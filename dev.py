# -*- coding: utf-8 -*-
"""成本结算平台 · 本地开发加速器

把以往每次都要手工做的事（找 PID → taskkill → 起服务 → 等 30-60 秒 →
探活 → py 脚本中文输出报错）压缩成一条命令。

用法（在项目根目录执行）：
    python dev.py up            启动本地服务（sqlite + 关闭模板缓存），就绪后才返回
    python dev.py restart       重启；改代码/改模板后一条命令搞定
    python dev.py restart -w    重启并预热常用页面（推荐）
    python dev.py down          只杀 runserver 进程（绝不误杀其它 python）
    python dev.py status        进程 / 端口 / HTTP / 启动参数 / 日志尾部 一览
    python dev.py log [n]       看服务端日志尾部（默认 40 行）
    python dev.py warm          预热常用页面
    python dev.py py x.py a b   用 venv 跑脚本（DB_ENGINE=sqlite，中文不再崩）
    python dev.py manage ...    透传 manage.py 命令，如 dev.py manage shell

设计要点（都是踩坑换来的）：
  * 进程枚举/杀进程一律走 psutil，不用 wmic / netstat / taskkill 子进程：
    本机 wmic 会挂死，且子进程继承 stdout 句柄会让外层命令"看起来卡住"。
  * 服务以 DETACHED_PROCESS 启动，日志重定向到文件，不随终端退出而死。
  * 带 --settings=cost_platform.dev_settings（显式非缓存模板加载器）：
    改 templates/*.html 后刷新浏览器即生效。
  * 中文输出用系统本地编码（GBK）并 errors=replace；不要加 -X utf8，
    实测那样输出会丢。
  * 日志：dev_server_stdout.log / dev_server_stderr.log
"""
from __future__ import annotations

import argparse
import locale
import os
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
from pathlib import Path

try:
    import psutil
except ImportError:  # 理论上不会发生，留个兜底
    psutil = None

BASE = Path(__file__).resolve().parent
VENV_PY = BASE / ".venv" / "Scripts" / "python.exe"
HOST = "127.0.0.1"
PORT = int(os.getenv("DEV_PORT", "8000"))
OUT_LOG = BASE / "dev_server_stdout.log"
ERR_LOG = BASE / "dev_server_stderr.log"
DEV_SETTINGS = "cost_platform.dev_settings"
WATCH_TASK = "CostPlatWatch"
WARM_URLS = ["/", "/accounts/login/", "/admin/login/"]

DETACHED_PROCESS = 0x00000008
CREATE_NEW_PROCESS_GROUP = 0x00000200
LOCALE = locale.getpreferredencoding(False) or "utf-8"

# 中文输出安全网：控制台是 GBK 时遇到编不出的字符用 ? 顶替，而不是抛
# UnicodeEncodeError 把整条命令弄崩。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")
    except Exception:
        pass


def base_url() -> str:
    return f"http://{HOST}:{PORT}"


def child_env() -> dict:
    """子进程环境：本地用 sqlite 库；中文按控制台编码输出且不抛异常。"""
    env = os.environ.copy()
    env["DB_ENGINE"] = env.get("DEV_DB_ENGINE", "sqlite")
    env.setdefault("DJANGO_DEBUG", "1")
    env["PYTHONIOENCODING"] = f"{LOCALE}:replace"
    return env


def listening(timeout: float = 1.0) -> bool:
    try:
        with socket.create_connection((HOST, PORT), timeout=timeout):
            return True
    except OSError:
        return False


def runserver_procs() -> list[tuple[int, str]]:
    """命令行含 'manage.py runserver' 的 python 进程 [(pid, cmdline)]。"""
    if psutil is None:
        return []
    found: list[tuple[int, str]] = []
    me = os.getpid()
    for proc in psutil.process_iter(["pid", "name", "cmdline"]):
        try:
            info = proc.info
            name = (info.get("name") or "").lower()
            if "python" not in name:
                continue
            cmd = " ".join(info.get("cmdline") or [])
            if "runserver" not in cmd or info["pid"] == me:
                continue
            found.append((info["pid"], cmd))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return found


def port_owner() -> str:
    """占用本端口的 PID（走 netstat，实测 0.02s）。

    注意：不要用 psutil.net_connections() —— Windows 上它会枚举全部连接，
    实测会卡住好几分钟（"命令看起来挂了"的元凶之一）。
    """
    try:
        out = subprocess.run(
            ["netstat", "-ano"], capture_output=True, text=True, timeout=20,
            encoding=LOCALE, errors="replace",
        ).stdout
        for line in out.splitlines():
            parts = line.split()
            if len(parts) >= 5 and parts[0].upper() == "TCP" and parts[3] == "LISTENING" \
                    and parts[1].endswith(f":{PORT}"):
                return parts[4]
    except Exception:
        pass
    return "?"


def http_status(url: str, timeout: float = 10.0):
    try:
        with urllib.request.urlopen(url, timeout=timeout) as resp:
            return resp.status
    except urllib.error.HTTPError as exc:
        return exc.code
    except Exception:
        return None


def tail(path: Path, n: int) -> None:
    if not path.exists():
        print(f"  (无) {path.name}")
        return
    try:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return
    print(f"  --- {path.name} 最后 {n} 行 ---")
    for line in lines[-n:]:
        print("   ", line)
    if not lines:
        print("    (空)")


def kill_procs(quiet: bool = False) -> int:
    procs = runserver_procs()
    if not procs and listening():
        owner = port_owner()
        if owner.isdigit():
            procs = [(int(owner), f"占用 {PORT} 端口（命令行未匹配 runserver）")]
    killed = 0
    for pid, cmd in procs:
        if not quiet:
            print(f"  kill {pid}  {cmd[:90]}")
        if psutil is not None:
            try:
                proc = psutil.Process(pid)
                proc.kill()
                proc.wait(timeout=10)
                killed += 1
                continue
            except Exception:
                pass
        subprocess.run(["taskkill", "/PID", str(pid), "/F"], capture_output=True)
        killed += 1
    deadline = time.time() + 15
    while time.time() < deadline and listening():
        time.sleep(0.5)
    return killed


def start_via_task() -> bool:
    """通过计划任务启动服务。

    关键：这样服务进程挂在「任务计划程序」下面，不会随着我这边执行命令的
    终端/任务结束被一起回收（否则会出现"起了又没了、守护脚本5分钟后再拉起"
    的反复横跳）。
    """
    try:
        r = subprocess.run(["schtasks", "/run", "/tn", WATCH_TASK],
                           capture_output=True, timeout=30,
                           encoding=LOCALE, errors="replace")
        return r.returncode == 0
    except Exception:
        return False


def wait_ready(timeout: int, label: str = "") -> bool:
    """轮询到 HTTP 有响应为止。"""
    t0 = time.time()
    print(f"[..] {label}启动中，最多等 {timeout}s", end="", flush=True)
    while time.time() - t0 < timeout:
        if http_status(base_url(), timeout=3):
            print(f"\n[ok] 就绪 {base_url()}  用时 {time.time() - t0:.1f}s")
            return True
        print(".", end="", flush=True)
        time.sleep(1.5)
    print()
    return False


def spawn_direct(timeout: int) -> int:
    """直接拉起 runserver（带 dev_settings）。计划任务方式不可用时的兜底。"""
    if not VENV_PY.exists():
        print(f"[x] 找不到虚拟环境解释器：{VENV_PY}")
        return 2
    flags = [f"{HOST}:{PORT}", "--noreload"]
    if (BASE / "cost_platform" / "dev_settings.py").exists():
        flags += ["--settings", DEV_SETTINGS]
    else:
        print("[!] 未找到 cost_platform/dev_settings.py，使用默认 settings")

    out = open(OUT_LOG, "ab")
    err = open(ERR_LOG, "ab")
    try:
        proc = subprocess.Popen(
            [str(VENV_PY), "manage.py", "runserver", *flags],
            cwd=str(BASE), env=child_env(), stdin=subprocess.DEVNULL,
            stdout=out, stderr=err,
            creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP,
        )
    finally:
        out.close()
        err.close()

    t0 = time.time()
    deadline = t0 + timeout
    while time.time() < deadline:
        if proc.poll() is not None:
            print(f"[x] 进程已退出（返回码 {proc.returncode}），看日志：")
            tail(ERR_LOG, 25)
            return 3
        if http_status(base_url(), timeout=3):
            print(f"[ok] 就绪 {base_url()}  用时 {time.time() - t0:.1f}s  PID {proc.pid}")
            return 0
        time.sleep(1.5)
    print("[!] 等待超时，服务可能仍在启动；用 `python dev.py status` 再看")
    return 4


def cmd_up(args) -> int:
    if listening() and not args.force:
        st = http_status(base_url(), timeout=5)
        if st:
            print(f"[=] 服务已在运行 {base_url()}  HTTP {st}  PID {port_owner()}（加 --force 可强制重启）")
            if args.warm:
                cmd_warm(args)
            return 0
        print("[!] 端口被占用但 HTTP 无响应，清理后重启")
    elif listening():
        print("[!] --force：先停掉现有服务")

    kill_procs()

    if not getattr(args, "no_task", False) and start_via_task():
        if wait_ready(args.timeout, "计划任务 CostPlatWatch "):
            if args.warm:
                cmd_warm(args)
            return 0
        print("[!] 计划任务方式没起来，改为直接启动")

    rc = spawn_direct(args.timeout)
    if rc == 0 and args.warm:
        cmd_warm(args)
    return rc


def cmd_autostart(args) -> int:
    """开关守护计划任务（CostPlatBoot / CostPlatWatch）。"""
    verb = {"on": "/enable", "off": "/disable", "status": None}[args.action]
    for name in ("CostPlatBoot", "CostPlatWatch"):
        try:
            if verb is None:
                r = subprocess.run(["schtasks", "/query", "/tn", name, "/fo", "LIST"],
                                   capture_output=True, timeout=30,
                                   encoding=LOCALE, errors="replace")
                state = "未知"
                for line in r.stdout.splitlines():
                    if "状态" in line or "Status" in line or "模式" in line:
                        state = line.split(":", 1)[-1].strip()
                print(f"  {name}: {state}")
            else:
                r = subprocess.run(["schtasks", "/change", "/tn", name, verb],
                                   capture_output=True, timeout=30,
                                   encoding=LOCALE, errors="replace")
                print(f"  {name}: {'成功' if r.returncode == 0 else '失败 ' + r.stderr.strip()[:80]}")
        except Exception as exc:
            print(f"  {name}: 探测失败 {type(exc).__name__}: {exc}")
    return 0


def cmd_down(args) -> int:
    n = kill_procs()
    print(f"[ok] 已停止 {n} 个 runserver 进程" if n else "[=] 没有 runserver 进程")
    print("[i] 提示：计划任务 CostPlatWatch 每 5 分钟会把服务拉回来（`dev.py autostart off` 可关）")
    if listening():
        print(f"[!] {PORT} 端口仍被 PID {port_owner()} 占用")
        return 1
    return 0


def cmd_restart(args) -> int:
    cmd_down(args)
    return cmd_up(args)


def listening_pids() -> set[int]:
    """真正监听本端口的 PID 集合。"""
    pids: set[int] = set()
    try:
        out = subprocess.run(
            ["netstat", "-ano"], capture_output=True, text=True, timeout=20,
            encoding=LOCALE, errors="replace",
        ).stdout
        for line in out.splitlines():
            parts = line.split()
            if len(parts) >= 5 and parts[0].upper() == "TCP" and parts[3] == "LISTENING" \
                    and parts[1].endswith(f":{PORT}") and parts[4].isdigit():
                pids.add(int(parts[4]))
    except Exception:
        pass
    return pids


def cmd_status(args) -> int:
    procs = runserver_procs()
    live = listening_pids()
    print(f"项目      : {BASE}")
    print(f"地址      : {base_url()}")
    print(f"端口监听  : {'是' if listening() else '否'}   占用 PID: {port_owner()}")
    st = http_status(base_url(), timeout=5) if listening() else None
    print(f"HTTP      : {st if st else '无响应'}")
    if procs:
        for pid, cmd in procs:
            tags = []
            tags.append("监听中" if pid in live else
                        ("venv 启动器父进程" if not live else "未监听"))
            tags.append("dev_settings" if DEV_SETTINGS in cmd else "默认 settings")
            print(f"进程 {pid:>6} : [{'] ['.join(tags)}]\n             {cmd}")
        if len(procs) > 1 and live:
            print("  (说明：venv 启动器 + 真实 runserver 属正常父子关系，不必紧张)")
    else:
        print("进程      : 无 runserver")
    print(f"venv      : {'有' if VENV_PY.exists() else '缺失'} {VENV_PY}")
    print(f"dev_settings: {'有' if (BASE / 'cost_platform' / 'dev_settings.py').exists() else '缺失'}")
    tail(ERR_LOG, 12)
    return 0


def cmd_log(args) -> int:
    tail(ERR_LOG, args.lines)
    tail(OUT_LOG, args.lines)
    return 0


def cmd_warm(args) -> int:
    if not listening():
        print("[x] 服务没在跑，先 `python dev.py up`")
        return 1
    print("[..] 预热常用页面（首次请求要建连接/编译模板，预热后人工点击才不卡）")
    for path in WARM_URLS:
        t0 = time.time()
        st = http_status(base_url() + path, timeout=90)
        print(f"  {path:<18} HTTP {st}  {time.time() - t0:.2f}s")
    return 0


def cmd_py(args) -> int:
    return subprocess.run(
        [str(VENV_PY), args.script, *args.rest], cwd=str(BASE), env=child_env()
    ).returncode


def cmd_manage(args) -> int:
    return subprocess.run(
        [str(VENV_PY), "manage.py", *args.rest], cwd=str(BASE), env=child_env()
    ).returncode


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="dev.py", description="成本结算平台 本地开发加速器",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：  python dev.py restart -w      python dev.py py _check_now.py",
    )
    sub = p.add_subparsers(dest="cmd")

    def add(name, fn, *aliases):
        sp = sub.add_parser(name, aliases=list(aliases))
        sp.set_defaults(fn=fn)
        return sp

    sp = add("up", cmd_up, "u")
    sp.add_argument("-f", "--force", action="store_true", help="已在运行也强制重启")
    sp.add_argument("-w", "--warm", action="store_true", help="就绪后预热页面")
    sp.add_argument("--no-task", action="store_true",
                    help="不用计划任务，直接拉起 runserver")
    sp.add_argument("-t", "--timeout", type=int, default=150, help="等待就绪秒数（默认 150）")

    add("down", cmd_down, "d", "stop")

    sp = add("restart", cmd_restart, "r")
    sp.add_argument("-f", "--force", action="store_true")
    sp.add_argument("-w", "--warm", action="store_true", default=True)
    sp.add_argument("--no-task", action="store_true",
                    help="不用计划任务，直接拉起 runserver")
    sp.add_argument("-t", "--timeout", type=int, default=150)

    add("status", cmd_status, "st")
    add("warm", cmd_warm, "w")

    sp = add("autostart", cmd_autostart)
    sp.add_argument("action", choices=["on", "off", "status"], nargs="?", default="status",
                    help="on/off 开关开机守护计划任务；默认看状态")

    sp = add("log", cmd_log, "logs")
    sp.add_argument("lines", nargs="?", type=int, default=40)

    sp = add("py", cmd_py, "run")
    sp.add_argument("script")
    sp.add_argument("rest", nargs=argparse.REMAINDER)

    sp = add("manage", cmd_manage, "m")
    sp.add_argument("rest", nargs=argparse.REMAINDER)

    return p


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    if not getattr(args, "fn", None):
        parser.print_help()
        return 0
    return args.fn(args)


if __name__ == "__main__":
    sys.exit(main())
