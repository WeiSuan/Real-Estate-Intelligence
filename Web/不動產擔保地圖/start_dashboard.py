import http.server
import os
import socket
import subprocess
import sys
import time
import webbrowser
from pathlib import Path


HOST = "0.0.0.0"
LOCAL_BROWSER_HOST = "127.0.0.1"
PORT = 8772
HTML_FILE = "index - V1.html"


def wait_before_exit() -> None:
    try:
        input("\n按 Enter 關閉視窗...")
    except EOFError:
        pass


def run_command(command: list[str], *, cwd: Path | None = None) -> int:
    print(f"> {' '.join(command)}")
    completed = subprocess.run(command, cwd=str(cwd) if cwd else None)
    return completed.returncode


def find_listening_pids(port: int) -> list[int]:
    completed = subprocess.run(
        ["netstat", "-ano"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if completed.returncode != 0:
        raise RuntimeError("netstat 執行失敗，無法檢查舊服務。")

    pids: set[int] = set()
    for line in completed.stdout.splitlines():
        parts = line.split()
        if len(parts) < 5:
            continue

        protocol, local_address, _foreign_address, state, pid_text = parts[:5]
        if protocol.upper() != "TCP":
            continue
        if state.upper() != "LISTENING":
            continue
        if not local_address.endswith(f":{port}"):
            continue

        try:
            pids.add(int(pid_text))
        except ValueError:
            continue

    return sorted(pids)


def stop_existing_servers(port: int) -> None:
    pids = find_listening_pids(port)
    if not pids:
        print(f"未發現舊服務占用 {HOST}:{port}")
        return

    for pid in pids:
        print(f"關閉舊服務 PID {pid}")
        result = subprocess.run(["taskkill", "/PID", str(pid), "/F"])
        if result.returncode != 0:
            raise RuntimeError(f"無法關閉舊服務 PID {pid}")

    time.sleep(1)


def refresh_sample(project_dir: Path) -> None:
    script_path = project_dir / "tools" / "refresh-sample.ps1"
    if not script_path.exists():
        raise FileNotFoundError(f"找不到資料刷新腳本：{script_path}")

    result = run_command(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_path),
        ],
        cwd=project_dir,
    )
    if result != 0:
        raise RuntimeError("資料刷新失敗，請確認 data\\總表_YYYYMM.xlsx 是否存在。")


def assert_port_available(host: str, port: int) -> None:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as probe:
        probe.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            probe.bind((host, port))
        except OSError as exc:
            raise RuntimeError(f"{host}:{port} 仍被占用，無法啟動服務。") from exc


def get_lan_hosts() -> list[str]:
    hosts: set[str] = set()
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            address = info[4][0]
            if not address.startswith("127."):
                hosts.add(address)
    except OSError:
        pass

    hosts.add(socket.gethostname())
    return sorted(hosts)


def start_server(project_dir: Path) -> None:
    os.chdir(project_dir)
    cache_buster = int(time.time())
    path = HTML_FILE.replace(" ", "%20")
    local_url = f"http://{LOCAL_BROWSER_HOST}:{PORT}/{path}?v={cache_buster}"
    lan_urls = [f"http://{host}:{PORT}/{path}?v={cache_buster}" for host in get_lan_hosts()]

    class DashboardHandler(http.server.SimpleHTTPRequestHandler):
        def end_headers(self) -> None:
            # Always fetch local dashboard assets again after the refresh step.
            self.send_header("Cache-Control", "no-store, max-age=0")
            super().end_headers()

    httpd = http.server.ThreadingHTTPServer((HOST, PORT), DashboardHandler)
    try:
        print()
        print(f"已啟動擔保地圖：{local_url}")
        print("同公司內網同仁可嘗試使用以下網址連線：")
        for lan_url in lan_urls:
            print(f"  {lan_url}")
        print("若同仁無法連線，請確認此電腦與同仁在同一內網，且 Windows 防火牆允許 Python 存取私人網路。")
        print("此視窗就是擔保地圖伺服器，請保持開啟。")
        print("要停止服務時，請在此視窗按 Ctrl+C，或直接關閉此視窗。")
        print()
        webbrowser.open(local_url)
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n收到停止指令。")
    finally:
        httpd.server_close()
        print("擔保地圖伺服器已停止。")


def main() -> int:
    project_dir = Path(__file__).resolve().parent
    try:
        stop_existing_servers(PORT)
        refresh_sample(project_dir)
        assert_port_available(HOST, PORT)
        start_server(project_dir)
        return 0
    except Exception as exc:
        print()
        print(f"啟動失敗：{exc}")
        wait_before_exit()
        return 1


if __name__ == "__main__":
    sys.exit(main())
