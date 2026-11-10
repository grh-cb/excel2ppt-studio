"""便携版入口：双击即启动本地服务并打开浏览器。

- 自动挑一个空闲端口（8765 起）；
- 只监听 127.0.0.1，数据不出机；
- 控制台窗口保留日志，关闭窗口即停止服务。
"""

from __future__ import annotations

import socket
import sys


def find_free_port(start: int = 8765, limit: int = 50) -> int:
    for port in range(start, start + limit):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            try:
                sock.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    return start


def main() -> int:
    from excel2ppt_studio.web.app import run

    port = find_free_port()
    print("=" * 56)
    print("  Excel2PPT Studio 便携版")
    print(f"  地址：http://127.0.0.1:{port}")
    print("  浏览器会自动打开；关闭本窗口即停止服务。")
    print("=" * 56)
    run(port=port, open_browser=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
