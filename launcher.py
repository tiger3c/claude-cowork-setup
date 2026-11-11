import argparse
import secrets
import socket
import threading
import webbrowser

import uvicorn

from menu_crawler.web import create_app


def main():
    parser = argparse.ArgumentParser(description="웹 메뉴 수집기 · 로컬 실행")
    parser.add_argument("--port", type=int, default=0)
    parser.add_argument("--no-open", action="store_true")
    args = parser.parse_args()
    token = secrets.token_urlsafe(32)
    # Bind before opening the UI so another process cannot take the chosen port.
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.bind(("127.0.0.1", args.port))
    sock.listen(128)
    url = f"http://127.0.0.1:{sock.getsockname()[1]}/#token={token}"
    print("웹 메뉴 수집기를 시작합니다. 종료하려면 이 창에서 Ctrl+C를 누르세요.", flush=True)
    print("실행 화면: " + url, flush=True)
    if not args.no_open:
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    server = uvicorn.Server(uvicorn.Config(create_app(token), log_level="warning", access_log=False))
    try:
        server.run(sockets=[sock])
    finally:
        sock.close()


if __name__ == "__main__":
    main()
