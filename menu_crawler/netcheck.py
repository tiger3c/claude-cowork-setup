"""Quick reachability check before a long multi-target run (VPN / 사내망 off, typo in a domain)."""
import socket
import ssl
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import urlsplit


def check(url, timeout=6):
    """-> (ok, text). Any HTTP answer, even 401/403/500, means the server is reachable."""
    host = urlsplit(url).hostname or ""
    try:
        socket.getaddrinfo(host, None)
    except OSError:
        return False, "주소 이름 확인 실패"
    ctx = ssl.create_default_context()
    ctx.check_hostname, ctx.verify_mode = False, ssl.CERT_NONE        # same as the crawler: self-signed is fine
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 menu-crawler-check"})
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as res:
            return True, f"응답 {res.status}"
    except urllib.error.HTTPError as exc:
        return True, f"응답 {exc.code}"
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        reason = getattr(exc, "reason", exc)
        if isinstance(reason, (TimeoutError, socket.timeout)) or "timed out" in str(reason):
            return False, "응답 없음(시간 초과)"
        if isinstance(reason, ConnectionRefusedError):
            return False, "연결 거부"
        return False, f"접속 실패({type(reason).__name__})"


def check_all(urls, timeout=6):
    unique = list(dict.fromkeys(urls))
    with ThreadPoolExecutor(max_workers=min(8, max(1, len(unique)))) as pool:
        results = dict(zip(unique, pool.map(lambda u: check(u, timeout), unique)))
    return [results[u] for u in urls]
