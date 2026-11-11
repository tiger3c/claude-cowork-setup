import re
from urllib.parse import parse_qsl, urlencode, urljoin, urlsplit, urlunsplit

TRACKING = {"gclid", "fbclid", "_ga", "_gl"}
SECRETS = {"token", "access_token", "id_token", "refresh_token", "authorization", "password", "passwd", "sessionid", "jsessionid", "ticket", "code"}
FILES = re.compile(r"\.(?:pdf|zip|exe|msi|msix|dmg|docx?|xlsx?|pptx?|png|jpe?g|gif|svg|webp|mp[34]|woff2?|css|js)(?:$|[?#])", re.I)
ACTION = re.compile(r"(?:^|[/_.?&=\-])(?:logout|signout|logoff|delete|remove|withdraw|unsubscribe|destroy|terminate)(?:$|[/_.?&=\-])", re.I)
AUTH = re.compile(r"(?:^|[/_.?&=\-])(?:login|signin|sign-in|sso|authenticate)(?:$|[/_.?&=\-])", re.I)


def normalize_host(value: str) -> str:
    value = value.strip()
    parsed = urlsplit(value if "://" in value else "https://" + value)
    if not parsed.hostname or parsed.username or parsed.password or "*" in value:
        raise ValueError("허용 호스트는 와일드카드 없이 입력하세요. 예: www.example.com")
    host = parsed.hostname.encode("idna").decode("ascii").lower()
    if ":" in host:
        host = f"[{host}]"
    return host + (f":{parsed.port}" if parsed.port else "")


def canonical_url(value: str, base: str = "") -> str:
    try:
        parsed = urlsplit(urljoin(base, value))
        if parsed.scheme.lower() not in {"http", "https"} or not parsed.hostname or parsed.username or parsed.password:
            return ""
        host = parsed.hostname.encode("idna").decode("ascii").lower()
        if ":" in host:
            host = f"[{host}]"
        port = parsed.port
        if port and (parsed.scheme.lower(), port) not in {("https", 443), ("http", 80)}:
            host += f":{port}"
        query = [(k, v) for k, v in parse_qsl(parsed.query, keep_blank_values=True)
                 if not k.lower().startswith("utm_") and k.lower() not in TRACKING]
        fragment = parsed.fragment if parsed.fragment.startswith(("/", "!/")) else ""
        return urlunsplit((parsed.scheme.lower(), host, parsed.path or "/", urlencode(query), fragment))
    except (ValueError, UnicodeError):
        return ""


def public_url(value: str) -> str:
    """Redact common authentication query parameters from UI and exported evidence."""
    try:
        p = urlsplit(value)
        pairs = [(k, "[가림]" if k.lower() in SECRETS else v) for k, v in parse_qsl(p.query, keep_blank_values=True)]
        return urlunsplit((p.scheme, p.netloc, p.path, urlencode(pairs), p.fragment))
    except ValueError:
        return ""


def visit_policy(url: str, hosts: set[str]) -> tuple[str, str]:
    if not url:
        return "동작 확인 필요", "주소가 없는 동적 메뉴입니다. 직접 메뉴를 연 뒤 현재 화면 추가를 사용하세요."
    if urlsplit(url).netloc.lower() not in hosts:
        return "범위 밖", "허용 호스트에 포함되지 않아 방문하지 않았습니다."
    if FILES.search(url):
        return "파일 링크", "다운로드 파일입니다. 메뉴 목록에는 남기고 실행·다운로드하지 않습니다."
    if ACTION.search(url):
        return "동작 확인 필요", "로그아웃·삭제 등 상태 변경 가능성이 있는 주소는 자동 방문하지 않습니다."
    if AUTH.search(url):
        return "로그인 필요", "로그인 경로입니다. 인증 후 현재 화면 추가 또는 재수집을 사용하세요."
    return "대기", ""
