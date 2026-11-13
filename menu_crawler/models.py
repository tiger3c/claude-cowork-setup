from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from typing import Literal

from pydantic import BaseModel, Field, field_validator

from .urls import canonical_url, normalize_host


MODES = {
    "public": "로그인 전 메뉴 수집",
    "login": "로그인 후 메뉴 수집",
    "compare": "로그인 전·후 비교 수집",
}
PHASES = {"before": "로그인 전", "after": "로그인 후"}


class CrawlConfig(BaseModel):
    url: str
    mode: Literal["public", "login", "compare"] = "public"
    login_url: str = Field(default="", max_length=2000)
    site_name: str = Field(default="메뉴 수집", max_length=120)
    extra_hosts: list[str] = Field(default_factory=list, max_length=30)
    max_pages: int = Field(default=0, ge=0, le=100000)      # 0 = 제한 없음
    max_seconds: int = Field(default=0, ge=0, le=172800)    # 0 = 제한 없음
    max_hops: int = Field(default=6, ge=1, le=20)
    wait_ms: int = Field(default=700, ge=0, le=5000)
    delay_ms: int = Field(default=300, ge=0, le=5000)
    timeout_ms: int = Field(default=15000, ge=1000, le=60000)
    browser: Literal["auto", "msedge", "chrome", "chromium"] = "auto"
    menu_selector: str = Field(default="", max_length=500)
    path_prefix: str = Field(default="", max_length=500)   # only pages under this path (several services on one host)
    click_menus: bool = True       # open script (javascript:) menus by clicking them, never logout/delete-like names
    headless: bool = False

    @field_validator("url")
    @classmethod
    def valid_url(cls, value: str) -> str:
        result = canonical_url(value.strip())
        if not result:
            raise ValueError("http:// 또는 https://로 시작하는 사이트 주소를 입력하세요.")
        return result

    @field_validator("login_url")
    @classmethod
    def valid_login_url(cls, value: str) -> str:
        value = value.strip()
        if not value:
            return ""
        result = canonical_url(value)
        if not result:
            raise ValueError("로그인 페이지 주소는 http:// 또는 https://로 시작해야 합니다.")
        return result

    @field_validator("extra_hosts")
    @classmethod
    def valid_hosts(cls, value: list[str]) -> list[str]:
        return list(dict.fromkeys(normalize_host(v) for v in value if v.strip()))


@dataclass
class Menu:
    id: str
    path: list[str]
    url: str
    source_url: str
    evidence: str
    kind: str = "link"
    status: str = "대기"
    reason: str = ""
    title: str = ""
    http_status: int | None = None
    final_url: str = ""
    hierarchy_review: bool = False
    seen: list[str] = field(default_factory=list)
    status_before: str = ""

    @property
    def visibility(self) -> str:
        if not self.seen:
            return ""
        if "before" in self.seen:
            return "로그인 전 공개"
        return "로그인 후 전용"

    @property
    def ready(self) -> bool:
        return (self.status == "수집됨" or (self.kind == "group" and len(self.path) == 3)) and not self.hierarchy_review and 1 <= len(self.path) <= 3

    def to_dict(self) -> dict:
        return {**asdict(self), "ready": self.ready, "visibility": self.visibility}


@dataclass
class Report:
    config: dict
    state: str = "starting"
    message: str = "브라우저를 여는 중입니다."
    started_at: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    finished_at: str = ""
    menus: list[Menu] = field(default_factory=list)
    pages: list[dict] = field(default_factory=list)
    issues: list[dict] = field(default_factory=list)
    visited_count: int = 0
    current_url: str = ""
    limit_reached: bool = False
    phase: str = ""
    phases_done: list[str] = field(default_factory=list)
    login_check: str = ""

    def to_dict(self) -> dict:
        return {**asdict(self), "menus": [m.to_dict() for m in self.menus]}
