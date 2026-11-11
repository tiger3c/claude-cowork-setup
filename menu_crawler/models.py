from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from typing import Literal

from pydantic import BaseModel, Field, field_validator

from .urls import canonical_url, normalize_host


class CrawlConfig(BaseModel):
    url: str
    site_name: str = Field(default="메뉴 수집", max_length=120)
    extra_hosts: list[str] = Field(default_factory=list, max_length=30)
    max_pages: int = Field(default=300, ge=1, le=2000)
    max_seconds: int = Field(default=1200, ge=10, le=7200)
    max_hops: int = Field(default=6, ge=1, le=20)
    wait_ms: int = Field(default=700, ge=0, le=5000)
    delay_ms: int = Field(default=300, ge=0, le=5000)
    timeout_ms: int = Field(default=15000, ge=1000, le=60000)
    browser: Literal["auto", "msedge", "chrome", "chromium"] = "auto"
    menu_selector: str = Field(default="", max_length=500)
    headless: bool = False

    @field_validator("url")
    @classmethod
    def valid_url(cls, value: str) -> str:
        result = canonical_url(value.strip())
        if not result:
            raise ValueError("http:// 또는 https://로 시작하는 사이트 주소를 입력하세요.")
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

    @property
    def ready(self) -> bool:
        return (self.status == "수집됨" or (self.kind == "group" and len(self.path) == 3)) and not self.hierarchy_review and 1 <= len(self.path) <= 3

    def to_dict(self) -> dict:
        return {**asdict(self), "ready": self.ready}


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

    def to_dict(self) -> dict:
        return {**asdict(self), "menus": [m.to_dict() for m in self.menus]}
