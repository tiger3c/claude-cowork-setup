"""Terminal output helpers: colors (Windows 10+ console included), Korean-width tables, step headers."""
import os
import shutil
import sys
import unicodedata

_COLOR = None


def _enable():
    global _COLOR
    if _COLOR is not None:
        return _COLOR
    _COLOR = sys.stdout.isatty() and not os.environ.get("NO_COLOR")
    if _COLOR and os.name == "nt":
        try:
            import ctypes
            k = ctypes.windll.kernel32
            h = k.GetStdHandle(-11)
            mode = ctypes.c_uint32()
            if k.GetConsoleMode(h, ctypes.byref(mode)):
                k.SetConsoleMode(h, mode.value | 0x0004)   # ENABLE_VIRTUAL_TERMINAL_PROCESSING
            else:
                _COLOR = False
        except (AttributeError, OSError):
            _COLOR = False
    return _COLOR


def paint(text, code):
    return f"\033[{code}m{text}\033[0m" if _enable() else text


def bold(t):
    return paint(t, "1")


def dim(t):
    return paint(t, "2")


def good(t):
    return paint(t, "32")


def warn(t):
    return paint(t, "33")


def bad(t):
    return paint(t, "31")


def accent(t):
    return paint(t, "36")


def width(text):
    return sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in str(text))


def fit(text, size):
    """Cut/pad to a display width (Korean characters take two cells)."""
    text, out, used = str(text), "", 0
    for ch in text:
        w = width(ch)
        if used + w > size:
            out = out[:-1] + "…" if out and size > 1 else out
            used = width(out)
            break
        out, used = out + ch, used + w
    return out + " " * max(0, size - used)


def screen():
    return max(60, min(110, shutil.get_terminal_size((100, 20)).columns - 1))


def header(text, sub=""):
    w = screen()
    print()
    print(accent("━" * w))
    print("  " + bold(text) + (dim("   " + sub) if sub else ""))
    print(accent("━" * w))


def step(no, total, text):
    print()
    print(accent(f"  [{no}/{total}] ") + bold(text))


def rule():
    print(dim("  " + "─" * (screen() - 2)))


def table(columns, rows, indent="  "):
    """columns: [(title, width)]; rows: lists of cell text. The last column takes the rest of the line."""
    widths = [w for _, w in columns]
    rest = max(10, screen() - len(indent) - sum(widths[:-1]) - 2 * (len(widths) - 1))
    widths[-1] = min(widths[-1], rest) if widths[-1] else rest
    print(indent + dim("  ".join(fit(t, w) for (t, _), w in zip(columns, widths))))
    for row in rows:
        print(indent + "  ".join(fit(c, w) for c, w in zip(row, widths)).rstrip())


def ok(text):
    print("  " + good("✔ ") + text)


def note(text):
    print("  " + warn("※ ") + text)


def fail(text):
    print("  " + bad("✖ ") + text)


# ---------- 질문 ----------

def ask(prompt, default=""):
    suffix = f" [{default}]" if default else ""
    try:
        value = input(f"{prompt}{suffix}: ").strip()
    except EOFError:
        value = ""
    return value or default


def ask_yes(prompt, default=False):
    mark = "Y/n" if default else "y/N"
    value = ask(f"{prompt} ({mark})").lower()
    return default if not value else value in {"y", "yes", "예", "ㅇ"}


def ask_int(prompt, default, low, high):
    while True:
        value = ask(prompt, str(default))
        if value.isdigit() and low <= int(value) <= high:
            return int(value)
        print(f"  {low}~{high} 사이 숫자를 입력하세요.")
