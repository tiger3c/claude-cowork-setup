"""Terminal output helpers: colors (Windows 10+ console included), Korean-width tables, step headers."""
import os
import re
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


# Korean Windows consoles (UTF-8 code page) keep one cell for "ambiguous" symbols but the font draws
# them two cells wide, so they overlap the next character. plain() replaces them before they reach
# the console; the width math below then only has to count Korean (wide) characters as two.
WIDE = "WFA" if os.environ.get("MENU_CRAWLER_WIDE_AMBIGUOUS") else "WF"
PLAIN = {"·": "/", "→": "->", "←": "<-", "‘": "'", "’": "'", "“": '"', "”": '"', "…": "...", "※": "*",
         "⚠": "!", "★": "*", "☆": "*", "■": "#", "□": "-", "▶": ">", "▷": ">", "└": "+", "├": "+", "─": "-",
         "━": "=", "│": "|", "×": "x", "°": "", "±": "+-", "✔": "v", "✖": "x", "•": "-", "º": "", "§": "S"}
_PLAIN_TABLE = str.maketrans({k: v for k, v in PLAIN.items()})


def plain(text):
    """Console-safe text: spaced middle dots become ' | ', other ambiguous symbols ASCII."""
    return str(text).replace(" · ", " | ").translate(_PLAIN_TABLE)


class ConsoleText:
    """stdout wrapper that writes plain() text; everything else is delegated."""

    def __init__(self, raw):
        self.raw = raw

    def write(self, text):
        return self.raw.write(plain(text))

    def __getattr__(self, name):
        return getattr(self.raw, name)


def install_console():
    if (os.name == "nt" or os.environ.get("MENU_CRAWLER_ASCII")) and not isinstance(sys.stdout, ConsoleText):
        sys.stdout = ConsoleText(sys.stdout)


def width(text):
    text = ANSI.sub("", plain(text) if os.name == "nt" else str(text))
    return sum(2 if unicodedata.east_asian_width(ch) in WIDE else 0 if unicodedata.combining(ch) else 1 for ch in str(text))


def fit(text, size):
    """Cut/pad to a display width so a line never wraps (Korean characters take two cells)."""
    out, used = "", 0
    for ch in str(text):
        w = width(ch)
        if used + w > size:
            while out and used + 2 > size:       # room for ".."
                used -= width(out[-1])
                out = out[:-1]
            out, used = out + "..", used + 2
            break
        out, used = out + ch, used + w
    return out + " " * max(0, size - used)


def screen():
    """Usable line width: one less than the console so the cursor never wraps."""
    return max(40, min(100, shutil.get_terminal_size((100, 20)).columns - 1))


def header(text, sub=""):
    w = screen()
    print()
    print(accent("=" * w))
    print(fit("  " + text, w).rstrip() if not sub else "  " + bold(text) + dim(fit("   " + sub, max(0, w - 2 - width(text))).rstrip()))
    print(accent("=" * w))


def step(no, total, text):
    print()
    print(accent(f"  [{no}/{total}] ") + bold(text))
    print(dim("  " + "-" * (screen() - 2)))


def rule():
    print(dim("  " + "-" * (screen() - 2)))


ANSI = re.compile(r"\033\[[0-9;]*m")


def say(text, indent="  "):
    """Print one line cut to the console width (a colour code around the text is kept whole)."""
    text = str(text)
    codes = ANSI.findall(text)
    body = ANSI.sub("", text)
    line = fit(indent + body, screen()).rstrip()
    if codes and _enable():
        line = line[:len(indent)] + codes[0] + line[len(indent):] + "\033[0m"
    print(line)


def table(columns, rows, indent="  "):
    """columns: [(title, width)]; rows: lists of cell text. A last width of 0 takes the rest of the line."""
    widths = [w for _, w in columns]
    rest = max(8, screen() - width(indent) - sum(widths[:-1]) - 2 * (len(widths) - 1))
    widths[-1] = min(widths[-1], rest) if widths[-1] else rest
    print(indent + dim("  ".join(fit(t, w) for (t, _), w in zip(columns, widths)).rstrip()))
    for row in rows:
        print(indent + "  ".join(fit(c, w) for c, w in zip(row, widths)).rstrip())


def grid(cells, indent="    ", min_width=26):
    """Short items in as many columns as fit the console."""
    size = max(min_width, max((width(c) for c in cells), default=0) + 2)
    per = max(1, (screen() - width(indent)) // size)
    for i in range(0, len(cells), per):
        print(indent + "".join(fit(c, size) for c in cells[i:i + per]).rstrip())


def ok(text):
    print("  " + good("[완료] ") + text)


def note(text):
    print("  " + warn("[참고] ") + text)


def fail(text):
    print("  " + bad("[확인] ") + text)


# ---------- 질문 ----------

def ask(prompt, default=""):
    suffix = f" [{default}]" if default else ""
    try:
        # input() writes its prompt straight to the console (not through sys.stdout) on Windows.
        text = f"{prompt}{suffix}: "
        value = input(plain(text) if isinstance(sys.stdout, ConsoleText) else text).strip()
    except EOFError:            # input closed (e.g. piped): stop instead of asking forever
        raise SystemExit("\n  입력이 끝나 종료합니다.") from None
    return value or default


YES = {"y", "yes", "ㅛ", "예", "네", "넹", "응", "ㅇ", "ㅇㅇ", "ok", "오케이", "좋아"}
NO = {"n", "no", "ㅜ", "아니", "아니요", "아니오", "ㄴ", "ㄴㄴ", "노"}


def ask_yes(prompt, default=False):
    """Y/N with Korean answers too ('네', or 'y' typed while the Korean keyboard is on = 'ㅛ'); asks again on anything else."""
    mark = "Y/n" if default else "y/N"
    while True:
        value = ask(f"{prompt} ({mark})").strip().lower()
        if not value:
            return default
        if value in YES:
            return True
        if value in NO:
            return False
        print("  y(예) 또는 n(아니요)로 답하세요.")


def ask_int(prompt, default, low, high):
    while True:
        value = ask(prompt, str(default))
        if value.isdigit() and low <= int(value) <= high:
            return int(value)
        print(f"  {low}~{high} 사이 숫자를 입력하세요.")


def beep():
    """Get attention when the user is needed (login) or the run is over; silent where not supported."""
    try:
        if os.name == "nt":
            import winsound
            winsound.MessageBeep()
        elif sys.stdout.isatty():
            sys.stdout.write("\a")
            sys.stdout.flush()
    except (ImportError, RuntimeError, OSError):
        pass


def window_title(text):
    """Progress in the console window title (visible in the taskbar while the user does other work)."""
    if os.name != "nt":
        return
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleTitleW(f"웹 메뉴 수집기 - {text}")
    except (AttributeError, OSError):
        pass
