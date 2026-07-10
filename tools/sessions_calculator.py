from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, timedelta

import requests
from bs4 import BeautifulSoup

DEFAULT_CALENDAR_URL = "https://www.educa.jcyl.es/es/informacion/calendario-escolar-2026-2027"

_MESES = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}

_SECUNDARIA_RE = re.compile(r"-\s*educaci[oó]n secundaria obligatoria\b", re.IGNORECASE)
_NAVIDAD_RE = re.compile(r"vacaciones de navidad", re.IGNORECASE)
_SEMANA_SANTA_RE = re.compile(r"vacaciones de semana santa", re.IGNORECASE)

_YEAR_IN_REF_RE = re.compile(r"_(\d{4})_")


@dataclass
class HolidayGroup:
    """Un festivo/día no lectivo con nombre propio (puede abarcar varios días)."""

    name: str
    category: str
    days: set[date] = field(default_factory=set)

    @property
    def start(self) -> date:
        return min(self.days)

    @property
    def end(self) -> date:
        return max(self.days)


@dataclass
class SchoolCalendar:
    course_start: date
    course_end: date
    christmas_days: set[date] = field(default_factory=set)
    easter_days: set[date] = field(default_factory=set)
    other_holidays: list[HolidayGroup] = field(default_factory=list)
    source_url: str = ""

    @property
    def holiday_days(self) -> set[date]:
        return self.christmas_days | self.easter_days


@dataclass
class _GridEvent:
    day: date
    event_type: str
    title: str


def _parse_grid_events(soup: BeautifulSoup) -> list[_GridEvent]:
    """Recorre la rejilla mensual del calendario (tablas <table class="month">).

    Cada día especial trae su clase CSS `eventType<Categoria>` y un atributo
    `title` con la descripción; el enlace interno codifica el año en su
    `href`/`name` (p.ej. "#11_2026_2_924756" -> año 2026).
    """
    events: list[_GridEvent] = []
    for month_table in soup.select("table.month"):
        caption = month_table.find("caption")
        if caption is None:
            continue
        month_num = _MESES.get(caption.get_text(strip=True).lower())
        if month_num is None:
            continue

        for td in month_table.select("td.calendarCustomFestiveCell"):
            classes = td.get("class") or []
            event_type = next((c[len("eventType"):] for c in classes if c.startswith("eventType")), None)
            if event_type is None:
                continue

            title = (td.get("title") or "").strip()
            anchor = td.find("a")
            day_text = anchor.get_text(strip=True) if anchor is not None else td.get_text(strip=True)
            try:
                day_num = int(day_text)
            except ValueError:
                continue

            year = None
            if anchor is not None:
                ref = anchor.get("href") or anchor.get("name") or ""
                match = _YEAR_IN_REF_RE.search(ref)
                if match:
                    year = int(match.group(1))
            if year is None:
                continue

            try:
                day = date(year, month_num, day_num)
            except ValueError:
                continue

            events.append(_GridEvent(day=day, event_type=event_type, title=title))

    return events


def _parse_calendar_html(html: str, source_url: str) -> SchoolCalendar:
    soup = BeautifulSoup(html, "html.parser")
    events = _parse_grid_events(soup)
    if not events:
        raise ValueError("No se han encontrado eventos en el calendario. Puede que haya cambiado su estructura.")

    course_start = None
    course_end = None
    christmas_days: set[date] = set()
    easter_days: set[date] = set()
    holiday_groups: dict[str, HolidayGroup] = {}

    for ev in events:
        event_type = ev.event_type.lower()
        if event_type.startswith("iniciodecurso"):
            if course_start is None and _SECUNDARIA_RE.search(ev.title):
                course_start = ev.day
        elif event_type.startswith("findecurso"):
            if course_end is None and _SECUNDARIA_RE.search(ev.title):
                course_end = ev.day
        elif event_type == "vacacionesescolares":
            if _NAVIDAD_RE.search(ev.title):
                christmas_days.add(ev.day)
            elif _SEMANA_SANTA_RE.search(ev.title):
                easter_days.add(ev.day)
        elif event_type in ("fiestaslaborales", "diasnolectivos"):
            group = holiday_groups.setdefault(ev.title, HolidayGroup(name=ev.title, category=ev.event_type))
            group.days.add(ev.day)

    if course_start is None or course_end is None:
        raise ValueError(
            "No se han podido localizar las fechas de inicio/fin de curso de Educación Secundaria "
            "Obligatoria en la página. Puede que haya cambiado su estructura."
        )

    other_holidays = sorted(holiday_groups.values(), key=lambda g: min(g.days))

    return SchoolCalendar(
        course_start=course_start,
        course_end=course_end,
        christmas_days=christmas_days,
        easter_days=easter_days,
        other_holidays=other_holidays,
        source_url=source_url,
    )


def fetch_school_calendar(url: str = DEFAULT_CALENDAR_URL) -> SchoolCalendar:
    response = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
    response.raise_for_status()
    return _parse_calendar_html(response.text, url)


def free_weeks_in_range(start: date, end: date, holiday_days: set[date]) -> float:
    if end < start:
        return 0.0
    total_days = (end - start).days + 1
    excluded = sum(1 for d in holiday_days if start <= d <= end)
    net_days = max(total_days - excluded, 0)
    return net_days / 7


def approx_sessions(weeks: float, weekly_hours: float, safety_factor: float) -> float:
    return max(weeks, 0) * max(weekly_hours, 0) * (1 - safety_factor)


def default_trimester_ranges(calendar: SchoolCalendar) -> tuple[tuple[date, date], tuple[date, date], tuple[date, date]]:
    """Propone una división en 3 trimestres en torno a Navidad y Semana Santa."""
    christmas_start = min(calendar.christmas_days) if calendar.christmas_days else None
    christmas_end = max(calendar.christmas_days) if calendar.christmas_days else None
    easter_start = min(calendar.easter_days) if calendar.easter_days else None
    easter_end = max(calendar.easter_days) if calendar.easter_days else None

    t1_end = (christmas_start - timedelta(days=1)) if christmas_start else calendar.course_start
    t2_start = (christmas_end + timedelta(days=1)) if christmas_end else calendar.course_start
    t2_end = (easter_start - timedelta(days=1)) if easter_start else calendar.course_end
    t3_start = (easter_end + timedelta(days=1)) if easter_end else calendar.course_end

    return (
        (calendar.course_start, t1_end),
        (t2_start, t2_end),
        (t3_start, calendar.course_end),
    )


def trimester_index_for_date(day: date, trimester_ranges: list[tuple[date, date]]) -> int | None:
    for i, (start, end) in enumerate(trimester_ranges):
        if start <= day <= end:
            return i
    return None
