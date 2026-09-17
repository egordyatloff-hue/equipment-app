# -*- coding: utf-8 -*-
"""Учёт поверки оборудования.

- список приборов: название, заводской номер, дата поверки, исполнитель;
- сортировка списка;
- отдельная кнопка/экран «Поверка в этом месяце»;
- индикатор справа: зелёный — срок не подошёл, жёлтый — сдача в этом
  месяце, красный — просрочено.
"""

import calendar
import json
import os
from datetime import date, datetime, timedelta

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from kivy.app import App
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.graphics import Color, Rectangle, RoundedRectangle
from kivy.metrics import dp
from kivy.properties import BooleanProperty, ListProperty, ObjectProperty, StringProperty
from kivy.uix.behaviors import ButtonBehavior
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.screenmanager import Screen, ScreenManager
from kivy.uix.scrollview import ScrollView
from kivy.uix.textinput import TextInput
from kivy.utils import get_color_from_hex, platform

from sync import (SyncClient, now_iso, device_id, SERVER_URL, API_TOKEN,
                  check_update, download_update, install_update, APP_VERSION)

APP_TITLE = "Учёт поверки"
VERSION = "1.0"

# Адрес сервера синхронизации и токен (вшиты в приложение)
SERVER_URL = "https://rezzonvoice.ru/api"
API_TOKEN = "hBqxlkwcoWrA65RuUaHstETn7ipOZ2801YIQD3GgSbeNmXJy"

C = {
    "bg": get_color_from_hex("#F2F4F8"),
    "primary": get_color_from_hex("#1565C0"),
    "primary_dark": get_color_from_hex("#0D47A1"),
    "accent": get_color_from_hex("#00897B"),
    "danger": get_color_from_hex("#C62828"),
    "warning": get_color_from_hex("#F9A825"),
    "ok": get_color_from_hex("#2E7D32"),
    "text": get_color_from_hex("#212121"),
    "text_light": get_color_from_hex("#757575"),
    "white": (1, 1, 1, 1),
    "chip_green": get_color_from_hex("#C8E6C9"),
    "chip_yellow": get_color_from_hex("#FFE082"),
    "chip_red": get_color_from_hex("#FFCDD2"),
    "chip_none": get_color_from_hex("#CFD8DC"),
    "line": get_color_from_hex("#B0BEC5"),
}

STATUS_INFO = {
    "green": ("ОК", C["ok"], C["chip_green"]),
    "yellow": ("На поверке", C["text"], C["chip_yellow"]),
    "red": ("Просрочено", C["white"], C["danger"]),
    "none": ("Нет даты", C["text_light"], C["chip_none"]),
}

STATES = [
    ("installed", "Установлен"),
    ("removed", "Снят"),
    ("verification", "На поверке"),
    ("reserve", "В резерве"),
    ("taken_from_verification", "Забрали с поверки"),
]

STATE_LABELS = dict(STATES)


# ------------------------- данные -------------------------

def crash_log_path():
    """Путь для crash.log: открытая пользователю папка на Android."""
    if platform == "android":
        for cand in ("/sdcard/Documents", "/sdcard/Download"):
            try:
                if os.path.isdir(cand):
                    return os.path.join(cand, "equipment_crash.log")
            except Exception:
                continue
        return "/sdcard/equipment_crash.log"
    return os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "equipment_crash.log")


def data_dir():
    if platform == "android":
        try:
            from android.storage import app_storage_dir
            return app_storage_dir()
        except Exception:
            return "/sdcard/Documents"
    return os.path.dirname(os.path.abspath(__file__))


def data_path():
    return os.path.join(data_dir(), "equipment.json")


def load_records():
    try:
        with open(data_path(), "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def save_records(records):
    try:
        with open(data_path(), "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=1)
    except Exception as e:
        print("save error:", e)


def export_excel(records, path):
    """Создать Excel-таблицу со всеми записями.

    Столбцы: Дата действия, Название, Заводской номер, Объект, Место,
    Исполнитель, Дата поверки, Состояние, Статус.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Приборы"

    headers = ["Дата действия", "Название прибора", "Заводской номер",
               "Дата поверки", "Объект", "Место", "Исполнитель",
               "Состояние", "Статус"]
    head_fill = PatternFill("solid", fgColor="1565C0")
    head_font = Font(bold=True, color="FFFFFF")
    thin_align = Alignment(horizontal="center", vertical="center")

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = head_fill
        cell.font = head_font
        cell.alignment = thin_align

    status_names = {"green": "ОК", "yellow": "В этом месяце",
                    "red": "Просрочено", "none": "Нет даты"}
    status_fills = {
        "green": PatternFill("solid", fgColor="C8E6C9"),
        "yellow": PatternFill("solid", fgColor="FFE082"),
        "red": PatternFill("solid", fgColor="FFCDD2"),
        "none": PatternFill("solid", fgColor="CFD8DC"),
    }

    recs = sorted((r for r in records if not r.get("deleted")),
                  key=lambda r: (r.get("added") or "", -r.get("id", 0)),
                  reverse=True)
    for i, r in enumerate(recs, 2):
        d = parse_date(r.get("verification_date", ""))
        st = effective_status(r, records)
        ws.cell(row=i, column=1, value=r.get("action_date") or "")
        ws.cell(row=i, column=2, value=r.get("name") or "")
        ws.cell(row=i, column=3, value=r.get("serial") or "")
        ws.cell(row=i, column=4, value=fmt_date(d))
        ws.cell(row=i, column=5, value=r.get("object") or "")
        ws.cell(row=i, column=6, value=r.get("location") or "")
        ws.cell(row=i, column=7, value=r.get("executor") or "")
        ws.cell(row=i, column=8,
                value=STATE_LABELS.get(r.get("state") or "installed", ""))
        c = ws.cell(row=i, column=9, value=status_names.get(st, ""))
        c.fill = status_fills.get(st, PatternFill())

    widths = [14, 30, 18, 14, 20, 20, 20, 16, 16]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"
    # Автофильтр: стрелки фильтрации/поиска в шапке каждого столбца
    ws.auto_filter.ref = "A1:I%d" % max(1, len(recs) + 1)
    wb.save(path)
    return len(recs)


def export_excel_unique(records, path):
    """Excel только с уникальными приборами: название, зав. номер, срок
    поверки. Уникальность по з/н, без з/н — по названию. Из всех записей
    прибора берётся самая новая (для актуального срока поверки)."""
    groups = {}
    for r in records:
        if r.get("deleted"):
            continue
        serial = (r.get("serial") or "").strip()
        name = (r.get("name") or "").strip()
        key = ("s:" + serial.lower()) if serial else (
            ("n:" + name.lower()) if name else None)
        if key is None:
            continue
        cur = groups.get(key)
        if cur is None or (r.get("added") or "", r.get("id", 0)) > \
                (cur.get("added") or "", cur.get("id", 0)):
            groups[key] = r

    recs = sorted(groups.values(),
                  key=lambda r: (r.get("added") or "", -r.get("id", 0)),
                  reverse=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Приборы"
    headers = ["Название прибора", "Заводской номер", "Дата поверки"]
    head_fill = PatternFill("solid", fgColor="1565C0")
    head_font = Font(bold=True, color="FFFFFF")
    thin_align = Alignment(horizontal="center", vertical="center")
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = head_fill
        cell.font = head_font
        cell.alignment = thin_align
    for i, r in enumerate(recs, 2):
        d = parse_date(r.get("verification_date", ""))
        ws.cell(row=i, column=1, value=r.get("name") or "")
        ws.cell(row=i, column=2, value=r.get("serial") or "")
        ws.cell(row=i, column=3, value=fmt_date(d))
    for col, w in enumerate([32, 20, 16], 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:C%d" % max(1, len(recs) + 1)
    wb.save(path)
    return len(recs)


# ------------------------- даты -------------------------

def parse_date(s):
    if not s:
        return None
    s = str(s).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def fmt_date(d):
    return d.strftime("%d.%m.%Y") if d else "—"


def months_until(d):
    today = date.today()
    return (d.year - today.year) * 12 + (d.month - today.month)


def status_of(d):
    if d is None:
        return "none"
    m = months_until(d)
    if m < 0:
        return "red"
    if m == 0:
        return "yellow"
    return "green"


def days_left(d):
    if d is None:
        return None
    return (d - date.today()).days


def effective_status(rec, records):
    """Статус записи с учётом других записей того же заводского номера.

    Правила:
    - состояние «На поверке» -> жёлтый, пока это последнее событие по
      прибору (даже если срок просрочен); если потом прибор вернули
      (установлен/снят) — запись превращается в зелёную (поверка
      завершена), а просрочки старых записей по-прежнему закрываются
      по наличию новой даты поверки;
    - просроченная запись становится зелёной, если по тому же заводскому
      номеру есть более поздняя запись «На поверке» или запись с более
      новой датой поверки; но если после «На поверке» появился
      «Установлен»/«Снят» БЕЗ новой поверки — просрочки снова красные.
    """
    d = parse_date(rec.get("verification_date", ""))
    st = status_of(d)
    state = rec.get("state") or "installed"
    serial = (rec.get("serial") or "").strip()

    # Запись «Забрали с поверки» = прибор вернули из поверки: всегда ОК
    if state == "taken_from_verification":
        return "green"

    same = [o for o in records
            if o is not rec and not o.get("deleted")
            and (o.get("serial") or "").strip() == serial]
    my_key = (rec.get("added") or "", rec.get("id", 0))

    # Все записи «На поверке» того же з/н, созданные ПОЗЖЕ текущей
    vers_after = [o for o in same
                  if (o.get("state") or "installed") == "verification"
                  and (o.get("added") or "", o.get("id", 0)) > my_key]
    # Самая поздняя запись-событие того же з/н
    latest_any_key = max(((o.get("added") or "", o.get("id", 0))
                          for o in same), default=None)
    # «На поверке» актуальна, если она сама является последним событием
    my_is_ver = state == "verification"
    latest_ver_key = max(((o.get("added") or "", o.get("id", 0))
                          for o in same
                          if (o.get("state") or "installed") == "verification"),
                         default=None)
    if my_is_ver:
        ver_actual = latest_any_key is None or my_key >= latest_any_key
    else:
        ver_actual = (latest_ver_key is not None
                      and not any((o.get("added") or "", o.get("id", 0)) > latest_ver_key
                                  for o in same))

    if my_is_ver:
        # «На поверке» — последняя по времени запись з/н, но если после
        # неё появились другие события — поверка завершена.
        if ver_actual:
            return "yellow"
        # Завершилась ли поверка УСПЕШНО (есть запись с новой датой поверки)?
        for o in same:
            od = parse_date(o.get("verification_date", ""))
            if od and d and od > d:
                return "green"
        # Поверка завершилась «Забрали с поверки» и прибор остался со
        # старым сроком — считаем зелёным (прибор цел, ждёт установки)
        later_taken = [o for o in same
                       if (o.get("state") or "installed") == "taken_from_verification"
                       and (o.get("added") or "", o.get("id", 0)) > my_key]
        if later_taken:
            return "green"
        return st
    if st != "red":
        return st
    if not serial:
        return st
    for o in same:
        od = parse_date(o.get("verification_date", ""))
        if od and d and od > d:
            return "green"
    if vers_after and ver_actual:
        return "green"
    # «Забрали с поверки» по этому прибору закрывает просрочку
    later_taken_any = [o for o in same
                       if (o.get("state") or "installed") == "taken_from_verification"
                       and (o.get("added") or "", o.get("id", 0)) > my_key]
    if later_taken_any:
        return "green"
    return st


def latest_verification_for(rec, records):
    """Самая новая дата поверки среди записей того же заводского номера.

    Старые записи прибора показывают актуальный срок поверки (он
    «переезжает» из новой записи), а не свой исторический.
    """
    serial = (rec.get("serial") or "").strip()
    if not serial:
        return parse_date(rec.get("verification_date", ""))
    best = None
    best_key = None
    for o in records:
        if o.get("deleted") or (o.get("serial") or "").strip() != serial:
            continue
        od = parse_date(o.get("verification_date", ""))
        if od and (best is None or od > best):
            best = od
        key = (o.get("added") or "", o.get("id", 0))
        if best_key is None or key > best_key:
            best_key = key
    return best


def _digits_filter(value, negative):
    allowed = "0123456789."
    return "".join(ch for ch in value if ch in allowed)[:10]


# ------------------------- UI-виджеты -------------------------

class AutoLabel(Label):
    """Label с подгонкой текста под размер, выравнивание сверху/слева."""

    def __init__(self, **kw):
        kw.setdefault("valign", "top")
        super().__init__(**kw)
        self.bind(size=self._sync)

    def _sync(self, w, size):
        self.text_size = size


class CalendarPicker(Popup):
    """Простой календарь для выбора даты."""

    MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
              "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
    DOW = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]

    def __init__(self, field, **kw):
        self.field = field
        super().__init__(title="Выбор даты", size_hint=(0.9, 0.85), **kw)
        self.year = date.today().year
        self.month = date.today().month
        box = BoxLayout(orientation="vertical", spacing=dp(6), padding=[dp(8)])
        nav = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(6))
        nav.add_widget(Button(text="<", bold=True, font_size="24sp",
                              background_normal="",
                              background_color=C["text_light"],
                              on_release=lambda *a: self._shift(-1)))
        self.month_label = AutoLabel(text="", bold=True, font_size="18sp",
                                     color=C["primary_dark"], halign="center",
                                     valign="middle")
        nav.add_widget(self.month_label)
        nav.add_widget(Button(text=">", bold=True, font_size="24sp",
                              background_normal="",
                              background_color=C["text_light"],
                              on_release=lambda *a: self._shift(1)))
        box.add_widget(nav)
        self.grid = BoxLayout(orientation="vertical", spacing=dp(2))
        box.add_widget(self.grid)
        box.add_widget(Button(text="Закрыть", size_hint_y=None, height=dp(48),
                              on_release=lambda *a: self.dismiss()))
        self.add_widget(box)
        self._rebuild()

    def _shift(self, delta):
        m = self.month + delta
        if m < 1:
            m, self.year = 12, self.year - 1
        elif m > 12:
            m, self.year = 1, self.year + 1
        self.month = m
        self._rebuild()

    def _rebuild(self):
        self.month_label.text = "%s %d" % (self.MONTHS[self.month - 1], self.year)
        self.grid.clear_widgets()
        dow = BoxLayout(spacing=dp(2), size_hint_y=None, height=dp(34))
        for d in self.DOW:
            dow.add_widget(AutoLabel(text=d, bold=True, font_size="14sp",
                                     color=C["text_light"], halign="center",
                                     valign="middle"))
        self.grid.add_widget(dow)
        cal = calendar.Calendar(firstweekday=0)
        for week in cal.monthdatescalendar(self.year, self.month):
            row = BoxLayout(spacing=dp(2), size_hint_y=None, height=dp(44))
            for day in week:
                is_other = day.month != self.month
                btn = Button(text=str(day.day), font_size="16sp",
                             background_normal="",
                             background_color=C["line"] if is_other else C["primary"],
                             color=C["text_light"] if is_other else C["white"],
                             on_release=lambda *a, dd=day: self._pick(dd))
                row.add_widget(btn)
            self.grid.add_widget(row)

    def _pick(self, day):
        self.field.text = day.strftime("%d.%m.%Y")
        self.dismiss()


class MonthPicker(Popup):
    """Выбор месяца и года."""

    MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
              "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

    def __init__(self, app, **kw):
        self.app = app
        today = date.today()
        self.year = today.year
        super().__init__(title="Выбор месяца", size_hint=(0.9, 0.85), **kw)
        box = BoxLayout(orientation="vertical", spacing=dp(8), padding=[dp(8)])
        nav = BoxLayout(size_hint_y=None, height=dp(52), spacing=dp(6))
        nav.add_widget(Button(text="<", bold=True, font_size="26sp",
                              background_normal="",
                              background_color=C["text_light"],
                              on_release=lambda *a: self._shift(-1)))
        self.year_label = AutoLabel(text=str(self.year), bold=True,
                                    font_size="22sp",
                                    color=C["primary_dark"],
                                    halign="center", valign="middle")
        nav.add_widget(self.year_label)
        nav.add_widget(Button(text=">", bold=True, font_size="26sp",
                              background_normal="",
                              background_color=C["text_light"],
                              on_release=lambda *a: self._shift(1)))
        box.add_widget(nav)

        grid = BoxLayout(orientation="vertical", spacing=dp(6))
        for row in range(4):
            r = BoxLayout(spacing=dp(6))
            for col in range(3):
                m = row * 3 + col + 1
                r.add_widget(Button(text=self.MONTHS[m - 1], bold=True,
                                    font_size="15sp", background_normal="",
                                    background_color=C["primary"],
                                    on_release=lambda *a, mm=m: self._pick(mm)))
            grid.add_widget(r)
        box.add_widget(grid)
        box.add_widget(Button(text="Закрыть", size_hint_y=None, height=dp(48),
                              on_release=lambda *a: self.dismiss()))
        self.add_widget(box)

    def _shift(self, delta):
        self.year += delta
        self.year_label.text = str(self.year)

    def _pick(self, month):
        self.dismiss()
        self.app.set_due_month(self.year, month)


class DateField(BoxLayout):
    """Строка ввода даты с иконкой календаря справа."""

    def __init__(self, hint, **kw):
        super().__init__(orientation="horizontal", spacing=dp(6),
                         size_hint_y=None, height=dp(56), **kw)
        self.input = TextInput(hint_text=hint, multiline=False,
                               font_size="18sp", input_filter=_digits_filter)
        self.add_widget(self.input)
        cal_btn = Button(text="Выбрать", font_size="14sp", size_hint_x=None,
                         width=dp(92), bold=True, background_normal="",
                         background_color=C["primary_dark"],
                         on_release=lambda *a: CalendarPicker(self.input).open())
        self.add_widget(cal_btn)


class ColorBar(BoxLayout):
    """Узкая цветная полоска слева строки."""

    def __init__(self, color, **kw):
        super().__init__(size_hint_x=None, width=dp(6), **kw)
        self.bar_color = color
        self.bind(pos=self._upd, size=self._upd)

    def _upd(self, *a):
        self.canvas.before.clear()
        with self.canvas.before:
            Color(*self.bar_color)
            RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(3)])


class StatusChip(AutoLabel):
    """Цветной ярлык статуса в правой части строки."""

    def __init__(self, bg, **kw):
        kw.setdefault("bold", True)
        kw.setdefault("font_size", "13sp")
        kw.setdefault("halign", "center")
        kw.setdefault("valign", "middle")
        super().__init__(**kw)
        self.chip_bg = bg
        self.bind(pos=self._upd, size=self._upd)

    def _upd(self, *a):
        self.canvas.before.clear()
        with self.canvas.before:
            Color(*self.chip_bg)
            RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(10)])


class RecordCard(ButtonBehavior, BoxLayout):
    """Строка прибора. Нажатие — редактирование."""

    def __init__(self, rec, app, **kw):
        super().__init__(orientation="horizontal", size_hint_y=None,
                         height=dp(86), spacing=dp(2),
                         padding=[dp(8), dp(4), dp(6), dp(4)], **kw)
        self.rec = rec
        self.app = app

        d = parse_date(rec.get("verification_date", ""))
        st = effective_status(rec, app.records)
        st_text, _fg, chip_bg = STATUS_INFO[st]

        # Состояние «В резерве» показываем своим ярлыком.
        # «Забрали с поверки» — своим ярлыком только пока запись актуальна;
        # если по прибору появилась более поздняя запись — обычный статус.
        rec_state = rec.get("state") or "installed"
        my_key = (rec.get("added") or "", rec.get("id", 0))
        serial = (rec.get("serial") or "").strip()
        name = (rec.get("name") or "").strip()
        has_later = False
        for o in app.records:
            if o is rec or o.get("deleted"):
                continue
            o_serial = (o.get("serial") or "").strip()
            o_name = (o.get("name") or "").strip()
            same = (serial and o_serial == serial) or \
                   (not serial and name and o_name == name)
            if same and (o.get("added") or "", o.get("id", 0)) > my_key:
                has_later = True
                break
        if rec_state == "reserve":
            st_text, chip_bg = "В резерве", C["line"]
        elif rec_state == "taken_from_verification" and not has_later:
            st_text, chip_bg = "Забрали", C["line"]

        with self.canvas.before:
            Color(0.16, 0.18, 0.22, 0.06)
            bg = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(8)])
        self.bind(pos=lambda w, p: setattr(bg, "pos", p),
                  size=lambda w, s: setattr(bg, "size", s))

        self.add_widget(ColorBar(_status_color(st)))

        info = BoxLayout(orientation="vertical", spacing=dp(1))
        info.add_widget(AutoLabel(text=rec.get("name") or "Без названия",
                                  bold=True, font_size="17sp",
                                  color=C["text"], shorten=True,
                                  shorten_from="right", size_hint_y=None,
                                  height=dp(21)))
        sub = ("Зав. номер: %s  •  %s" %
               (rec.get("serial") or "—",
                " / ".join(x for x in (rec.get("object"), rec.get("location")) if x)
                or "место не указано"))
        info.add_widget(AutoLabel(text=sub, font_size="13sp",
                                  color=C["text"], bold=True,
                                  shorten=True,
                                  shorten_from="right", size_hint_y=None,
                                  height=dp(16)))
        left = days_left(d)
        if st == "yellow" and status_of(d) == "yellow":
            extra = "Дата поверки: %s  •  осталось %s дн." % (fmt_date(d), left)
        elif st == "red":
            extra = "Дата поверки: %s  •  просрочено на %s дн." % (fmt_date(d), abs(left))
        else:
            extra = "Дата поверки: %s" % fmt_date(d)
        info.add_widget(AutoLabel(text=extra, font_size="13sp",
                                  color=C["text"], size_hint_y=None,
                                  height=dp(16)))

        st_label = STATE_LABELS.get(rec.get("state") or "installed", "Установлен")
        act_date = rec.get("action_date") or ""
        state_line = "Состояние: %s  •  от %s" % (st_label, act_date or "—")
        if getattr(self, "reason_text", None):
            state_line = "Причина: %s  •  действие: %s" % (self.reason_text, act_date or "—")
        info.add_widget(AutoLabel(text=state_line, font_size="13sp",
                                  color=C["text_light"], size_hint_y=None,
                                  height=dp(16)))
        self.add_widget(info)

        chip = StatusChip(chip_bg, text=st_text, size_hint_x=None,
                          width=dp(96))
        self.add_widget(chip)

        self.bind(on_release=lambda *a: app.open_edit(rec))


def _status_color(st):
    return {"green": C["ok"], "yellow": C["warning"],
            "red": C["danger"], "none": C["chip_none"]}[st]


def _reason_labels(reasons):
    labels = []
    for r in reasons:
        labels.append({"action": "действие",
                       "due": "к поверке",
                       "overdue": "просрочено"}.get(r, r))
    return labels


class ListScreen(Screen):
    pass


class EditScreen(Screen):
    pass


class DueScreen(Screen):
    pass


class VerificationScreen(Screen):
    pass


class Root(BoxLayout):
    pass


# ------------------------- приложение -------------------------

class EquipmentApp(App):
    title = APP_TITLE
    records = ListProperty()
    editing_rec = ObjectProperty({}, allow_none=True)
    search_text = StringProperty("")

    def build(self):
        if platform not in ("android", "ios"):
            try:
                Window.size = (440, 800)
            except Exception:
                pass
        try:
            Window.softinput_mode = "below_target"
        except Exception:
            pass
        self.records = load_records()
        self.sync_client = SyncClient(self, SERVER_URL, API_TOKEN)

        root = Root()
        self.sm = ScreenManager(size_hint=(1, 1))
        root.add_widget(self.sm)
        self.sm.add_widget(ListScreen(name="list"))
        self.sm.add_widget(EditScreen(name="edit"))
        self.sm.add_widget(DueScreen(name="due"))
        self.sm.add_widget(VerificationScreen(name="verification"))
        Clock.schedule_once(lambda dt: self.build_static(), 0)
        return root

    def save_and_refresh(self):
        save_records(self.records)
        self.refresh_all()

    # ---------- построение UI ----------
    def build_static(self):
        try:
            self._build_list_screen(self.sm.get_screen("list"))
            self._build_edit_screen(self.sm.get_screen("edit"))
            self._build_due_screen(self.sm.get_screen("due"))
            self._build_verification_screen(self.sm.get_screen("verification"))
            self.show_list()
            Clock.schedule_interval(self.periodic_sync, 300)
        except Exception:
            import traceback
            from kivy.uix.label import Label as _L
            from kivy.uix.popup import Popup as _P
            _P(title="Ошибка старта", content=_L(
                text=traceback.format_exc()[-1500:],
                font_size="10sp"),
                size_hint=(0.95, 0.9)).open()
        # Обновления: проверить в фоне через 10 секунд после старта
        Clock.schedule_once(lambda dt: self._startup_sync(), 10)

    def _startup_sync(self):
        from threading import Thread
        Thread(target=lambda: self.sync_client.sync_now(),
               daemon=True).start()

    def periodic_sync(self, dt):
        self.sync_client.sync_now()

    def _top_bar(self):
        bar = BoxLayout(size_hint_y=None, height=dp(64), spacing=dp(6),
                        padding=[dp(8), dp(8), dp(8), 0])

        def tb(text, color, cb, hx):
            b = Button(text=text, bold=True, background_normal="",
                       background_color=color, font_size="16sp",
                       on_release=cb)
            b.size_hint_x = hx
            bar.add_widget(b)
            return b

        tb("+ Добавить", C["primary"], lambda *a: self.open_add(), 1.0)
        self.due_btn = tb("! Месяц", C["warning"],
                          lambda *a: self.open_due(), 0.85)
        tb("На поверке", C["accent"], lambda *a: self.open_verification(), 0.95)
        tb("Excel", C["primary_dark"], lambda *a: self.export_to_excel(), 0.55)
        tb("Обновл.", C["text_light"], lambda *a: self.check_updates(), 0.62)
        return bar

    def _bg_panel(self, box):
        with box.canvas.before:
            Color(*C["bg"])
            rect = Rectangle(pos=box.pos, size=box.size)
        box.bind(pos=lambda w, p: setattr(rect, "pos", p),
                 size=lambda w, s: setattr(rect, "size", s))

    def _build_list_screen(self, scr):
        scr.clear_widgets()
        box = BoxLayout(orientation="vertical")
        self._bg_panel(box)
        box.add_widget(self._top_bar())

        wrap = BoxLayout(size_hint_y=None, height=dp(52),
                         padding=[dp(10), dp(6), dp(10), 0])
        self.search_input = TextInput(hint_text="Поиск: название / з/н / объект / место",
                                      multiline=False, font_size="16sp")
        self.search_input.bind(text=lambda inst, v: self.set_search(v))
        wrap.add_widget(self.search_input)
        box.add_widget(wrap)

        cwrap = BoxLayout(size_hint_y=None, height=dp(24),
                          padding=[dp(14), 0, dp(14), 0])
        self.count_label = AutoLabel(font_size="14sp", color=C["text_light"],
                                     halign="left")
        cwrap.add_widget(self.count_label)
        box.add_widget(cwrap)

        self.list_scroll = ScrollView(bar_width=dp(4))
        self.list_container = BoxLayout(orientation="vertical",
                                        size_hint_y=None, spacing=dp(6),
                                        padding=[dp(8), dp(2), dp(8), dp(16)])
        self.list_container.bind(minimum_height=lambda w, h: setattr(w, "height", h))
        self.list_scroll.add_widget(self.list_container)
        box.add_widget(self.list_scroll)
        scr.add_widget(box)

    def _build_edit_screen(self, scr):
        scr.clear_widgets()
        box = BoxLayout(orientation="vertical", spacing=dp(8),
                        padding=[dp(14), dp(14), dp(14), dp(14)])
        self._bg_panel(box)

        titlewrap = BoxLayout(size_hint_y=None, height=dp(32),
                              padding=[dp(4), 0, 0, 0])
        self.ed_title = AutoLabel(text="Новая запись", bold=True,
                                  font_size="24sp", color=C["primary_dark"],
                                  halign="left")
        titlewrap.add_widget(self.ed_title)
        box.add_widget(titlewrap)

        self.action_field = DateField("сегодня, изменить по календарю")
        box.add_widget(self.action_field)
        self.in_action_date = self.action_field.input

        for attr, hint in (
                ("in_name", "Название прибора"),
                ("in_serial", "Заводской номер")):
            w = TextInput(hint_text=hint, multiline=False, font_size="18sp")
            w.size_hint_y = None
            w.height = dp(56)
            setattr(self, attr, w)
            box.add_widget(w)

        self.in_date = TextInput(hint_text="Дата поверки",
                                 multiline=False, font_size="18sp",
                                 input_filter=_digits_filter)
        self.in_date.size_hint_y = None
        self.in_date.height = dp(56)
        box.add_widget(self.in_date)

        for attr, hint in (
                ("in_object", "Объект"),
                ("in_location", "Место"),
                ("in_executor", "Исполнитель")):
            w = TextInput(hint_text=hint, multiline=False, font_size="18sp")
            w.size_hint_y = None
            w.height = dp(56)
            setattr(self, attr, w)
            box.add_widget(w)

        self.state_label_btn = Button(text="Установлен", bold=True,
                                      font_size="17sp", background_normal="",
                                      background_color=C["ok"])
        self.state_label_btn.size_hint_y = None
        self.state_label_btn.height = dp(56)
        self.state_label_btn.bind(on_release=lambda *a: self.open_state_menu())
        box.add_widget(self.state_label_btn)

        box.add_widget(BoxLayout())  # распорка

        brow = BoxLayout(size_hint_y=None, height=dp(64), spacing=dp(8))
        brow.add_widget(Button(text="Изменить", bold=True, font_size="17sp",
                               background_normal="",
                               background_color=C["ok"],
                               on_release=lambda *a: self.save_record(False)))
        brow.add_widget(Button(text="Создать", bold=True, font_size="17sp",
                               background_normal="",
                               background_color=C["primary"],
                               on_release=lambda *a: self.save_record(True)))
        brow.add_widget(Button(text="Отмена", font_size="17sp",
                               background_normal="",
                               background_color=C["text_light"],
                               on_release=lambda *a: self.show_list()))
        box.add_widget(brow)

        self.del_btn = Button(text="Удалить запись", bold=True, font_size="16sp",
                              background_normal="",
                              background_color=C["danger"],
                              size_hint_y=None, height=dp(56),
                              on_release=lambda *a: self.delete_record())
        box.add_widget(self.del_btn)
        scr.add_widget(box)

    def _build_due_screen(self, scr):
        scr.clear_widgets()
        box = BoxLayout(orientation="vertical")
        self._bg_panel(box)

        top = BoxLayout(size_hint_y=None, height=dp(64),
                        padding=[dp(8), dp(8), dp(8), 0])
        top.add_widget(Button(text="< Назад к списку", bold=True, font_size="16sp",
                              background_normal="",
                              background_color=C["text_light"],
                              on_release=lambda *a: self.show_list()))
        box.add_widget(top)
        duewrap = BoxLayout(size_hint_y=None, height=dp(28),
                            padding=[dp(14), 0, dp(14), 0])
        self.due_label = AutoLabel(text="", bold=True, font_size="19sp",
                                   color=C["primary_dark"], halign="left")
        duewrap.add_widget(self.due_label)
        box.add_widget(duewrap)

        self.due_scroll = ScrollView(bar_width=dp(4))
        self.due_container = BoxLayout(orientation="vertical",
                                       size_hint_y=None, spacing=dp(6),
                                       padding=[dp(8), dp(4), dp(8), dp(16)])
        self.due_container.bind(minimum_height=lambda w, h: setattr(w, "height", h))
        self.due_scroll.add_widget(self.due_container)
        box.add_widget(self.due_scroll)
        scr.add_widget(box)

    def _build_verification_screen(self, scr):
        scr.clear_widgets()
        box = BoxLayout(orientation="vertical")
        self._bg_panel(box)

        top = BoxLayout(size_hint_y=None, height=dp(64),
                        padding=[dp(8), dp(8), dp(8), 0])
        top.add_widget(Button(text="< Назад к списку", bold=True, font_size="16sp",
                              background_normal="",
                              background_color=C["text_light"],
                              on_release=lambda *a: self.show_list()))
        box.add_widget(top)

        verwrap = BoxLayout(size_hint_y=None, height=dp(28),
                            padding=[dp(14), 0, dp(14), 0])
        self.verif_label = AutoLabel(text="", bold=True, font_size="19sp",
                                     color=C["primary_dark"], halign="left")
        verwrap.add_widget(self.verif_label)
        box.add_widget(verwrap)

        self.verif_scroll = ScrollView(bar_width=dp(4))
        self.verif_container = BoxLayout(orientation="vertical",
                                         size_hint_y=None, spacing=dp(6),
                                         padding=[dp(8), dp(4), dp(8), dp(16)])
        self.verif_container.bind(minimum_height=lambda w, h: setattr(w, "height", h))
        self.verif_scroll.add_widget(self.verif_container)
        box.add_widget(self.verif_scroll)
        scr.add_widget(box)

    # ---------- вспомогательное ----------
    def open_state_menu(self):
        """Выбор состояния из списка (вместо переключения по кругу)."""
        box = BoxLayout(orientation="vertical", spacing=dp(6),
                        padding=[dp(10)])
        pop = Popup(title="Состояние прибора", content=box,
                    size_hint=(0.8, 0.55))
        colors = {
            "installed": C["ok"],
            "removed": C["text_light"],
            "verification": C["warning"],
            "reserve": C["primary"],
            "taken_from_verification": C["primary_dark"],
        }
        for key, label in STATES:
            b = Button(text=label, bold=True, font_size="16sp",
                       background_normal="",
                       background_color=colors[key],
                       size_hint_y=None, height=dp(52))
            b.bind(on_release=lambda *a, k=key, p=pop: (self._set_state(k),
                                                        p.dismiss()))
            box.add_widget(b)
        pop.open()

    def _set_state(self, key):
        self._edit_state = key
        self._update_state_btn()

    def _update_state_btn(self):
        cur = getattr(self, "_edit_state", "installed")
        label, color = {
            "installed": ("Установлен", C["ok"]),
            "removed": ("Снят", C["text_light"]),
            "verification": ("На поверке", C["warning"]),
            "reserve": ("В резерве", C["primary"]),
            "taken_from_verification": ("Забрали с поверки", C["primary_dark"]),
        }[cur]
        self.state_label_btn.text = label
        self.state_label_btn.background_color = color

    def set_search(self, text):
        self.search_text = text
        self.refresh_list()

    def export_to_excel(self):
        """Выбор варианта экспорта: все данные или только приборы."""
        box = BoxLayout(orientation="vertical", spacing=dp(10),
                        padding=[dp(12)])
        box.add_widget(Label(text="Что выгрузить в Excel?"))
        b_all = Button(text="Все данные", bold=True, font_size="16sp",
                       background_normal="", background_color=C["primary"],
                       size_hint_y=None, height=dp(56))
        b_dev = Button(text="Только приборы", bold=True, font_size="16sp",
                       background_normal="", background_color=C["ok"],
                       size_hint_y=None, height=dp(56))
        b_cancel = Button(text="Отмена", size_hint_y=None, height=dp(48))
        box.add_widget(b_all)
        box.add_widget(b_dev)
        box.add_widget(b_cancel)
        pop = Popup(title="Экспорт в Excel", content=box,
                    size_hint=(0.8, 0.45))
        b_cancel.bind(on_release=lambda *a: pop.dismiss())
        b_all.bind(on_release=lambda *a: (pop.dismiss(),
                                          self._do_export("all")))
        b_dev.bind(on_release=lambda *a: (pop.dismiss(),
                                          self._do_export("devices")))
        pop.open()

    def _do_export(self, mode):
        try:
            stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
            if mode == "devices":
                path = os.path.join(data_dir(),
                                    "Приборы_список_%s.xlsx" % stamp)
                n = export_excel_unique(self.records, path)
            else:
                path = os.path.join(data_dir(),
                                    "Приборы_%s.xlsx" % stamp)
                n = export_excel(self.records, path)
            self._popup("Экспорт завершён",
                        "Файл: %s\nЗаписей: %d" % (os.path.basename(path), n))
        except Exception as e:
            self._popup("Ошибка экспорта", str(e))

    # ---------- обновление приложения ----------
    def check_updates(self):
        self._popup("Обновление", "Проверяю обновления...")
        from threading import Thread

        def work():
            ver, apk = check_update()
            Clock.schedule_once(lambda dt: self._update_result(ver, apk), 0)

        Thread(target=work, daemon=True).start()

    def _update_result(self, ver, apk):
        if ver is None:
            self._popup("Обновление", "У вас последняя версия (%s)." % APP_VERSION)
            return
        box = BoxLayout(orientation="vertical", spacing=dp(10), padding=[dp(10)])
        box.add_widget(Label(text="Доступна версия %s (у вас %s).\nСкачать и установить?"
                             % (ver, APP_VERSION)))
        brow = BoxLayout(size_hint_y=None, height=dp(48), spacing=dp(10))
        pop = Popup(title="Обновление", content=box, size_hint=(0.85, 0.4))

        def do_update(*a):
            pop.dismiss()
            self._popup("Обновление", "Скачиваю...")
            from threading import Thread

            def dl():
                try:
                    dest = os.path.join(data_dir(), "update_%s.apk" % ver)
                    download_update(apk, dest)
                    Clock.schedule_once(
                        lambda dt: self._popup(
                            "Обновление",
                            "Скачано. Открываю установку...\n"
                            "Разрешите установку, если спросит Android."),
                        0)
                    install_update(dest)
                except Exception as e:
                    Clock.schedule_once(
                        lambda dt: self._popup("Ошибка обновления", str(e)), 0)

            Thread(target=dl, daemon=True).start()

        brow.add_widget(Button(text="Обновить", bold=True, font_size="16sp",
                               background_normal="",
                               background_color=C["ok"],
                               on_release=do_update))
        brow.add_widget(Button(text="Позже",
                               on_release=lambda *a: pop.dismiss()))
        box.add_widget(brow)
        pop.open()

    def visible_records(self):
        recs = [r for r in self.records if not r.get("deleted")]
        if self.search_text:
            q = self.search_text.lower()
            recs = [r for r in recs if
                    q in (r.get("name") or "").lower()
                    or q in (r.get("serial") or "").lower()
                    or q in (r.get("object") or "").lower()
                    or q in (r.get("location") or "").lower()
                    or q in (r.get("executor") or "").lower()]

        def key(r):
            # Новые записи сверху: сортировка по времени создания (убывание)
            return ((r.get("added") or ""), -r.get("id", 0))

        return sorted(recs, key=key, reverse=True)

    # ---------- обновление экранов ----------
    def refresh_all(self):
        n_due = 0
        n_red = 0
        for r in self.records:
            if r.get("deleted"):
                continue
            st = effective_status(r, self.records)
            if st == "yellow":
                n_due += 1
            elif st == "red":
                n_red += 1
        if hasattr(self, "due_btn"):
            self.due_btn.text = ("! Месяц (%d)" % n_due) if n_due else "! Месяц"
            self.due_btn.background_color = C["warning"] if (n_due or n_red) else C["text_light"]
        self.refresh_list()

    def _unique_devices(self, records):
        """Уникальные приборы: по заводскому номеру; без з/н — по названию."""
        keys = set()
        for r in records:
            if r.get("deleted"):
                continue
            serial = (r.get("serial") or "").strip()
            name = (r.get("name") or "").strip()
            if serial:
                keys.add("s:" + serial.lower())
            elif name:
                keys.add("n:" + name.lower())
        return len(keys)

    def refresh_list(self):
        if not hasattr(self, "list_container"):
            return
        self.list_container.clear_widgets()
        recs = self.visible_records()
        if self.search_text:
            self.count_label.text = ("Найдено записей: %d  •  приборов: %d"
                                     % (len(recs), self._unique_devices(recs)))
        else:
            self.count_label.text = "Приборов: %d" % self._unique_devices(self.records)
        if not recs:
            self.list_container.add_widget(AutoLabel(
                text="Нет записей.\nНажмите «+ Добавить».",
                halign="center", color=C["text_light"], font_size="15sp"))
        for r in recs:
            self.list_container.add_widget(RecordCard(r, self))

    def refresh_due(self):
        if not hasattr(self, "due_container"):
            return
        self.due_container.clear_widgets()
        y = getattr(self, "due_year", date.today().year)
        m = getattr(self, "due_month", date.today().month)
        today = date.today()
        # Прошлый или текущий месяц -> история (действия + просрочки),
        # будущий месяц -> только приборы со сдачей на поверку в этом месяце.
        is_past_or_current = (y, m) <= (today.year, today.month)
        first_day = date(y, m, 1)
        # последний день месяца
        last_day = (date(y + m // 12, m % 12 + 1, 1) - timedelta(days=1))

        items = []
        for r in self.records:
            if r.get("deleted"):
                continue
            vd = parse_date(r.get("verification_date", ""))
            ad = parse_date(r.get("action_date", ""))
            st = effective_status(r, self.records)
            reasons = []
            if is_past_or_current:
                # Действия, совершённые в выбранном месяце
                if ad and first_day <= ad <= last_day:
                    reasons.append("action")
                # Прибор просрочен (и не закрыт) — надо было сдать
                if st == "red":
                    reasons.append("overdue")
            # Сдача на поверку в выбранном месяце: дата поверки попадает
            # в месяц И это ещё не просрочено на момент этого месяца
            if (vd and first_day <= vd <= last_day
                    and (st in ("yellow", "green") or not is_past_or_current)):
                reasons.append("due")
            if reasons:
                items.append((r, vd, ad, reasons))

        def key(t):
            # Новые сверху, как в основном списке
            return (t[0].get("added") or "", -t[0].get("id", 0))
        items.sort(key=key, reverse=True)
        month_name = "%s %d" % (MonthPicker.MONTHS[m - 1], y)
        n_act = sum(1 for i in items if "action" in i[3])
        n_due = sum(1 for i in items if "due" in i[3])
        n_over = sum(1 for i in items if "overdue" in i[3])
        self.due_label.text = ("%s: %d  •  к поверке: %d  •  просрочено: %d"
                               % (month_name, n_act, n_due, n_over))
        if not items:
            self.due_container.add_widget(AutoLabel(
                text="В этом месяце нет приборов\nс действиями или сдачей на поверку.",
                halign="center", color=C["text_light"], font_size="15sp"))
        for r, vd, ad, reasons in items:
            card = RecordCard(r, self)
            tag = ", ".join(_reason_labels(reasons))
            if tag:
                card.reason_text = tag
            self.due_container.add_widget(card)

    # ---------- навигация ----------
    def open_add(self):
        self.editing_rec = {}
        self.in_name.text = ""
        self.in_serial.text = ""
        self.in_object.text = ""
        self.in_location.text = ""
        self.in_executor.text = ""
        self.in_date.text = ""
        self.in_action_date.text = date.today().strftime("%d.%m.%Y")
        self._edit_state = "installed"
        self._update_state_btn()
        self.ed_title.text = "Новая запись"
        self.del_btn.disabled = True
        self.del_btn.opacity = 0.4
        self.sm.current = "edit"

    def open_edit(self, rec):
        self.editing_rec = rec
        self.in_name.text = rec.get("name", "")
        self.in_serial.text = rec.get("serial", "")
        self.in_object.text = rec.get("object", "")
        self.in_location.text = rec.get("location", "")
        self.in_executor.text = rec.get("executor", "")
        self.in_date.text = rec.get("verification_date", "")
        self.in_action_date.text = rec.get("action_date", "")
        self._edit_state = rec.get("state") or "installed"
        self._update_state_btn()
        self.ed_title.text = "Редактирование прибора"
        self.del_btn.disabled = False
        self.del_btn.opacity = 1
        self.sm.current = "edit"

    def show_list(self):
        self.sm.current = "list"
        self.refresh_all()

    def open_due(self):
        MonthPicker(self).open()

    def set_due_month(self, year, month):
        self.due_year = year
        self.due_month = month
        self.refresh_due()
        self.sm.current = "due"

    def open_verification(self):
        self.refresh_verification()
        self.sm.current = "verification"

    def refresh_verification(self):
        if not hasattr(self, "verif_container"):
            return
        self.verif_container.clear_widgets()
        items = []
        for r in self.records:
            if r.get("deleted"):
                continue
            if (r.get("state") or "installed") != "verification":
                continue
            # Только АКТУАЛЬНО на поверке: нет более поздней записи
            # (установлен/снят/в резерве) по тому же зав. номеру/названию
            serial = (r.get("serial") or "").strip()
            name = (r.get("name") or "").strip()
            my_key = (r.get("added") or "", r.get("id", 0))
            superseded = False
            for o in self.records:
                if o is r or o.get("deleted"):
                    continue
                o_serial = (o.get("serial") or "").strip()
                o_name = (o.get("name") or "").strip()
                same = (serial and o_serial == serial) or \
                       (not serial and name and o_name == name)
                if not same:
                    continue
                if (o.get("added") or "", o.get("id", 0)) > my_key:
                    superseded = True
                    break
            if not superseded:
                items.append(r)
        items.sort(key=lambda r: (r.get("added") or "", -r.get("id", 0)),
                   reverse=True)
        self.verif_label.text = "На поверке: %d" % len(items)
        if not items:
            self.verif_container.add_widget(AutoLabel(
                text="Нет приборов со состоянием\n«На поверке».",
                halign="center", color=C["text_light"], font_size="15sp"))
        for r in items:
            card = RecordCard(r, self)
            card.reason_text = "на поверке"
            self.verif_container.add_widget(card)

    # ---------- CRUD ----------
    def save_record(self, as_new=False):
        editing = bool(self.editing_rec and self.editing_rec.get("id")) and not as_new
        name = self.in_name.text.strip()
        if not name:
            self._popup("Ошибка", "Укажите название прибора.")
            return
        ds = self.in_date.text.strip()
        if ds and parse_date(ds) is None:
            self._popup("Ошибка", "Неверный формат даты поверки.\nИспользуйте ДД.ММ.ГГГГ")
            return
        ads = self.in_action_date.text.strip()
        if ads and parse_date(ads) is None:
            self._popup("Ошибка", "Неверный формат даты действия.\nИспользуйте ДД.ММ.ГГГГ")
            return
        state = getattr(self, "_edit_state", "installed")
        src = self.editing_rec if (self.editing_rec and self.editing_rec.get("id")) else {}
        data = {
            "id": src["id"] if editing else self._next_id(),
            "name": name,
            "serial": self.in_serial.text.strip(),
            "object": self.in_object.text.strip(),
            "location": self.in_location.text.strip(),
            "executor": self.in_executor.text.strip(),
            "verification_date": ds,
            "state": state,
            "action_date": ads,
            "added": (src.get("added") if editing
                      else datetime.now().isoformat(timespec="seconds")),
            "updated_at": now_iso(),
            "deleted": False,
        }
        if editing:
            for i, r in enumerate(self.records):
                if r.get("id") == data["id"]:
                    self.records[i] = data
                    break
        else:
            self.records.append(data)
        save_records(self.records)
        self.show_list()

    def delete_record(self):
        if not (self.editing_rec and self.editing_rec.get("id")):
            return
        rec = self.editing_rec
        box = BoxLayout(orientation="vertical", spacing=dp(10),
                        padding=[dp(10)])
        box.add_widget(Label(text="Удалить «%s»?" % rec.get("name", "")))
        brow = BoxLayout(size_hint_y=None, height=dp(44), spacing=dp(10))
        pop = Popup(title="Подтверждение", content=box, size_hint=(0.85, 0.4))

        def do_del(*a):
            for i, r in enumerate(self.records):
                if r.get("id") == rec.get("id"):
                    r["deleted"] = True
                    r["updated_at"] = now_iso()
                    break
            save_records(self.records)
            pop.dismiss()
            self.show_list()

        brow.add_widget(Button(text="Удалить", background_normal="",
                               background_color=C["danger"],
                               on_release=do_del))
        brow.add_widget(Button(text="Отмена",
                               on_release=lambda *a: pop.dismiss()))
        box.add_widget(brow)
        pop.open()

    def _popup(self, title, msg):
        Popup(title=title, content=Label(text=msg),
              size_hint=(0.8, 0.35)).open()

    def _next_id(self):
        return max((r.get("id", 0) for r in self.records), default=0) + 1


if __name__ == "__main__":
    import traceback
    try:
        EquipmentApp().run()
    except Exception:
        err = traceback.format_exc()
        # Сохранить лог в доступную папку, чтобы можно было снять с телефона
        try:
            log_path = crash_log_path()
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(err)
            print("CRASH LOG SAVED:", log_path)
        except Exception:
            pass
        # Показать на экране
        try:
            from kivy.base import runTouchApp
            from kivy.uix.label import Label as _L
            from kivy.uix.button import Button as _B
            from kivy.uix.boxlayout import BoxLayout as _BL
            from kivy.uix.popup import Popup as _P
            bl = _BL(orientation="vertical", padding=[10], spacing=[10])
            bl.add_widget(_L(text="Ошибка запуска (скопирована в crash.log):",
                             font_size="13sp"))
            bl.add_widget(_L(text=err[-1800:], font_size="9sp"))
            runTouchApp(bl)
        except Exception:
            print(err)
        raise
