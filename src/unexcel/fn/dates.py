"""DATE, YEAR, MONTH, DAY, WEEKDAY, DAYS, EDATE, EOMONTH, DAYS360, DATEDIF, and TEXT's date formats.

Excel stores a date as a serial number: 1 = 1900-01-01, and it counts a 29 February 1900 that never existed
(serial 60), so serials from 61 on are days since 1899-12-30. Serial 0 is "1900-01-00". Workbooks using the
1904 date system are not supported yet.
"""
from __future__ import annotations

from unexcel.expr import Expr, errs, num
from unexcel.fn.lookup import int_literal
from unexcel.fn.text import FORMATS
from unexcel.functions import arity, register
from unexcel.model import Unsupported
from unexcel.parser import Bool, Str

MAX_SERIAL = 2958465  # 9999-12-31


def _check(c) -> None:
    if c.book.date1904:
        raise Unsupported("1904 date system not supported yet")


def _serial(c, node, sheet: str) -> Expr:
    _check(c)
    s = num(c.expr(node, sheet))
    return Expr(f"floor({s.sql})", "num", errs(s.err, f"(CASE WHEN {s.sql} < 0 OR {s.sql} > {MAX_SERIAL} THEN '#NUM!' END)"))


def to_date(s: str) -> str:
    """A whole serial (61 and up, or 1..59) as a DATE."""
    return (f"(CASE WHEN {s} >= 61 THEN DATE '1899-12-30' + CAST({s} AS INTEGER) "
            f"ELSE DATE '1899-12-31' + CAST({s} AS INTEGER) END)")


def from_date(d: str) -> str:
    return (f"CAST(CASE WHEN {d} >= DATE '1900-03-01' THEN {d} - DATE '1899-12-30' "
            f"ELSE {d} - DATE '1899-12-31' END AS DOUBLE)")


def _ok(s: Expr) -> str:
    return f"(({s.err}) IS NULL)"


def _part(name: str, fn: str, at0: int, at60: int):
    def impl(c, args, sheet):
        arity(name, args, 1, 1)
        s = _serial(c, args[0], sheet)
        return Expr(f"(CASE WHEN NOT {_ok(s)} THEN NULL WHEN {s.sql} = 0 THEN {at0} WHEN {s.sql} = 60 THEN {at60} "
                    f"ELSE CAST({fn}({to_date(s.sql)}) AS DOUBLE) END)", "num", s.err)
    register(name)(impl)


_part("YEAR", "year", 1900, 1900)
_part("MONTH", "month", 1, 2)
_part("DAY", "day", 0, 29)


@register("DATE")
def date(c, args, sheet):
    arity("DATE", args, 3, 3)
    _check(c)
    y, m, d = (num(c.expr(a, sheet)) for a in args)
    year = f"(CASE WHEN trunc({y.sql}) < 1900 THEN trunc({y.sql}) + 1900 ELSE trunc({y.sql}) END)"
    ok = (f"(trunc({y.sql}) BETWEEN 0 AND 9999 AND abs({m.sql}) < 120000 AND abs({d.sql}) < 4000000)")
    day = (f"CAST(make_date(CAST({year} AS INTEGER), 1, 1) + to_months(CAST(trunc({m.sql}) - 1 AS INTEGER)) "
           f"+ to_days(CAST(trunc({d.sql}) - 1 AS INTEGER)) AS DATE)")
    # Real Excel: DATE(1900,2,29) is 60 - it counts a 29 February 1900 that never existed - so inside 1900 a date is
    # (days before the month, with 1900 a leap year) + day, for any day (DATE(1900,3,0) is 60). Months outside 1-12
    # use the real calendar.
    before = "[0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335]"
    in_1900 = f"({year} = 1900 AND trunc({m.sql}) BETWEEN 1 AND 12)"
    in_1900_serial = f"(CAST({before}[CAST(trunc({m.sql}) AS INTEGER)] AS DOUBLE) + trunc({d.sql}))"
    serial = f"(CASE WHEN NOT {ok} THEN NULL WHEN {in_1900} THEN {in_1900_serial} ELSE {from_date(day)} END)"
    return Expr(serial, "num", errs(y.err, m.err, d.err,
                                    f"(CASE WHEN NOT {ok} OR {serial} < 0 OR {serial} > {MAX_SERIAL} THEN '#NUM!' END)"))


@register("WEEKDAY")
def weekday(c, args, sheet):
    arity("WEEKDAY", args, 1, 2)
    s = _serial(c, args[0], sheet)
    kind = 1 if len(args) == 1 or args[1] is None else int_literal(args[1])
    sunday0 = f"(CAST({s.sql} AS BIGINT) + 6) % 7"  # serial 1 was a Sunday in Excel's calendar
    sql = {1: f"({sunday0} + 1)", 2: f"(({sunday0} + 6) % 7 + 1)", 3: f"(({sunday0} + 6) % 7)"}.get(kind)
    if sql is None:
        raise Unsupported("WEEKDAY return type must be 1, 2 or 3")
    return Expr(f"(CASE WHEN {_ok(s)} THEN CAST({sql} AS DOUBLE) END)", "num", s.err)


@register("DAYS")
def days(c, args, sheet):
    arity("DAYS", args, 2, 2)
    end, start = _serial(c, args[0], sheet), _serial(c, args[1], sheet)
    return Expr(f"({end.sql} - {start.sql})", "num", errs(end.err, start.err))


def _shift(name: str, last: bool):
    def impl(c, args, sheet):
        arity(name, args, 2, 2)
        s, m = _serial(c, args[0], sheet), num(c.expr(args[1], sheet))
        d = f"CAST({to_date(s.sql)} + to_months(CAST(trunc({m.sql}) AS INTEGER)) AS DATE)"
        if last:
            d = f"last_day({d})"
        ok = f"({_ok(s)} AND ({m.err}) IS NULL AND trunc({m.sql}) BETWEEN -120000 AND 120000)"
        serial = f"(CASE WHEN {ok} THEN {from_date(d)} END)"
        bad = f"(CASE WHEN NOT {ok} OR {serial} NOT BETWEEN 0 AND {MAX_SERIAL} THEN '#NUM!' END)"
        return Expr(serial, "num", errs(s.err, m.err, bad))
    register(name)(impl)


_shift("EDATE", last=False)
_shift("EOMONTH", last=True)


@register("DAYS360")
def days360(c, args, sheet):
    arity("DAYS360", args, 2, 3)
    if len(args) == 3 and not isinstance(args[2], Bool):
        raise Unsupported("DAYS360 method must be TRUE/FALSE written in the formula")
    european = len(args) == 3 and args[2].value
    a, b = _serial(c, args[0], sheet), _serial(c, args[1], sheet)
    d1, d2 = to_date(a.sql), to_date(b.sql)
    y1, m1, y2, m2 = f"year({d1})", f"month({d1})", f"year({d2})", f"month({d2})"
    if european:
        day1, day2, m2 = f"least(day({d1}), 30)", f"least(day({d2}), 30)", m2
    else:  # U.S. (NASD), as Microsoft documents it for DAYS360
        day1 = f"(CASE WHEN {d1} = last_day({d1}) THEN 30 ELSE day({d1}) END)"
        # Real Excel: the last day of February ends the period unchanged (28 or 29) unless the start is also the
        # last day of February; only then does it become the 30th. Only a 31st is moved (to the 1st of the next
        # month, or the 30th), never the 30th of a short month: Jan 15 -> Apr 30 is 105, not 106.
        feb_end = f"({d2} = last_day({d2}) AND month({d2}) = 2)"
        feb_start = f"({d1} = last_day({d1}) AND month({d1}) = 2)"
        end_last = f"(day({d2}) = 31)"
        day2 = (f"(CASE WHEN {feb_end} AND {feb_start} THEN 30 WHEN {end_last} AND {day1} < 30 THEN 1 "
                f"WHEN {end_last} THEN 30 ELSE day({d2}) END)")
        m2 = f"({m2} + CASE WHEN {end_last} AND {day1} < 30 THEN 1 ELSE 0 END)"
    sql = f"CAST(({y2} - {y1}) * 360 + ({m2} - {m1}) * 30 + ({day2} - {day1}) AS DOUBLE)"
    return Expr(f"(CASE WHEN {_ok(a)} AND {_ok(b)} THEN {sql} END)", "num", errs(a.err, b.err))


@register("DATEDIF")
def datedif(c, args, sheet):
    arity("DATEDIF", args, 3, 3)
    if not isinstance(args[2], Str) or args[2].value.upper() not in ("Y", "M", "D"):
        raise Unsupported("DATEDIF unit must be \"Y\", \"M\" or \"D\" written in the formula")
    unit = args[2].value.upper()
    a, b = _serial(c, args[0], sheet), _serial(c, args[1], sheet)
    d1, d2 = to_date(a.sql), to_date(b.sql)
    months = (f"((year({d2}) - year({d1})) * 12 + month({d2}) - month({d1}) "
              f"- CASE WHEN day({d2}) < day({d1}) THEN 1 ELSE 0 END)")
    sql = {"D": f"({b.sql} - {a.sql})", "M": months, "Y": f"floor({months} / 12)"}[unit]
    return Expr(f"(CASE WHEN {_ok(a)} AND {_ok(b)} AND {a.sql} <= {b.sql} THEN CAST({sql} AS DOUBLE) END)", "num",
                errs(a.err, b.err, f"(CASE WHEN {a.sql} > {b.sql} THEN '#NUM!' END)"))


def _date_format(pattern: str):
    def sql(x: str) -> str:
        s = f"floor({x})"
        return f"(CASE WHEN {s} BETWEEN 1 AND {MAX_SERIAL} AND {s} <> 60 THEN strftime({to_date(s)}, '{pattern}') END)"
    sql.is_date = True  # TEXT refuses date formats in 1904-system workbooks
    return sql


DATE_FORMATS = {
    "yyyy-mm-dd": "%Y-%m-%d", "mm/dd/yyyy": "%m/%d/%Y", "m/d/yyyy": "%-m/%-d/%Y", "dd/mm/yyyy": "%d/%m/%Y",
    "d/m/yyyy": "%-d/%-m/%Y", "dd-mmm-yyyy": "%d-%b-%Y", "d-mmm-yy": "%-d-%b-%y", "dd-mmm-yy": "%d-%b-%y",
    "mmm": "%b", "mmmm": "%B", "yyyy": "%Y", "yy": "%y", "mmm-yy": "%b-%y", "mmm yyyy": "%b %Y",
    "mmmm yyyy": "%B %Y", "ddd": "%a", "dddd": "%A", "dd": "%d", "d": "%-d", "mm": "%m",
}
FORMATS.update({k: _date_format(v) for k, v in DATE_FORMATS.items()})
