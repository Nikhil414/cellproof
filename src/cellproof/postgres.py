"""The Postgres target: translate each verified DuckDB batch, run it on Postgres, verify every cell again.

See docs/superpowers/specs/2026-10-31-postgres-target-design.md. Needs the `postgres` extra (sqlglot, psycopg;
pgserver for a Postgres of its own when no URL is given).
"""
from __future__ import annotations

import tempfile
from collections import defaultdict
from contextlib import contextmanager
from pathlib import Path

from cellproof.compiler import Compiled, batch_rows, input_rows
from cellproof.model import Book, Key, Unsupported, Verdict, a1
from cellproof.refs import resolve

NOT_ON_PG = "not yet on Postgres"


class Untranslatable(Exception):
    pass


# Functions with DuckDB's behaviour, created in pg_temp (session-scoped: nothing stays in the database). Postgres
# never finds a pg_temp function by its bare name, so the translation calls them as pg_temp.<name>.
PRELUDE = r"""
-- the exact decimal value of |v|, from its IEEE-754 bits (DuckDB's printf rounds this, not the shortest form)
CREATE OR REPLACE FUNCTION pg_temp.cellproof_exact(v double precision) RETURNS text
LANGUAGE plpgsql IMMUTABLE AS $$
DECLARE b bigint; ex int; m numeric; k int; s text;
BEGIN
  b := ('x' || encode(float8send(v), 'hex'))::bit(64)::bigint;
  ex := ((b >> 52) & 2047)::int;
  m := (b & 4503599627370495)::numeric;
  IF ex = 0 THEN k := 1074; ELSE m := m + 4503599627370496; k := 1075 - ex; END IF;
  IF k <= 0 THEN RETURN trunc(m * power(2::numeric, -k))::text; END IF;
  s := trunc(m * power(5::numeric, k))::text;
  IF length(s) <= k THEN s := repeat('0', k - length(s) + 1) || s; END IF;
  RETURN substr(s, 1, length(s) - k) || '.' || rtrim(substr(s, length(s) - k + 1), '0');
END $$;

-- "int.frac" rounded to n fraction digits, ties to even, as text with exactly n fraction digits
CREATE OR REPLACE FUNCTION pg_temp.cellproof_round(ip text, fp text, n int) RETURNS text
LANGUAGE plpgsql IMMUTABLE AS $$
DECLARE kept text; rest text; up boolean; r text;
BEGIN
  kept := ip || rpad(substr(fp, 1, n), n, '0');
  rest := substr(fp, n + 1);
  up := rest ~ '^[6-9]' OR rest ~ '^5.*[1-9]' OR (rest ~ '^50*$' AND right(kept, 1) ~ '[13579]');
  r := CASE WHEN up THEN (kept::numeric + 1)::text ELSE ltrim(kept, '0') END;
  IF length(r) < n + 1 THEN r := lpad(r, n + 1, '0'); END IF;  -- lpad would cut a longer r
  RETURN CASE WHEN n = 0 THEN r ELSE substr(r, 1, length(r) - n) || '.' || right(r, n) END;
END $$;

-- DuckDB's printf for the formats cellproof writes: %.<p>f, %.<p>e and %.<p>g
CREATE OR REPLACE FUNCTION pg_temp.cellproof_printf(fmt text, v double precision) RETURNS text
LANGUAGE plpgsql IMMUTABLE AS $$
DECLARE
  m text[] := regexp_match(fmt, '^%\.(\d+)([efg])$');
  p int; x text; ip text; fp text; e int; d text; mant text; neg text;
BEGIN
  IF v IS NULL THEN RETURN NULL; END IF;
  IF m IS NULL THEN RAISE EXCEPTION 'cellproof_printf: format % not supported', fmt; END IF;
  p := m[1]::int;
  neg := CASE WHEN v < 0 OR (v = 0 AND float8send(v) <> float8send(0::float8)) THEN '-' ELSE '' END;
  x := pg_temp.cellproof_exact(v);
  ip := split_part(x, '.', 1); fp := split_part(x, '.', 2);
  IF m[2] = 'f' THEN RETURN neg || pg_temp.cellproof_round(ip, fp, p); END IF;
  IF m[2] = 'g' THEN
    IF p = 0 THEN p := 1; END IF;
    e := split_part(pg_temp.cellproof_printf('%.' || (p - 1) || 'e', v), 'e', 2)::int;
    IF e < -4 OR e >= p THEN
      d := pg_temp.cellproof_printf('%.' || (p - 1) || 'e', v);
      RETURN regexp_replace(split_part(d, 'e', 1), '\.?0+$', '') || 'e' || split_part(d, 'e', 2);
    END IF;
    d := pg_temp.cellproof_printf('%.' || (p - 1 - e) || 'f', v);
    RETURN CASE WHEN position('.' in d) > 0 THEN regexp_replace(d, '\.?0+$', '') ELSE d END;
  END IF;
  -- %e: the first significant digit, p more, ties to even; 9.99 -> 10.0 moves the exponent
  IF v = 0 THEN RETURN neg || '0' || CASE WHEN p > 0 THEN '.' || repeat('0', p) ELSE '' END || 'e+00'; END IF;
  IF ltrim(ip, '0') <> '' THEN
    d := ltrim(ip, '0') || fp; e := length(ltrim(ip, '0')) - 1;
  ELSE
    e := -(length(fp) - length(ltrim(fp, '0')) + 1); d := ltrim(fp, '0');
  END IF;
  mant := pg_temp.cellproof_round(substr(d, 1, 1), substr(d, 2), p);
  IF split_part(mant, '.', 1) = '10' THEN
    e := e + 1;
    mant := '1' || CASE WHEN p > 0 THEN '.' || repeat('0', p) ELSE '' END;
  END IF;
  RETURN neg || mant || 'e' || CASE WHEN e < 0 THEN '-' ELSE '+' END || CASE WHEN abs(e) < 10 THEN '0' ELSE '' END || abs(e)::text;
END $$;

-- DuckDB's % on doubles is C's fmod: the exact remainder, with the dividend's sign. Postgres's numeric mod is exact
-- on the exact values, and the remainder of two doubles is itself a double.
CREATE OR REPLACE FUNCTION pg_temp.cellproof_numeric(v double precision) RETURNS numeric
LANGUAGE sql IMMUTABLE AS $$
  SELECT CASE WHEN v < 0 THEN -1 ELSE 1 END * pg_temp.cellproof_exact(v)::numeric $$;
CREATE OR REPLACE FUNCTION pg_temp.cellproof_fmod(a double precision, b double precision) RETURNS double precision
LANGUAGE sql IMMUTABLE AS $$
  SELECT CASE WHEN b = 0 THEN NULL ELSE mod(pg_temp.cellproof_numeric(a), pg_temp.cellproof_numeric(b))::float8 END $$;

-- Postgres raises where DuckDB returns inf/NaN/NULL (power(0, -1), exp(1000), ln(0), sqrt(-1)), and it folds
-- constants while planning, so even a CASE branch cellproof's guards never take can raise. These never raise.
CREATE OR REPLACE FUNCTION pg_temp.cellproof_pow(a double precision, b double precision) RETURNS double precision
LANGUAGE plpgsql IMMUTABLE AS $$
BEGIN
  IF a = 0 AND b < 0 THEN RETURN 'Infinity'; END IF;
  IF a < 0 AND b <> trunc(b) THEN RETURN 'NaN'; END IF;
  RETURN power(a, b);
EXCEPTION WHEN numeric_value_out_of_range THEN
  RETURN CASE WHEN a < 0 AND mod(b::numeric, 2) = 1 THEN '-Infinity'::float8 WHEN abs(a) < 1 THEN 0 ELSE 'Infinity' END;
END $$;
CREATE OR REPLACE FUNCTION pg_temp.cellproof_exp(x double precision) RETURNS double precision
LANGUAGE sql IMMUTABLE AS $$ SELECT CASE WHEN x > 709.782712893384 THEN 'Infinity'::float8
                                         WHEN x < -745.2 THEN 0 ELSE exp(x) END $$;
CREATE OR REPLACE FUNCTION pg_temp.cellproof_ln(x double precision) RETURNS double precision
LANGUAGE sql IMMUTABLE AS $$ SELECT CASE WHEN x > 0 THEN ln(x) END $$;
CREATE OR REPLACE FUNCTION pg_temp.cellproof_log10(x double precision) RETURNS double precision
LANGUAGE sql IMMUTABLE AS $$ SELECT CASE WHEN x > 0 THEN log(x) END $$;
CREATE OR REPLACE FUNCTION pg_temp.cellproof_sqrt(x double precision) RETURNS double precision
LANGUAGE sql IMMUTABLE AS $$ SELECT CASE WHEN x >= 0 THEN sqrt(x) END $$;
CREATE OR REPLACE FUNCTION pg_temp.cellproof_isfinite(x double precision) RETURNS boolean
LANGUAGE sql IMMUTABLE AS $$ SELECT x NOT IN ('Infinity'::float8, '-Infinity'::float8, 'NaN'::float8) $$;

-- DuckDB's substr: a negative start counts from the end, a negative length takes the characters before the start,
-- and everything is clipped to the text (Postgres raises on a negative length)
CREATE OR REPLACE FUNCTION pg_temp.cellproof_substr(t text, a bigint, k bigint) RETURNS text
LANGUAGE sql IMMUTABLE AS $$
  SELECT CASE WHEN t IS NULL OR a IS NULL OR k IS NULL THEN NULL ELSE (
    SELECT CASE WHEN k >= 0 THEN substr(t, greatest(s, 1)::int, greatest(least(s + k, length(t) + 1) - greatest(s, 1), 0)::int)
                ELSE substr(t, greatest(s + k, 1)::int, greatest(least(s, length(t) + 1) - greatest(s + k, 1), 0)::int) END
    FROM (SELECT CASE WHEN a < 0 THEN length(t) + 1 + a ELSE a END AS s) b) END $$;
CREATE OR REPLACE FUNCTION pg_temp.cellproof_substr(t text, a bigint) RETURNS text
LANGUAGE sql IMMUTABLE AS $$ SELECT pg_temp.cellproof_substr(t, a, 9223372036854775807 / 2) $$;

-- DuckDB's regexp_escape: a backslash before every character but letters, digits, _ and non-ASCII ones
CREATE OR REPLACE FUNCTION pg_temp.cellproof_regexp_escape(t text) RETURNS text
LANGUAGE sql IMMUTABLE AS $$ SELECT regexp_replace(t, '([^A-Za-z0-9_\u0080-\U0010ffff])', '\\\1', 'g') $$;

-- DuckDB's / divides as doubles, and x / 0 is inf, -inf or NaN (Postgres raises, and divides integers as integers)
CREATE OR REPLACE FUNCTION pg_temp.cellproof_div(a double precision, b double precision) RETURNS double precision
LANGUAGE sql IMMUTABLE AS $$
  SELECT CASE WHEN b <> 0 OR b IS NULL OR a IS NULL THEN a / b WHEN a > 0 THEN 'Infinity'::float8
              WHEN a < 0 THEN '-Infinity'::float8 ELSE 'NaN'::float8 END $$;

-- DuckDB's string_split: '' splits into characters, and '' itself is [''] (string_to_array gives the whole text, {})
CREATE OR REPLACE FUNCTION pg_temp.cellproof_split(t text, d text) RETURNS text[]
LANGUAGE sql IMMUTABLE AS $$
  SELECT CASE WHEN t = '' THEN ARRAY['']::text[] WHEN d = '' THEN regexp_split_to_array(t, '')
              ELSE string_to_array(t, d) END $$;

-- DuckDB's regexp_extract: group g of the first match, '' when there is none (Postgres's regexp_match gives NULL)
CREATE OR REPLACE FUNCTION pg_temp.cellproof_regexp_extract(t text, p text, g integer) RETURNS text
LANGUAGE sql IMMUTABLE AS $$
  SELECT CASE WHEN t IS NULL OR p IS NULL THEN NULL WHEN g = 0 THEN COALESCE(substring(t from p), '')
              ELSE COALESCE((regexp_match(t, p))[g], '') END $$;

-- RE2's \p{L} (PROPER's "is a letter"): exact for ASCII and for letters with case; a caseless non-ASCII character
-- (CJK, Arabic, ...) raises, so its batch is UNRESOLVED rather than guessed
CREATE OR REPLACE FUNCTION pg_temp.cellproof_is_letter(c text) RETURNS boolean
LANGUAGE plpgsql IMMUTABLE AS $$
BEGIN
  IF c IS NULL THEN RETURN NULL; END IF;
  IF c ~ '^[\x01-\x7f]$' THEN RETURN c ~ '[A-Za-z]'; END IF;
  IF lower(c) <> upper(c) THEN RETURN true; END IF;
  RAISE EXCEPTION 'cellproof: is % a letter? (no Unicode letter class on Postgres)', c;
END $$;

CREATE OR REPLACE AGGREGATE pg_temp.cellproof_product(double precision) (SFUNC = float8mul, STYPE = double precision);
"""


_DIALECTS = None


def _dialects():
    """DuckDB and Postgres as sqlglot reads/writes them, but with list indexes passed through untouched: both are
    1-based, and sqlglot's 1-based -> 0-based -> 1-based round trip broke xs[LEAST(i + 1, n)] into xs[LEAST(...) + -1]
    (PERCENTILE on Postgres took the wrong element; the second verification caught it)."""
    global _DIALECTS
    if _DIALECTS is None:
        from sqlglot.dialects.duckdb import DuckDB
        from sqlglot.dialects.postgres import Postgres

        class Duck(DuckDB):
            INDEX_OFFSET = 0

        class Pg(Postgres):
            INDEX_OFFSET = 0
        _DIALECTS = (Duck, Pg)
    return _DIALECTS


def _is_int_cast(e) -> bool:
    from sqlglot import exp
    return isinstance(e, exp.Cast) and e.to.this == exp.DataType.Type.INT


def _is_array(e) -> bool:
    """Whether e gives a list (so DuckDB's len() is Postgres's cardinality(), not length())."""
    from sqlglot import exp
    if isinstance(e, (exp.Paren, exp.Subquery)):
        return _is_array(e.this)
    if isinstance(e, exp.Select):
        return len(e.expressions) == 1 and _is_array(e.expressions[0].unalias())
    if isinstance(e, exp.Coalesce):
        return _is_array(e.this)
    if isinstance(e, exp.Bracket):  # a slice of a list is a list; one element is not
        return any(isinstance(x, exp.Slice) for x in e.expressions)
    return isinstance(e, (exp.ArrayAgg, exp.Array, exp.Transform, exp.ArrayFilter, exp.ArraySort)) or (
        isinstance(e, exp.Anonymous) and e.name.lower() in ("string_to_array", "array", "pg_temp.cellproof_split"))


def _unnest(arr, body: str, names, where: str = "", order: str = "i", agg: bool = True):
    """ARRAY(SELECT body FROM unnest(arr) WITH ORDINALITY AS u(x, i) [WHERE ...] ORDER BY i) - a DuckDB lambda over
    a list, with the lambda's own names; agg=False: the SELECT alone (a scalar subquery's body)."""
    import sqlglot
    from sqlglot import exp
    x, i = (list(names) + ["cellproof_i"])[:2]
    sql = (f"SELECT {body} FROM unnest(cellproof_arr) WITH ORDINALITY AS cellproof_u({x}, {i})"
           + (f" WHERE {where}" if where else "") + (f" ORDER BY {order.replace('i', i)}" if order else ""))
    tree = sqlglot.parse_one(f"SELECT ARRAY({sql})" if agg else sql, read=_dialects()[1])
    tree = tree.transform(lambda n: arr.copy() if isinstance(n, exp.Column) and n.name == "cellproof_arr" else n)
    return tree.expressions[0] if agg else tree


def _rewrite(tree):
    """sqlglot tree rewrites for what DuckDB has and Postgres has not. Every rewrite is idempotent and the tree is
    rewritten until nothing changes: sqlglot does not visit inside a node a rewrite just replaced."""
    import sqlglot
    from sqlglot import exp

    def as_int(e):
        return e if _is_int_cast(e) else exp.Cast(this=e, to=exp.DataType.build("INT"))

    def lam(node):
        f = node.expression
        return f.this, [n.name for n in f.expressions]

    def visit(node):
        if isinstance(node, exp.Anonymous):
            name = node.name.upper()
            if name == "PRINTF":
                return exp.Anonymous(this="pg_temp.cellproof_printf", expressions=node.expressions)
            if name == "PRODUCT":
                return exp.Anonymous(this="pg_temp.cellproof_product", expressions=node.expressions)
            if name == "ISFINITE":
                return exp.Anonymous(this="pg_temp.cellproof_isfinite", expressions=node.expressions)
            if name == "TO_MONTHS":
                return exp.Anonymous(this="make_interval", expressions=[
                    exp.PropertyEQ(this=exp.var("months"), expression=as_int(node.expressions[0]))])
            if name == "LIST_SUM":  # empty: NULL, as DuckDB
                return exp.Subquery(this=_unnest(node.expressions[0], "sum(x)", ["x"], order="", agg=False))
            if name == "LIST_ZIP":
                raise Untranslatable("list_zip outside a list_transform has no Postgres form yet")
        if isinstance(node, (exp.Max, exp.Min)) and isinstance(node.this, (exp.Predicate, exp.Not, exp.Boolean)) \
                and not node.expressions:  # max/min of TRUE/FALSE: Postgres has only bool_or/bool_and (same values)
            return exp.Anonymous(this="bool_or" if isinstance(node, exp.Max) else "bool_and", expressions=[node.this])
        if isinstance(node, exp.RegexpExtract):  # sqlglot parses DuckDB's regexp_extract into this, not Anonymous
            group = node.args.get("group") or exp.Literal.number(0)
            return exp.Anonymous(this="pg_temp.cellproof_regexp_extract",
                                 expressions=[node.this, node.expression, as_int(group)])
        if isinstance(node, exp.Transform) and isinstance(node.this, exp.Anonymous) and node.this.name.upper() == "LIST_ZIP":
            # list_transform(list_zip(a, b), lambda x: ... x[1] ... x[2] ...): the two lists side by side
            body, names = lam(node)
            x = names[0]
            body = body.transform(lambda n: exp.column(f"cellproof_z{n.expressions[0].name}")
                                  if isinstance(n, exp.Bracket) and isinstance(n.this, (exp.Column, exp.Identifier)) and n.this.name == x
                                  and len(n.expressions) == 1 and isinstance(n.expressions[0], exp.Literal) else n)
            a, b = node.this.expressions
            sql = (f"SELECT ARRAY(SELECT {body.sql(dialect=_dialects()[1])} FROM unnest(cellproof_a) WITH ORDINALITY "
                   f"AS za(cellproof_z1, i) JOIN unnest(cellproof_b) WITH ORDINALITY AS zb(cellproof_z2, j) ON i = j "
                   f"ORDER BY i)")
            tree = sqlglot.parse_one(sql, read=_dialects()[1])
            tree = tree.transform(lambda n: (a if n.name == "cellproof_a" else b).copy()
                                  if isinstance(n, exp.Column) and n.name in ("cellproof_a", "cellproof_b") else n)
            return tree.expressions[0]
        if isinstance(node, exp.RegexpLike) and isinstance(node.expression, exp.Literal) and \
                node.expression.name in ("\\p{L}", "\\\\p{L}"):
            return exp.Anonymous(this="pg_temp.cellproof_is_letter", expressions=[node.this])
        if isinstance(node, exp.GenerateSeries) and node.args.get("is_end_exclusive"):
            # DuckDB's range(a, b[, s]) stops before b: the last value is b - sign(s) (sqlglot wrote b - 1, wrong for a
            # negative step); used as a value (unnest(range(...))) it is a list, in FROM a set
            start, end, step = node.args["start"], node.args["end"], node.args.get("step")
            last = (exp.Sub(this=end, expression=exp.Literal.number(1)) if step is None else
                    as_int(exp.Sub(this=end, expression=exp.Anonymous(this="sign", expressions=[step.copy()]))))  # sign() is float8
            series = exp.Anonymous(this="generate_series", expressions=[start, last] + ([step] if step is not None else []))
            in_from = node.find_ancestor(exp.From, exp.Join) is not None and node.find_ancestor(exp.Unnest) is None
            if in_from:
                return series
            return sqlglot.parse_one("SELECT ARRAY(SELECT * FROM generate_series(1, 1))", read=_dialects()[1]) \
                .expressions[0].transform(lambda n: series.copy() if isinstance(n, exp.GenerateSeries) else n)
        if isinstance(node, exp.Transform):
            body, names = lam(node)
            return _unnest(node.this, body.sql(dialect=_dialects()[1]), names)
        if isinstance(node, exp.ArrayFilter):
            body, names = lam(node)
            return _unnest(node.this, names[0], names, where=body.sql(dialect=_dialects()[1]))
        if isinstance(node, exp.ArraySort):
            asc = node.args.get("asc")
            desc = asc is not None and asc.this is False  # list_reverse_sort, list_sort(xs, 'DESC')
            return _unnest(node.this, "x", ["x"], order="x DESC" if desc else "x")
        if isinstance(node, exp.Length) and _is_array(node.this):
            return exp.Anonymous(this="cardinality", expressions=[node.this])
        if isinstance(node, exp.Array) and not node.expressions:
            return exp.Literal.string("{}")  # an untyped '{}' takes the other side's array type (COALESCE)
        if isinstance(node, exp.Bracket):
            if isinstance(node.this, (exp.Func, exp.Anonymous)) and not isinstance(node.this, exp.Paren):
                node.set("this", exp.Paren(this=node.this))  # Postgres: (f(x))[i], not f(x)[i]
            return node  # Postgres takes bigint subscripts: sqlglot's index offsets stay untouched
        if isinstance(node, exp.Substring):
            args = [node.this, node.args["start"]] + ([node.args["length"]] if node.args.get("length") is not None else [])
            return exp.Anonymous(this="pg_temp.cellproof_substr", expressions=args)
        if isinstance(node, exp.Anonymous) and node.name.upper() == "REGEXP_ESCAPE":
            return exp.Anonymous(this="pg_temp.cellproof_regexp_escape", expressions=node.expressions)
        if isinstance(node, exp.Div) and not node.args.get("typed"):
            return exp.Anonymous(this="pg_temp.cellproof_div", expressions=[node.this, node.expression])
        if isinstance(node, (exp.Left, exp.Right)):
            node.set("expression", as_int(node.expression))
            return node
        safe = {exp.Pow: "pow", exp.Exp: "exp", exp.Ln: "ln", exp.Sqrt: "sqrt"}
        if type(node) in safe:
            args = [node.this] + ([node.expression] if isinstance(node, exp.Pow) else [])
            return exp.Anonymous(this=f"pg_temp.cellproof_{safe[type(node)]}", expressions=args)
        if isinstance(node, exp.Log) and isinstance(node.this, exp.Literal) and node.this.name == "10":
            return exp.Anonymous(this="pg_temp.cellproof_log10", expressions=[node.expression])  # log10(x)
        if isinstance(node, exp.ToDays):
            return exp.Anonymous(this="make_interval", expressions=[
                exp.PropertyEQ(this=exp.var("days"), expression=as_int(node.this))])
        if isinstance(node, exp.Repeat):
            node.set("times", as_int(node.args["times"]))
            return node
        if isinstance(node, exp.Round) and node.args.get("decimals") is None and not (
                isinstance(node.this, exp.Cast) and node.this.to.this == exp.DataType.Type.DECIMAL):
            # Postgres's round(double) rounds ties to even (100.5 -> 100); numeric's, like DuckDB's, away from zero
            inner = exp.Round(this=exp.Cast(this=node.this, to=exp.DataType.build("NUMERIC")))
            return exp.Cast(this=inner, to=exp.DataType.build("DOUBLE PRECISION"))
        if isinstance(node, exp.Round) and node.args.get("decimals") is not None and not (
                isinstance(node.this, exp.Cast) and node.this.to.this == exp.DataType.Type.DECIMAL):
            # Postgres rounds numeric only: decimal rounding at the 15 digits a double casts to, as Excel does
            # (ROUND(1.005, 2) is 1.01 in Excel and here; DuckDB's binary rounding gives 1.0)
            inner = exp.Round(this=exp.Cast(this=node.this, to=exp.DataType.build("NUMERIC")),
                              decimals=as_int(node.args["decimals"]))
            return exp.Cast(this=inner, to=exp.DataType.build("DOUBLE PRECISION"))
        if isinstance(node, exp.Mod):
            return exp.Anonymous(this="pg_temp.cellproof_fmod", expressions=[node.this, node.expression])
        if isinstance(node, exp.DateFromParts):
            return exp.Anonymous(this="make_date", expressions=[as_int(node.args[k]) for k in ("year", "month", "day")])
        if isinstance(node, exp.Split):
            return exp.Anonymous(this="pg_temp.cellproof_split", expressions=[node.this, node.expression])
        if isinstance(node, exp.Select) and len(node.expressions) == 1:
            agg = node.expressions[0].unalias()
            if isinstance(agg, (exp.ArgMax, exp.ArgMin)) and not node.args.get("group"):
                # arg_max(v, k): the v of the row with the largest k, skipping NULL v (DuckDB) - one row or none
                value, key = agg.this, agg.expression
                order = exp.Ordered(this=key, desc=isinstance(agg, exp.ArgMax), nulls_first=False)
                return (node.select(value, append=False)
                        .where(exp.Not(this=exp.Is(this=value.copy(), expression=exp.Null()))).order_by(order).limit(1))
        return node

    for _ in range(12):
        before = tree.sql(dialect=_dialects()[1])
        tree = tree.transform(visit)
        if tree.sql(dialect=_dialects()[1]) == before:
            return tree
    raise Untranslatable("the rewrite did not settle")


def translate(sql: str) -> str:
    """DuckDB SQL of one batch -> Postgres SQL."""
    import sqlglot
    try:
        duck, pg = _dialects()
        return _rewrite(sqlglot.parse_one(sql, read=duck)).sql(dialect=pg)
    except Exception as err:  # noqa: BLE001 - sqlglot raises several types
        raise Untranslatable(f"{type(err).__name__}: {str(err).splitlines()[0][:120]}") from err


def _readers(book: Book) -> dict[Key, set[Key]]:
    """Formula cell -> the formula cells that read it (one step)."""
    from cellproof.r1c1 import lexed
    formulas = [k for k, c in book.cells.items() if c.formula is not None]
    by_sheet = defaultdict(list)
    for k in formulas:
        by_sheet[k[0]].append(k)
    out: dict[Key, set[Key]] = defaultdict(set)
    lex = lexed(book)
    for key in formulas:
        shape, texts = lex.get(key, (None, ()))
        for text in texts:
            if text[:6].lower() == "_xlpm.":
                continue
            try:
                area = resolve(text, key[0], book)
            except Unsupported:
                continue
            for other in by_sheet.get(area.sheet, ()):
                if area.r1 <= other[1] <= area.r2 and area.c1 <= other[2] <= area.c2:
                    out[other].add(key)
    return out


def _spread(book: Book, failed: dict[Key, str]) -> dict[Key, str]:
    """failed cells plus every cell reading them, transitively, with the reason."""
    readers = _readers(book)
    out, todo = dict(failed), list(failed)
    while todo:
        key = todo.pop()
        for r in readers.get(key, ()):
            if r not in out:
                out[r] = f"{NOT_ON_PG}: depends on {a1(key)}, which is not"
                todo.append(r)
    return out


@contextmanager
def connect(url: str | None = None):
    """A psycopg connection: to `url`, or to a Postgres of cellproof's own (pgserver) in a temporary folder that is
    deleted afterwards. Nothing stays in a given database either: the cells table is temporary."""
    import psycopg
    if url is not None:
        with psycopg.connect(url, autocommit=True) as con:
            yield con
        return
    import pgserver
    with tempfile.TemporaryDirectory(prefix="cellproof_pg_") as tmp:
        server = pgserver.get_server(Path(tmp) / "data", cleanup_mode="stop")
        try:
            with psycopg.connect(server.get_uri(), autocommit=True) as con:
                yield con
        finally:
            server.cleanup()


def execute(book: Book, compiled: Compiled, con) -> tuple[dict[Key, tuple], dict[Key, str]]:
    """Run every batch on Postgres in a temporary cells table. -> (results, cells not run with why)."""
    con.execute(PRELUDE)
    con.execute("CREATE TEMP TABLE cells (sheet TEXT, row_num BIGINT, col_num BIGINT, kind TEXT, "
                "num DOUBLE PRECISION, txt TEXT) ON COMMIT PRESERVE ROWS")
    try:
        with con.cursor().copy("COPY cells FROM STDIN") as cp:
            for row in input_rows(book):
                cp.write_row(row)
        failed: dict[Key, str] = {}
        got: dict[Key, tuple] = {}
        for batch in compiled.batches:
            keys = [s.key for s in batch]
            try:
                rows = con.execute(translate(batch_rows(batch, "cells"))).fetchall()
            except Untranslatable as err:
                failed.update({k: f"{NOT_ON_PG}: the SQL does not translate ({err})" for k in keys})
                continue
            except Exception as err:  # noqa: BLE001 - psycopg's errors: the batch did not run
                failed.update({k: f"{NOT_ON_PG}: Postgres refused the SQL ({str(err).splitlines()[0][:160]})"
                               for k in keys})
                continue
            with con.cursor().copy("COPY cells FROM STDIN") as cp:
                for row in rows:
                    cp.write_row(row)
            got.update({(s, r, c): (k, n, t) for s, r, c, k, n, t in rows})
        return got, failed
    finally:
        con.execute("DROP TABLE IF EXISTS cells")


MAX_ROUNDS = 20


def verify_book(book: Book, compile_fn, con) -> tuple[Compiled, list[Verdict]]:
    """(compiled, verdicts) on Postgres, where `compiled` holds only SQL Postgres ran: a batch that does not translate
    or run makes its cells UNRESOLVED with the reason and the book is compiled again, so every cell reading them is
    UNRESOLVED too (the DuckDB verifier's rule). The dbt project written from `compiled` then never reads a missing
    value. Each round removes at least one batch."""
    from cellproof.verifier import verdicts
    for _ in range(MAX_ROUNDS):
        compiled = compile_fn(book)
        got, failed = execute(book, compiled, con)
        if not failed:
            return compiled, verdicts(book, compiled, got)
        book.unread.update(failed)
    raise RuntimeError(f"Postgres still refused SQL after {MAX_ROUNDS} rounds; please report it")


def dbt_env(con, schema: str | None = None) -> dict[str, str]:
    """The environment the generated dbt-postgres profile reads, pointing at `con`'s database."""
    i = con.info
    env = {"CELLPROOF_PG_HOST": i.host, "CELLPROOF_PG_PORT": str(i.port), "CELLPROOF_PG_USER": i.user,
           "CELLPROOF_PG_PASSWORD": i.password or "", "CELLPROOF_PG_DBNAME": i.dbname}
    return env | ({"CELLPROOF_PG_SCHEMA": schema} if schema else {})


def verify(book: Book, compiled: Compiled, con) -> list[Verdict]:
    """Every formula cell DuckDB compiled, verified again on Postgres."""
    from cellproof.verifier import verdicts
    got, failed = execute(book, compiled, con)
    return verdicts(book, compiled, got, _spread(book, failed))
