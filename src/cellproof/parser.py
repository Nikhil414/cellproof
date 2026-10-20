"""Excel formula text -> small AST. Pratt parser over openpyxl's tokenizer."""
from __future__ import annotations

import hashlib
from dataclasses import dataclass
from typing import Iterator, Union

from openpyxl.formula.tokenizer import Token, Tokenizer, TokenizerError
from openpyxl.utils import get_column_letter

from cellproof.model import Unsupported


@dataclass(frozen=True)
class Num:
    value: float


@dataclass(frozen=True)
class Str:
    value: str


@dataclass(frozen=True)
class Bool:
    value: bool


@dataclass(frozen=True)
class ErrLit:
    value: str


@dataclass(frozen=True)
class RefNode:
    text: str


@dataclass(frozen=True)
class Unary:
    op: str
    operand: "Node"


@dataclass(frozen=True)
class Postfix:
    op: str
    operand: "Node"


@dataclass(frozen=True)
class Binary:
    op: str
    left: "Node"
    right: "Node"


@dataclass(frozen=True)
class Call:
    name: str
    args: tuple  # Node or None (empty argument)


Node = Union[Num, Str, Bool, ErrLit, RefNode, Unary, Postfix, Binary, Call]

# An array constant ({1,2;3,4}) becomes a reference to a hidden sheet holding its values, so every function that takes a
# range takes it unchanged. The compiler adds those cells to the book (see add_arrays); used as one value it is refused.
ARRAY_SHEET = "#array-"
ARRAYS: dict[str, tuple[tuple[tuple[str, object], ...], ...]] = {}  # hidden sheet -> rows of (kind, value)

# Excel precedence, low -> high. Unary minus binds tighter than ^, so -2^2 = 4.
BINARY = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5}
PREFIX = 6


def parse(formula: str) -> Node:
    """Parse formula text without the leading '='."""
    try:
        tokens = [t for t in Tokenizer("=" + formula).items if t.type != Token.WSPACE]
    except TokenizerError as err:
        raise Unsupported(f"cannot parse formula: {err}") from None
    p = _Parser(tokens)
    node = p.expr(0)
    if p.peek() is not None:
        raise Unsupported(f"unexpected {p.peek().value!r} in formula")
    return node


def _fixed_reference(args) -> bool:
    """INDIRECT of text written in the formula, A1 style (no second argument, or TRUE)."""
    return (len(args) in (1, 2) and isinstance(args[0], Str) and args[0].value.strip() != ""
            and (len(args) == 1 or (isinstance(args[1], Bool) and args[1].value)))


def _refs(node) -> Iterator[str]:
    if isinstance(node, RefNode):
        yield node.text
    elif isinstance(node, (Unary, Postfix)):
        yield from _refs(node.operand)
    elif isinstance(node, Binary):
        yield from _refs(node.left)
        yield from _refs(node.right)
    elif isinstance(node, Call):
        for arg in node.args:
            if arg is not None:
                yield from _refs(arg)


def _array(text: str) -> str | None:
    """The hidden sheet of an array constant's reference, else None (a real sheet may share the prefix)."""
    if not text.startswith("'" + ARRAY_SHEET):
        return None
    name = text[1:text.find("'!")]
    return name if name in ARRAYS else None


def refs(node) -> Iterator[str]:
    """Every reference text in the tree, left to right; array constants are constants, not dependencies."""
    return (t for t in _refs(node) if _array(t) is None)


def arrays(node) -> Iterator[str]:
    """The hidden sheet of every array constant in the tree."""
    return (name for t in _refs(node) if (name := _array(t)) is not None)


class _Parser:
    def __init__(self, tokens):
        self.tokens, self.i = tokens, 0

    def peek(self):
        return self.tokens[self.i] if self.i < len(self.tokens) else None

    def take(self):
        tok = self.peek()
        if tok is None:
            raise Unsupported("formula ends unexpectedly")
        self.i += 1
        return tok

    def at(self, type_, subtype):
        tok = self.peek()
        return tok is not None and tok.type == type_ and tok.subtype == subtype

    def expr(self, min_bp):
        left = self.atom()
        while (tok := self.peek()) is not None:
            if tok.type == Token.OP_POST:
                self.take()
                left = Postfix(tok.value, left)
                continue
            if tok.type != Token.OP_IN:
                break
            bp = BINARY.get(tok.value)
            if bp is None:
                raise Unsupported(f"operator {tok.value!r} not supported")
            if bp <= min_bp:
                break
            self.take()
            left = Binary(tok.value, left, self.expr(bp))
        return left

    def atom(self):
        tok = self.take()
        if tok.type == Token.OP_PRE:
            operand = self.expr(PREFIX)
            return operand if tok.value == "+" else Unary("-", operand)
        if tok.type == Token.OPERAND:
            if tok.subtype == Token.NUMBER:
                return Num(float(tok.value))
            if tok.subtype == Token.TEXT:
                return Str(tok.value[1:-1].replace('""', '"'))
            if tok.subtype == Token.LOGICAL:
                return Bool(tok.value.upper() == "TRUE")
            if tok.subtype == Token.ERROR:
                return ErrLit(tok.value)
            return RefNode(tok.value)
        if tok.type == Token.PAREN and tok.subtype == Token.OPEN:
            inner = self.expr(0)
            if not self.at(Token.PAREN, Token.CLOSE):
                raise Unsupported("missing ')'")
            self.take()
            return inner
        if tok.type == Token.FUNC and tok.subtype == Token.OPEN:
            return self.call(tok.value[:-1].upper().removeprefix("_XLFN."))
        if tok.type == Token.ARRAY and tok.subtype == Token.OPEN:
            return self.array()
        raise Unsupported(f"unexpected {tok.value!r} in formula")

    def call(self, name):
        if self.at(Token.FUNC, Token.CLOSE):
            self.take()
            return Call(name, ())
        args = []
        while True:
            if self.at(Token.SEP, Token.ARG) or self.at(Token.FUNC, Token.CLOSE):
                args.append(None)
            else:
                args.append(self.expr(0))
            tok = self.take()
            if tok.type == Token.FUNC and tok.subtype == Token.CLOSE:
                if name == "INDIRECT" and _fixed_reference(args):  # INDIRECT("Sales") is just the range Sales
                    return RefNode(args[0].value)
                return Call(name, tuple(args))
            if not (tok.type == Token.SEP and tok.subtype == Token.ARG):
                raise Unsupported(f"unexpected {tok.value!r} in {name}()")

    def array(self):
        rows, row = [], []
        while True:
            tok = self.take()
            sign = 1
            if tok.type == Token.OP_PRE and tok.value in "+-":
                sign, tok = (-1 if tok.value == "-" else 1), self.take()
            if tok.type != Token.OPERAND or tok.subtype == Token.RANGE or (sign != 1 and tok.subtype != Token.NUMBER):
                raise Unsupported(f"unexpected {tok.value!r} in array constant")
            if tok.subtype == Token.NUMBER:
                row.append(("n", sign * float(tok.value)))
            elif tok.subtype == Token.TEXT:
                if tok.value == '""':  # the cells table cannot tell empty text from no text
                    raise Unsupported("array constant holding empty text \"\"")
                row.append(("s", tok.value[1:-1].replace('""', '"')))
            elif tok.subtype == Token.LOGICAL:
                row.append(("b", tok.value.upper() == "TRUE"))
            else:
                row.append(("e", tok.value.upper()))
            sep = self.take()
            if sep.type == Token.SEP and sep.subtype == Token.ARG:
                continue
            rows.append(tuple(row))
            row = []
            if sep.type == Token.SEP and sep.subtype == Token.ROW:
                continue
            if sep.type == Token.ARRAY and sep.subtype == Token.CLOSE:
                break
            raise Unsupported(f"unexpected {sep.value!r} in array constant")
        if len({len(r) for r in rows}) != 1:
            raise Unsupported("array constant with rows of different lengths")
        values = tuple(rows)
        name = ARRAY_SHEET + hashlib.sha1(repr(values).encode()).hexdigest()[:16]
        ARRAYS[name] = values
        end = get_column_letter(len(values[0])) + str(len(values))
        return RefNode(f"'{name}'!A1" + ("" if end == "A1" else f":{end}"))
