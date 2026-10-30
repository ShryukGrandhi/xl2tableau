"""Tokenizer and precedence-climbing parser for Excel formulas.

Operator precedence follows Excel (highest first):
negation, %, ^, * /, + -, &, comparison. Note that negation binds tighter
than ^, so ``=-2^2`` evaluates to 4 in Excel, and ^ is left associative.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import List, Optional

from .ast import (
    MAX_ROW, Binary, Bool, Err, Func, Name, Num, Percent, Ref, Str, Unary, col_to_index,
)

ERROR_CODES = ("#DIV/0!", "#N/A", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#NULL!")


class FormulaParseError(ValueError):
    pass


@dataclass
class Token:
    kind: str  # NUM STR BOOL ERR REF FUNC NAME OP LP RP COMMA
    text: str
    value: object = None


_CELL = r"\$?[A-Za-z]{1,3}\$?[0-9]+"
_COL = r"\$?[A-Za-z]{1,3}"
_SHEET_BARE = r"[A-Za-z_][A-Za-z0-9_.]*"
_SHEET_QUOTED = r"'(?:[^']|'')+'"
_SHEET = rf"(?:(?P<sheet>{_SHEET_BARE}|{_SHEET_QUOTED})!)?"
_END = r"(?![A-Za-z0-9_(.])"
CELL_RANGE_RE = re.compile(rf"{_SHEET}(?P<a>{_CELL})(?::(?P<b>{_CELL}))?{_END}")
COL_RANGE_RE = re.compile(rf"{_SHEET}(?P<a>{_COL}):(?P<b>{_COL}){_END}")
IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*")
NUMBER_RE = re.compile(r"(?:[0-9]+\.?[0-9]*|\.[0-9]+)(?:[eE][+-]?[0-9]+)?")
CELL_PARTS_RE = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)")


def _unquote_sheet(s: Optional[str]) -> Optional[str]:
    if s is None:
        return None
    if s.startswith("'"):
        return s[1:-1].replace("''", "'")
    return s


def _make_ref(m: re.Match, whole_col: bool) -> Ref:
    sheet = _unquote_sheet(m.group("sheet"))
    if whole_col:
        a, b = m.group("a"), m.group("b")
        c1, c2 = col_to_index(a.lstrip("$")), col_to_index(b.lstrip("$"))
        return Ref(sheet, min(c1, c2), 1, max(c1, c2), MAX_ROW,
                   a.startswith("$"), True, b.startswith("$"), True, whole_col=True)
    pa = CELL_PARTS_RE.fullmatch(m.group("a"))
    pb = CELL_PARTS_RE.fullmatch(m.group("b") or m.group("a"))
    c1, r1 = col_to_index(pa.group(2)), int(pa.group(4))
    c2, r2 = col_to_index(pb.group(2)), int(pb.group(4))
    if r1 < 1 or r2 < 1:
        raise FormulaParseError(f"invalid row in reference {m.group(0)!r}")
    return Ref(sheet, min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2),
               bool(pa.group(1)), bool(pa.group(3)), bool(pb.group(1)), bool(pb.group(3)))


def tokenize(text: str) -> List[Token]:
    if text.startswith("="):
        text = text[1:]
    toks: List[Token] = []
    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        if ch in " \t\r\n":
            i += 1
            continue
        if ch == '"':
            j, buf = i + 1, []
            while True:
                if j >= n:
                    raise FormulaParseError("unterminated string literal")
                if text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    break
                buf.append(text[j])
                j += 1
            toks.append(Token("STR", text[i:j + 1], "".join(buf)))
            i = j + 1
            continue
        if ch == "#":
            for code in ERROR_CODES:
                if text.upper().startswith(code, i):
                    toks.append(Token("ERR", code, code))
                    i += len(code)
                    break
            else:
                raise FormulaParseError(f"unknown error literal at {i}")
            continue
        if ch.isdigit() or (ch == "." and i + 1 < n and text[i + 1].isdigit()):
            # A digit can also start a row-range like 1:3, which we do not support.
            m = NUMBER_RE.match(text, i)
            toks.append(Token("NUM", m.group(0), float(m.group(0))))
            i = m.end()
            continue
        if ch.isalpha() or ch in "_$'":
            m = CELL_RANGE_RE.match(text, i)
            if m:
                toks.append(Token("REF", m.group(0), _make_ref(m, False)))
                i = m.end()
                continue
            m = COL_RANGE_RE.match(text, i)
            if m:
                toks.append(Token("REF", m.group(0), _make_ref(m, True)))
                i = m.end()
                continue
            m = IDENT_RE.match(text, i)
            if not m:
                raise FormulaParseError(f"unexpected character {ch!r} at {i}")
            word = m.group(0)
            j = m.end()
            if j < n and text[j] == "(":
                toks.append(Token("FUNC", word, word.upper()))
            elif word.upper() in ("TRUE", "FALSE"):
                toks.append(Token("BOOL", word, word.upper() == "TRUE"))
            else:
                toks.append(Token("NAME", word, word))
            i = j
            continue
        two = text[i:i + 2]
        if two in ("<=", ">=", "<>"):
            toks.append(Token("OP", two))
            i += 2
            continue
        if ch in "+-*/^&=<>%":
            toks.append(Token("OP", ch))
        elif ch == "(":
            toks.append(Token("LP", ch))
        elif ch == ")":
            toks.append(Token("RP", ch))
        elif ch in ",;":
            toks.append(Token("COMMA", ch))
        else:
            raise FormulaParseError(f"unexpected character {ch!r} at {i}")
        i += 1
    return toks


_BINARY_LEVELS = [
    ("=", "<>", "<", ">", "<=", ">="),
    ("&",),
    ("+", "-"),
    ("*", "/"),
    ("^",),
]


class _Parser:
    def __init__(self, toks: List[Token]):
        self.toks = toks
        self.pos = 0

    def peek(self) -> Optional[Token]:
        return self.toks[self.pos] if self.pos < len(self.toks) else None

    def take(self) -> Token:
        tok = self.peek()
        if tok is None:
            raise FormulaParseError("unexpected end of formula")
        self.pos += 1
        return tok

    def parse(self):
        if not self.toks:
            raise FormulaParseError("empty formula")
        node = self.binary(0)
        if self.peek() is not None:
            raise FormulaParseError(f"unexpected token {self.peek().text!r}")
        return node

    def binary(self, level: int):
        if level == len(_BINARY_LEVELS):
            return self.percent()
        node = self.binary(level + 1)
        ops = _BINARY_LEVELS[level]
        while True:
            tok = self.peek()
            if tok is None or tok.kind != "OP" or tok.text not in ops:
                return node
            self.take()
            node = Binary(tok.text, node, self.binary(level + 1))

    def percent(self):
        node = self.unary()
        while (tok := self.peek()) is not None and tok.kind == "OP" and tok.text == "%":
            self.take()
            node = Percent(node)
        return node

    def unary(self):
        tok = self.peek()
        if tok is not None and tok.kind == "OP" and tok.text in "+-":
            self.take()
            return Unary(tok.text, self.unary())
        return self.primary()

    def primary(self):
        tok = self.take()
        if tok.kind == "NUM":
            return Num(tok.value)
        if tok.kind == "STR":
            return Str(tok.value)
        if tok.kind == "BOOL":
            return Bool(tok.value)
        if tok.kind == "ERR":
            return Err(tok.value)
        if tok.kind == "REF":
            return tok.value
        if tok.kind == "NAME":
            return Name(tok.value)
        if tok.kind == "LP":
            node = self.binary(0)
            if self.take().kind != "RP":
                raise FormulaParseError("expected ')'")
            return node
        if tok.kind == "FUNC":
            self.take()  # '('
            args = []
            if self.peek() is not None and self.peek().kind == "RP":
                self.take()
                return Func(tok.value, ())
            while True:
                args.append(self.binary(0))
                nxt = self.take()
                if nxt.kind == "RP":
                    break
                if nxt.kind != "COMMA":
                    raise FormulaParseError(f"expected ',' or ')' in {tok.value}()")
            return Func(tok.value, tuple(args))
        raise FormulaParseError(f"unexpected token {tok.text!r}")


def parse(text: str):
    """Parse an Excel formula (with or without the leading '=') into an AST."""
    return _Parser(tokenize(text)).parse()
