"""AST node types for Excel formulas."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional, Tuple

# Largest row index Excel supports; used for whole-column references like C:C.
MAX_ROW = 1_048_576


def col_to_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def index_to_col(idx: int) -> str:
    out = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        out = chr(65 + rem) + out
    return out


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
class Err:
    code: str


@dataclass(frozen=True)
class Ref:
    """A cell or rectangular range reference.

    A single cell has (col1, row1) == (col2, row2). A whole-column reference
    such as C:C has row1 = 1 and row2 = MAX_ROW with ``whole_col`` set.
    """

    sheet: Optional[str]
    col1: int
    row1: int
    col2: int
    row2: int
    abs_col1: bool = False
    abs_row1: bool = False
    abs_col2: bool = False
    abs_row2: bool = False
    whole_col: bool = False

    @property
    def is_cell(self) -> bool:
        return self.col1 == self.col2 and self.row1 == self.row2 and not self.whole_col

    def key(self, default_sheet: str) -> Tuple[str, int, int, int, int]:
        return (self.sheet or default_sheet, self.col1, self.row1, self.col2, self.row2)

    def a1(self) -> str:
        def part(c, r, ac, ar):
            return f"{'$' if ac else ''}{index_to_col(c)}{'$' if ar else ''}{r}"

        prefix = ""
        if self.sheet is not None:
            s = self.sheet
            if not s.replace("_", "").isalnum():
                s = "'" + s.replace("'", "''") + "'"
            prefix = s + "!"
        if self.whole_col:
            a = f"{'$' if self.abs_col1 else ''}{index_to_col(self.col1)}"
            b = f"{'$' if self.abs_col2 else ''}{index_to_col(self.col2)}"
            return f"{prefix}{a}:{b}"
        first = part(self.col1, self.row1, self.abs_col1, self.abs_row1)
        if self.is_cell:
            return prefix + first
        return f"{prefix}{first}:{part(self.col2, self.row2, self.abs_col2, self.abs_row2)}"


@dataclass(frozen=True)
class Name:
    """A defined name (named range or named constant)."""

    name: str


@dataclass(frozen=True)
class Func:
    name: str
    args: Tuple["Node", ...] = field(default_factory=tuple)


@dataclass(frozen=True)
class Unary:
    op: str  # '-' or '+'
    operand: "Node"


@dataclass(frozen=True)
class Percent:
    operand: "Node"


@dataclass(frozen=True)
class Binary:
    op: str
    left: "Node"
    right: "Node"


Node = object  # union of the dataclasses above


def walk(node):
    """Yield every node in the tree, depth first."""
    yield node
    if isinstance(node, Func):
        for a in node.args:
            yield from walk(a)
    elif isinstance(node, (Unary, Percent)):
        yield from walk(node.operand)
    elif isinstance(node, Binary):
        yield from walk(node.left)
        yield from walk(node.right)
