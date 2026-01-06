"""簡單的指數與對數計算器 (CLI)

功能：
- power: 計算 a^b
- exp: 計算 e^x
- log: 計算對數，支援自然對數、10 為底或自定底數

範例：
  python tt1.py power 2 3      # 2^3 = 8
  python tt1.py exp 1         # e^1
  python tt1.py log 100 --base 10

此檔案不依賴第三方套件，只使用標準函式庫
"""

from __future__ import annotations
import argparse
import math
import sys
from typing import Optional


def power(a: float, b: float) -> float:
	"""返回 a 的 b 次方（a^b）。"""
	return a ** b


def exp(x: float) -> float:
	"""返回 e 的 x 次方。"""
	return math.exp(x)


def log(x: float, base: Optional[float] = None) -> float:
	"""計算對數。

	- 如果 base 為 None：使用自然對數（ln）
	- 如果 base 為 10：使用常用對數（log10）
	- 否則使用 change-of-base 計算： log(x)/log(base)
	"""
	if x <= 0:
		raise ValueError("log: x 必須為正數")
	if base is None:
		return math.log(x)
	if base == 10:
		return math.log10(x)
	if base <= 0 or base == 1:
		raise ValueError("log: base 必須為正數且不等於 1")
	return math.log(x) / math.log(base)


def build_parser() -> argparse.ArgumentParser:
	p = argparse.ArgumentParser(prog="tt1", description="簡單的指數與對數計算器")
	sub = p.add_subparsers(dest="cmd", required=True)

	p_pow = sub.add_parser("power", help="計算 a^b")
	p_pow.add_argument("a", type=float, help="底數 a")
	p_pow.add_argument("b", type=float, help="指數 b")

	p_exp = sub.add_parser("exp", help="計算 e^x")
	p_exp.add_argument("x", type=float, help="指數 x")

	p_log = sub.add_parser("log", help="計算對數")
	p_log.add_argument("x", type=float, help="要計算對數的正數 x")
	p_log.add_argument("--base", "-b", type=float, default=None, help="底數（預設自然對數）")

	return p


def main(argv: Optional[list[str]] = None) -> int:
	argv = argv if argv is not None else sys.argv[1:]
	parser = build_parser()
	args = parser.parse_args(argv)

	try:
		if args.cmd == "power":
			res = power(args.a, args.b)
			print(res)

		elif args.cmd == "exp":
			res = exp(args.x)
			print(res)

		elif args.cmd == "log":
			res = log(args.x, args.base)
			print(res)

		else:
			parser.print_help()
			return 1
	except Exception as e:
		print(f"錯誤: {e}", file=sys.stderr)
		return 2

	return 0


if __name__ == "__main__":
	raise SystemExit(main())

