#!/usr/bin/env python3
# Проверка бота ПЕРЕД заливкой. Не запускает боевой сервер, не требует токенов.
# Если что-то не так — выходит с кодом 1, и заливка отменяется.
import sys, ast, py_compile

F = "bot.py"
errs = []

# 1) синтаксис
try:
    py_compile.compile(F, doraise=True)
except py_compile.PyCompileError as e:
    print("❌ Синтаксическая ошибка в bot.py:")
    print(e)
    sys.exit(1)

data = open(F, "rb").read()
src = data.decode("utf-8")

# 2) целостность файла
if b"\x00" in data:
    errs.append("в файле есть null-байты (повреждён/обрезан)")
if not src.rstrip().endswith("main()"):
    errs.append("файл обрезан: не заканчивается на main()")

# 3) разбор: все функции-обработчики должны быть определены (нет «мёртвых» ссылок)
tree = ast.parse(src)
defined = {n.name for n in ast.walk(tree)
           if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))}
used = set()
for n in ast.walk(tree):
    if isinstance(n, ast.Call) and isinstance(n.func, ast.Name) and \
       n.func.id in ("MessageHandler", "CommandHandler"):
        for a in n.args:
            if isinstance(a, ast.Name):
                used.add(a.id)
missing = sorted(h for h in used if h not in defined)
if missing:
    errs.append("обработчики не определены (мёртвые ссылки): " + ", ".join(missing))

# 4) обязательные функции на месте
for fn in ["main", "start", "handle_photo", "save_production", "save_sale", "save_expense"]:
    if fn not in defined:
        errs.append(f"нет обязательной функции: {fn}")

if errs:
    print("❌ Проверка НЕ пройдена — заливка отменена:")
    for e in errs:
        print("   -", e)
    sys.exit(1)

print("✅ Проверка пройдена: синтаксис ок, файл целый, все обработчики на месте.")
sys.exit(0)
