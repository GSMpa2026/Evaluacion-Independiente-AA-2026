#!/usr/bin/env python3
"""Calcula los números de página del índice estático.
Uso: python3 paginas_indice.py Informe.pdf   (lee toc.json, escribe pages.json)
Busca cada título de forma secuencial desde la página 3 para no confundirlo con menciones en tablas.
Repite generar -> convertir a PDF -> este script hasta que pages.json no cambie."""
import subprocess, json, re, sys
pdf = sys.argv[1]
n = int(re.search(r'Pages:\s+(\d+)', subprocess.run(['pdfinfo', pdf], capture_output=True, text=True).stdout).group(1))
titles = json.load(open('toc.json'))
texts = {p: subprocess.run(['pdftotext', '-f', str(p), '-l', str(p), '-layout', pdf, '-'], capture_output=True, text=True).stdout for p in range(3, n + 1)}
out, start = {}, 3
for t in titles:
    key = t[:40]
    for p in range(start, n + 1):
        if any(l.strip().startswith(key) for l in texts[p].splitlines()):
            out[t] = p; start = p; break
missing = [t for t in titles if t not in out]
json.dump(out, open('pages.json', 'w'), ensure_ascii=False)
print(json.dumps(out, ensure_ascii=False, indent=0)); print('MISSING', missing)
