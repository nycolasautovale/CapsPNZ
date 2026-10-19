#!/usr/bin/env python3
"""Monta o HTML final auto-contido a partir dos fontes em src/ e vendor/."""
import base64, pathlib

ROOT = pathlib.Path(__file__).parent
def read(p): return (ROOT/p).read_text(encoding="utf-8")
def b64img(p): return "data:image/png;base64," + base64.b64encode((ROOT/p).read_bytes()).decode()

shell   = read("src/shell.html")
css     = read("src/styles.css")
icons   = read("src/icons.js")
app     = read("src/app.js")
xlsx    = read("vendor/xlsx.full.min.js")
echarts = read("vendor/echarts.min.js")
brasao  = b64img("assets/image2.png")
saude   = b64img("assets/image1.png")

out = (shell
    .replace("__CSS__", css)
    .replace("__ICONS__", icons)
    .replace("__APP__", app)
    .replace("__XLSX__", xlsx)
    .replace("__ECHARTS__", echarts)
    .replace("__LOGO_BRASAO__", brasao)
    .replace("__LOGO_SAUDE__", saude)
    .replace("__ICON_UPLOAD__", "<svg width=\"40\" height=\"40\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"1.6\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><path d=\"M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4\"/><polyline points=\"17 8 12 3 7 8\"/><line x1=\"12\" y1=\"3\" x2=\"12\" y2=\"15\"/></svg>")
    .replace("__ICON_LOCK__", "<svg width=\"14\" height=\"14\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"1.8\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><rect x=\"3\" y=\"11\" width=\"18\" height=\"11\" rx=\"2\"/><path d=\"M7 11V7a5 5 0 0 1 10 0v4\"/></svg>")
)

dest = ROOT/"painel-caps-petrolina.html"
dest.write_text(out, encoding="utf-8")
kb = dest.stat().st_size/1024
print(f"Gerado: {dest.name}  ({kb:.0f} KB)")
