def H1(text): return {"type": "h1", "text": text}
def H2(text): return {"type": "h2", "text": text}
def H3(text): return {"type": "h3", "text": text}
def P(text, bold_prefix=None):
    d = {"type": "p", "text": text}
    if bold_prefix: d["bold_prefix"] = bold_prefix
    return d
def C(text): return {"type": "callout", "text": text}
def B(items, numbered=False, level=0): return {"type": "bullets", "items": items, "numbered": numbered, "level": level}
def T(headers, rows, widths=None, font_size=10.5, landscape=False):
    d = {"type": "table", "headers": headers, "rows": rows, "font_size": font_size, "landscape": landscape}
    if widths: d["widths"] = widths
    return d
def PB(): return {"type": "pagebreak"}
def SIG(text="本文件经学校规定程序审议通过后施行。"): return {"type": "signature", "text": text}
