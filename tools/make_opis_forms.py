# Копия генератора из D:\Загрузки\grist mods\accounting\blank (v0.7.0: ширины колонок эл. описи под MD5).
# Запуск из корня репозитория: python tools/make_opis_forms.py .
# Параметрические бланки описей на выдачу документации (по образцу «Опись копий.docx» и «Опись 13-2025.docx»).
#   opis_copies_1.pdf / opis_copies_1a.pdf — опись копий (Обозначение | Кол-во лис. | Кол-во экз.)
#   opis_edocs_1.pdf  / opis_edocs_1a.pdf  — опись электронных документов (том + Инв. № | Обозначение | Изм./Расширение | Сигнатура | лис. | экз.)
# *_1 — первый лист, *_1a — лист продолжения (виджет размножает его и переименовывает поля в L{n}_*).
# A4 вертикально, поля листа как в docx: слева 30, справа 15, сверху/снизу 20 мм; Times New Roman 11 pt.
# Поля AcroForm: шрифты Times (обычный и полужирный), урезанные до cp1251, встроены в /AcroForm/DR.
import json
import os
import sys
import pymupdf as fitz
from fontTools.ttLib import TTFont
from fontTools.agl import AGL2UV
from fontTools import subset

MM = 72 / 25.4
LW = 0.5                                   # толщина линий таблиц, pt (одинарная граница Word)
FS = 11                                    # кегль, pt
X0, XR = 30, 195                           # левая/правая граница текста, мм
W = XR - X0                                # 165 мм
TOP, BOTTOM = 20, 277                      # верх/низ области текста, мм
ROW = 5                                    # высота строки таблицы документов, мм
LINE = 5.1                                 # межстрочный шаг абзацев (14,5 pt), мм

SYS_FONTS = r"C:\Windows\Fonts"
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
FORM_FONTS = {                             # ресурс /DR -> (системный шрифт, урезанная копия для виджета/pdf-lib)
    "TiNR": ("times.ttf", "times-cp1251.ttf"),
    "TiNRB": ("timesbd.ttf", "timesbd-cp1251.ttf"),
}
TEXT_FONTS = {"R": "times.ttf", "B": "timesbd.ttf", "BI": "timesbi.ttf"}
ORG = "ОНТД ООО «ЛАБОРАТОРИЯ МИКРОПРИБОРОВ»"


def norm(widths):
    """Ширины колонок из twips docx -> мм, сумма ровно W."""
    k = W / sum(widths)
    return [w * k for w in widths]


def make_subset(src, dst):
    chars = set()
    for code in range(32, 256):
        try:
            chars.add(ord(bytes([code]).decode("cp1251")))
        except UnicodeDecodeError:
            pass
    opts = subset.Options()
    opts.name_IDs = ["*"]
    opts.name_languages = ["*"]
    opts.notdef_outline = True
    opts.glyph_names = True
    opts.hinting = False
    opts.layout_features = []
    font = subset.load_font(os.path.join(SYS_FONTS, src), opts)
    sub = subset.Subsetter(opts)
    sub.populate(unicodes=sorted(chars))
    sub.subset(font)
    subset.save_font(font, os.path.join(OUT, dst), opts)


def r(x, y, w, h):
    return fitz.Rect(x * MM, y * MM, (x + w) * MM, (y + h) * MM)


class Sheet:
    def __init__(self, doc):
        self.doc = doc
        self.page = doc.new_page(width=210 * MM, height=297 * MM)
        self.measure = {k: fitz.Font(fontfile=os.path.join(SYS_FONTS, f)) for k, f in TEXT_FONTS.items()}
        self.used = set()                  # встраиваем в страницу только реально использованные начертания
        self.sh = self.page.new_shape()
        self.fields = []

    # --- графика
    def hl(self, x1, x2, y):
        self.sh.draw_line((x1 * MM, y * MM), (x2 * MM, y * MM))

    def vl(self, x, y1, y2):
        self.sh.draw_line((x * MM, y1 * MM), (x * MM, y2 * MM))

    def grid(self, x, y, cols, rows):
        """Полная сетка: cols — ширины колонок, rows — высоты строк (мм). Возвращает (xs, ys)."""
        xs = [x]
        for c in cols:
            xs.append(xs[-1] + c)
        ys = [y]
        for h in rows:
            ys.append(ys[-1] + h)
        for yy in ys:
            self.hl(xs[0], xs[-1], yy)
        for xx in xs:
            self.vl(xx, ys[0], ys[-1])
        return xs, ys

    # --- статический текст
    def width(self, text, font="R", size=FS):
        return self.measure[font].text_length(text, fontsize=size) / MM

    def text(self, x, y, text, font="R", size=FS):
        """Строка от левого края x, y — верх строки (мм)."""
        if font not in self.used:
            self.page.insert_font(fontname=font, fontfile=os.path.join(SYS_FONTS, TEXT_FONTS[font]))
            self.used.add(font)
        self.page.insert_text((x * MM, y * MM + size * 0.89), text, fontname=font, fontsize=size)

    def label(self, x, y, w, h, text, font="R", size=FS, fit=False):
        """Надпись в ячейке: по центру по горизонтали и вертикали; \\n — перенос строки.
        fit — уменьшать кегль, пока самая длинная строка не влезет в ширину ячейки."""
        lines = text.split("\n")
        while fit and size > 6 and max(self.width(ln, font, size) for ln in lines) > w - 1:
            size -= 0.5
        lh = size * 1.15 / MM
        y0 = y + (h - lh * len(lines)) / 2
        for i, ln in enumerate(lines):
            self.text(x + (w - self.width(ln, font, size)) / 2, y0 + i * lh + (lh - size / MM) / 2, ln, font, size)

    # --- поля
    def field(self, name, x, y, w, h, size=FS, align=1, multiline=False, bold=False):
        wd = fitz.Widget()
        wd.field_type = fitz.PDF_WIDGET_TYPE_TEXT
        wd.field_name = name
        wd.rect = r(x + 0.4, y + 0.3, w - 0.8, h - 0.6)
        wd.text_font = "Helv"
        wd.text_fontsize = size
        wd.border_width = 0
        wd.fill_color = None
        wd.field_flags = fitz.PDF_TX_FIELD_IS_MULTILINE if multiline else 0
        annot = self.page.add_widget(wd)
        self.fields.append((annot.xref, name, "TiNRB" if bold else "TiNR", size, align))

    def commit(self):
        self.sh.finish(width=LW, color=(0, 0, 0), closePath=False)
        self.sh.commit()


# ---------- общие блоки ----------

def title(s):
    """«Опись № [num]» / НА ВЫДАЧУ ДОКУМЕНТАЦИИ / ИЗ ОНТД ООО «…». Возвращает низ блока, мм."""
    cx = 105
    lab = "Опись № "
    s.text(cx - s.width(lab, "B"), TOP, lab, "B")
    s.field("num", cx, TOP - 0.3, 45, LINE, bold=True, align=0)
    s.label(X0, TOP + LINE, W, LINE, "НА ВЫДАЧУ ДОКУМЕНТАЦИИ")
    a, b = "ИЗ ", ORG
    x = cx - (s.width(a) + s.width(b, "BI")) / 2
    s.text(x, TOP + 2 * LINE, a)
    s.text(x + s.width(a), TOP + 2 * LINE, b, "BI")
    return TOP + 3 * LINE


def addressee(s, y):
    """Таблица «Вид документа / Изделие / Тема | Предприятие / Адрес / Телефон»."""
    cols = norm([1651, 1936, 1828, 3930])
    rows = [9, 17, 5]
    xs, ys = s.grid(X0, y, cols, rows)
    for i, (l1, f1, l2, f2) in enumerate([("Вид документа", "docType", "Предприятие", "company"),
                                          ("Изделие", "product", "Адрес", "address"),
                                          ("Тема", "topic", "Телефон", "phone")]):
        h = rows[i]
        s.label(xs[0], ys[i], cols[0], h, l1)
        s.field(f1, xs[1], ys[i], cols[1], h, multiline=h > ROW)
        s.label(xs[2], ys[i], cols[2], h, l2)
        s.field(f2, xs[3], ys[i], cols[3], h, multiline=h > ROW)
    return ys[-1]


def sheet_header(s):
    """Шапка листа продолжения: «Опись № [num]» слева, «Лист [sheet]» справа."""
    lab = "Опись № "
    s.text(X0, TOP, lab, "B")
    s.field("num", X0 + s.width(lab, "B"), TOP - 0.3, 45, LINE, bold=True, align=0)
    lab = "Лист"
    s.text(XR - 15 - s.width(lab) - 1, TOP, lab)
    s.field("sheet", XR - 15, TOP - 0.3, 15, LINE, align=0)
    return TOP + LINE


def table_head(s, y, cols, head, head_h):
    """Только шапка таблицы документов. Строки (по ROW мм) и блок подписей дорисовывает виджет —
    ровно столько строк, сколько данных; геометрия для него — в opis_layout.json."""
    xs, ys = s.grid(X0, y, cols, [head_h])
    for j, t in enumerate(head):
        s.label(xs[j], y, cols[j], head_h, t, fit=True)
    return ys[-1]


# Блок «Составил/Получил» (рисует виджет сразу под таблицей): отступ после таблицы, затем 2 строки по LINE.
# Координаты x, мм: подпись-надпись X0; линия фамилии от x_name до 88.2; подпись 92.5–120.6; дата 130–152.
SIGN = {"gap": 8, "line": LINE, "underline": 4.3, "name_to": 88.2, "sign": [92.5, 120.6], "date": [130, 152],
        "rows": [["Составил", "compiled"], ["Получил", "received"]]}
SIGN_H = SIGN["gap"] + 2 * LINE
# groupRow — высота строки-заголовка группы (изделие), две строки текста.
LAYOUT = {"mm": MM, "fontSize": FS, "x0": X0, "bottom": BOTTOM, "row": ROW, "groupRow": 2 * ROW, "sign": SIGN}


# ---------- опись копий ----------

COPIES_COLS = norm([6083, 1743, 1559])
COPIES_HEAD = ["Обозначение документа", "Кол-во лис.", "Кол-во экз."]
COPIES_KEYS = ["designation", "sheets", "copies"]


def copies_1(doc):
    s = Sheet(doc)
    y = addressee(s, title(s) + 5)
    y = table_head(s, y + 5, COPIES_COLS, COPIES_HEAD, 10)
    s.commit()
    return s, y


def copies_1a(doc):
    s = Sheet(doc)
    y = table_head(s, sheet_header(s) + 3, COPIES_COLS, COPIES_HEAD, 10)
    s.commit()
    return s, y


# ---------- опись электронных документов ----------

# Ширины в мм (сумма = W). Сигнатура — MD5 (32 hex-символа), поэтому колонка
# широкая; Изм./Расширение («14.pdf») и Кол-во лис./экз. ужаты.
EDOCS_COLS = [19, 50, 16, 63, 8.5, 8.5]
EDOCS_HEAD = ["Инв. №\nподл.", "Обозначение документа", "Изм.\nРасширение\nфайла", "Сигнатура",
              "Кол-\nво\nлис.", "Кол-\nво\nэкз."]
EDOCS_KEYS = ["inv", "designation", "ext", "signature", "sheets", "copies"]


def volume(s, y):
    """Таблица тома: Изделие | Лист/листов; Идентификатор тома | Кол-во томов; Владелец | Занято, МБ."""
    cols = norm([2346, 2210, 2365, 2424])
    rows = [17, ROW, ROW]
    xs, ys = s.grid(X0, y, cols, rows)
    s.label(xs[0], ys[0], cols[0], rows[0], "Изделие")
    s.field("volumeProduct", xs[1], ys[0], cols[1], rows[0], multiline=True)
    s.label(xs[2], ys[0], cols[2], rows[0], "Лист/листов")
    half = cols[3] / 2
    s.label(xs[3], ys[0], cols[3], rows[0], "/")
    s.field("sheet", xs[3], ys[0], half - 1.5, rows[0], align=2)
    s.field("sheets", xs[3] + half + 1.5, ys[0], half - 1.5, rows[0], align=0)
    for i, (l1, f1, l2, f2) in enumerate([("Идентификатор тома", "volumeId", "Кол-во томов", "volumes"),
                                          ("Владелец", "owner", "Занято, МБ", "sizeMb")], start=1):
        s.label(xs[0], ys[i], cols[0], ROW, l1)
        s.field(f1, xs[1], ys[i], cols[1], ROW)
        s.label(xs[2], ys[i], cols[2], ROW, l2)
        s.field(f2, xs[3], ys[i], cols[3], ROW)
    return ys[-1]


def edocs_1(doc):
    s = Sheet(doc)
    y = addressee(s, title(s) + 5)
    y = volume(s, y + 5)
    y = table_head(s, y + 5, EDOCS_COLS, EDOCS_HEAD, 13.5)
    s.commit()
    return s, y


def edocs_1a(doc):
    s = Sheet(doc)
    y = table_head(s, sheet_header(s) + 3, EDOCS_COLS, EDOCS_HEAD, 13.5)
    s.commit()
    return s, y


# ---------- встраивание шрифтов в /AcroForm /DR ----------

def cp1251_encoding_and_widths(tt):
    cmap = tt.getBestCmap()
    upm = tt["head"].unitsPerEm
    hmtx = tt["hmtx"]
    uv2name = {}
    for name, uv in AGL2UV.items():
        afii = name.startswith("afii")
        if uv not in uv2name or (afii and 0x400 <= uv <= 0x4FF):
            uv2name[uv] = name
    diffs, widths = [], []
    for code in range(32, 256):
        try:
            ch = bytes([code]).decode("cp1251")
        except UnicodeDecodeError:
            ch = None
        g = cmap.get(ord(ch)) if ch else None
        widths.append(round(hmtx[g][0] * 1000 / upm) if g else 0)
        if code >= 128 and ch:
            name = uv2name.get(ord(ch)) or f"uni{ord(ch):04X}"
            diffs.append(f"{code} /{name}")
    return " ".join(diffs), widths


def embed_font(doc, path, bold):
    data = open(path, "rb").read()
    tt = TTFont(path)
    head, hhea, os2 = tt["head"], tt["hhea"], tt["OS/2"]
    k = 1000 / head.unitsPerEm
    ps = "".join(c for c in (tt["name"].getDebugName(6) or "") if c.isalnum() or c == "-")
    diffs, widths = cp1251_encoding_and_widths(tt)
    ff = doc.get_new_xref()
    doc.update_object(ff, "<<>>")
    doc.update_stream(ff, data, compress=True)
    doc.xref_set_key(ff, "Length1", str(len(data)))
    fd = doc.get_new_xref()
    bbox = " ".join(str(round(v * k)) for v in (head.xMin, head.yMin, head.xMax, head.yMax))
    doc.update_object(fd, (
        f"<</Type/FontDescriptor/FontName/{ps}/Flags {34 | (1 << 18) if bold else 34}/FontBBox[{bbox}]"
        f"/ItalicAngle {tt['post'].italicAngle}/Ascent {round(hhea.ascent * k)}"
        f"/Descent {round(hhea.descent * k)}/CapHeight {round(getattr(os2, 'sCapHeight', hhea.ascent) * k)}"
        f"/StemV {140 if bold else 80}/FontFile2 {ff} 0 R>>"))
    font = doc.get_new_xref()
    doc.update_object(font, (
        f"<</Type/Font/Subtype/TrueType/BaseFont/{ps}/FirstChar 32/LastChar 255"
        f"/Widths[{' '.join(map(str, widths))}]/FontDescriptor {fd} 0 R"
        f"/Encoding<</Type/Encoding/BaseEncoding/WinAnsiEncoding/Differences[{diffs}]>>>>"))
    return font


def embed_form_fonts(doc, fields):
    cat = doc.pdf_catalog()
    for res, (_, sub) in FORM_FONTS.items():
        xref = embed_font(doc, os.path.join(OUT, sub), bold=res.endswith("B"))
        doc.xref_set_key(cat, f"AcroForm/DR/Font/{res}", f"{xref} 0 R")
    doc.xref_set_key(cat, "AcroForm/DA", f"(/TiNR {FS} Tf 0 g)")
    doc.xref_set_key(cat, "AcroForm/NeedAppearances", "true")
    for xref, name, res, size, align in fields:
        doc.xref_set_key(xref, "DA", f"(/{res} {size} Tf 0 g)")
        doc.xref_set_key(xref, "Q", str(align))
        doc.xref_set_key(xref, "AP", "null")      # отображение строит просмотрщик (NeedAppearances)


def build(maker, name):
    path = os.path.join(OUT, name)
    doc = fitz.open()
    s, top = maker(doc)
    doc.subset_fonts()
    doc.save(path)                                # фиксируем виджеты
    doc = fitz.open(path)
    embed_form_fonts(doc, s.fields)
    doc.save(path + ".tmp", garbage=3, deflate=True)
    doc.close()
    os.replace(path + ".tmp", path)
    print(f"{name}: строки таблицы с {top:.1f} мм, полей {len(s.fields)}, {os.path.getsize(path)} байт")
    return round(top, 2)


if __name__ == "__main__":
    for src, dst in FORM_FONTS.values():
        make_subset(src, dst)
    for kind, cols, keys, (f1, fa) in [("copies", COPIES_COLS, COPIES_KEYS, (copies_1, copies_1a)),
                                       ("edocs", EDOCS_COLS, EDOCS_KEYS, (edocs_1, edocs_1a))]:
        LAYOUT[kind] = {"cols": [round(c, 2) for c in cols], "keys": keys,
                        "top1": build(f1, f"opis_{kind}_1.pdf"), "topA": build(fa, f"opis_{kind}_1a.pdf")}
    with open(os.path.join(OUT, "opis_layout.json"), "w", encoding="utf-8") as f:
        json.dump(LAYOUT, f, ensure_ascii=False, indent=1)
