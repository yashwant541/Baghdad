"""A tiny, dependency-free PDF writer (landscape/portrait A4, Helvetica text, filled rectangles, wrapped
table cells, page numbers). Written because a Dataiku code environment cannot be assumed to have reportlab.

Text is drawn with the standard Helvetica fonts, so only Latin (Windows-1252) characters can be shown; any
other character (e.g. Arabic) is drawn as '?'. Python 3.9 compatible.
"""
import zlib

# Helvetica advance widths (1/1000 em) for ASCII 32..126
_W = [278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278,
      556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556,
      1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778,
      667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556,
      333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556,
      556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584]

_REPLACE = {"→": "->", "←": "<-", "↔": "<->", "≤": "<=", "≥": ">=", "−": "-", "‑": "-",
            "‐": "-", " ": " ", "✓": "v", "⚠": "!", "Δ": "Delta ", "≈": "~", "…": "..."}


def clean(text):
    s = "" if text is None else str(text)
    for k, v in _REPLACE.items():
        s = s.replace(k, v)
    out = []
    for ch in s:
        try:
            ch.encode("cp1252")
            out.append(ch)
        except UnicodeEncodeError:
            out.append("?")
    return "".join(out)


def text_width(s, size, bold=False):
    w = 0
    for ch in s:
        o = ord(ch)
        w += _W[o - 32] if 32 <= o <= 126 else 556
    return w * size / 1000.0 * (1.06 if bold else 1.0)


def wrap(text, width, size, bold=False):
    """Greedy word wrap; very long tokens are split."""
    lines = []
    for para in clean(text).split("\n"):
        cur = ""
        for word in para.split(" "):
            while text_width(word, size, bold) > width and len(word) > 1:      # token wider than the cell
                k = len(word)
                while k > 1 and text_width(word[:k], size, bold) > width:
                    k -= 1
                if cur:
                    lines.append(cur)
                    cur = ""
                lines.append(word[:k])
                word = word[k:]
            trial = (cur + " " + word) if cur else word
            if text_width(trial, size, bold) <= width or not cur:
                cur = trial
            else:
                lines.append(cur)
                cur = word
        lines.append(cur)
    return lines or [""]


def _esc(s):
    return s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


class SimplePDF(object):
    def __init__(self, landscape=True, margin=28, title="", footer=""):
        self.w, self.h = (842.0, 595.0) if landscape else (595.0, 842.0)
        self.m = margin
        self.title = clean(title)
        self.footer = clean(footer)
        self.pages = []
        self.cur = None
        self.y = 0
        self.new_page()

    # ---- page handling
    def new_page(self):
        self.cur = []
        self.pages.append(self.cur)
        self.y = self.h - self.m
        return self.y

    @property
    def bottom(self):
        return self.m + 14

    def room(self, need):
        return self.y - need >= self.bottom

    def ensure(self, need):
        if not self.room(need):
            self.new_page()

    # ---- drawing primitives (y measured from the bottom, PDF style)
    def text(self, x, y, s, size=8, bold=False, color=(0, 0, 0)):
        self.cur.append("BT /%s %.2f Tf %.3f %.3f %.3f rg %.2f %.2f Td (%s) Tj ET" % (
            "F2" if bold else "F1", size, color[0], color[1], color[2], x, y, _esc(clean(s))))

    def rect(self, x, y, w, h, fill=None, stroke=None, lw=0.4):
        ops = []
        if fill:
            ops.append("%.3f %.3f %.3f rg" % fill)
        if stroke:
            ops.append("%.3f %.3f %.3f RG %.2f w" % (stroke + (lw,)))
        ops.append("%.2f %.2f %.2f %.2f re %s" % (x, y, w, h, "B" if (fill and stroke) else ("f" if fill else "S")))
        self.cur.append(" ".join(ops))

    def line(self, x1, y1, x2, y2, color=(0.6, 0.6, 0.6), lw=0.5):
        self.cur.append("%.3f %.3f %.3f RG %.2f w %.2f %.2f m %.2f %.2f l S" % (color + (lw, x1, y1, x2, y2)))

    # ---- text blocks
    def paragraph(self, s, size=9, bold=False, color=(0, 0, 0), indent=0, gap=3):
        width = self.w - 2 * self.m - indent
        lead = size * 1.25
        for ln in wrap(s, width, size, bold):
            self.ensure(lead)
            self.y -= lead
            self.text(self.m + indent, self.y, ln, size, bold, color)
        self.y -= gap

    def heading(self, s, size=13, color=(0.12, 0.22, 0.42)):
        self.ensure(size * 2.4)
        self.y -= size * 0.6
        self.paragraph(s, size, True, color, gap=2)
        self.line(self.m, self.y + 1, self.w - self.m, self.y + 1, color=color, lw=0.8)
        self.y -= 4

    # ---- tables
    def table(self, headers, rows, widths, size=7, head_fill=(0.12, 0.22, 0.42), aligns=None, fills=None, bold_rows=None):
        """widths are relative; scaled to the page width. rows: list of lists of str. fills: optional list (per row) of
        dict col_index -> rgb. Repeats the header on every new page."""
        total = float(sum(widths))
        avail = self.w - 2 * self.m
        cw = [avail * x / total for x in widths]
        aligns = aligns or ["l"] * len(headers)
        pad = 2.5
        lead = size * 1.28

        def draw_header():
            hl = [wrap(h, cw[i] - 2 * pad, size, True) for i, h in enumerate(headers)]
            hh = max(len(x) for x in hl) * lead + 2 * pad
            self.ensure(hh + lead * 2)
            x = self.m
            for i in range(len(headers)):
                self.rect(x, self.y - hh, cw[i], hh, fill=head_fill)
                for k, ln in enumerate(hl[i]):
                    self.text(x + pad, self.y - pad - (k + 1) * lead + 2, ln, size, True, (1, 1, 1))
                x += cw[i]
            self.y -= hh

        draw_header()
        for r, row in enumerate(rows):
            cells = [wrap(str(c), cw[i] - 2 * pad, size, bool(bold_rows and r in bold_rows)) for i, c in enumerate(row)]
            rh = max(len(c) for c in cells) * lead + 2 * pad
            if not self.room(rh):
                self.new_page()
                draw_header()
            x = self.m
            top = self.y
            rowfill = (fills[r] if fills else None) or {}
            for i in range(len(headers)):
                bg = rowfill.get(i) or ((0.97, 0.97, 0.99) if r % 2 else None)
                self.rect(x, top - rh, cw[i], rh, fill=bg, stroke=(0.82, 0.83, 0.88), lw=0.3)
                for k, ln in enumerate(cells[i]):
                    ty = top - pad - (k + 1) * lead + 2
                    if aligns[i] == "r":
                        tx = x + cw[i] - pad - text_width(ln, size, bool(bold_rows and r in bold_rows))
                    else:
                        tx = x + pad
                    self.text(tx, ty, ln, size, bool(bold_rows and r in bold_rows))
                x += cw[i]
            self.y -= rh
        self.y -= 6

    # ---- output
    def output(self):
        objs = []                                           # 1-indexed pdf objects as bytes

        def add(b):
            objs.append(b)
            return len(objs)

        add(b"")                                            # 1 catalog (filled later)
        add(b"")                                            # 2 pages  (filled later)
        f1 = add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
        f2 = add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>")
        kids = []
        n = len(self.pages)
        for i, ops in enumerate(self.pages, 1):
            foot = "%s   Page %d of %d" % (self.footer, i, n)
            extra = ["BT /F1 7 Tf 0.45 0.45 0.5 rg %.2f %.2f Td (%s) Tj ET" % (self.m, self.m - 4, _esc(clean(self.title))),
                     "BT /F1 7 Tf 0.45 0.45 0.5 rg %.2f %.2f Td (%s) Tj ET" % (
                         self.w - self.m - text_width(clean(foot), 7), self.m - 4, _esc(clean(foot)))]
            stream = "\n".join(ops + extra).encode("cp1252", "replace")
            comp = zlib.compress(stream)
            cid = add(b"<< /Length %d /Filter /FlateDecode >>\nstream\n" % len(comp) + comp + b"\nendstream")
            pid = add(("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 %.0f %.0f] /Contents %d 0 R "
                       "/Resources << /Font << /F1 %d 0 R /F2 %d 0 R >> >> >>" % (self.w, self.h, cid, f1, f2)).encode())
            kids.append(pid)
        objs[0] = b"<< /Type /Catalog /Pages 2 0 R >>"
        objs[1] = ("<< /Type /Pages /Kids [%s] /Count %d >>" % (" ".join("%d 0 R" % k for k in kids), len(kids))).encode()
        out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = []
        for i, body in enumerate(objs, 1):
            offsets.append(len(out))
            out += b"%d 0 obj\n" % i + body + b"\nendobj\n"
        xref = len(out)
        out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
        for off in offsets:
            out += b"%010d 00000 n \n" % off
        out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, xref)
        return bytes(out)
