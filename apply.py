#!/usr/bin/env python3
"""
apply.py  -  Excel budget estimates to A4 landscape PDF
=======================================================
Source : INPUT/FINAL Budget Anuman 31 October 2026.xlsx  (sheet "Details")
Output : OUTPUT/budget_paged.pdf

Column layout (A-R = indices 0-17):
  Visible : A B C D E F G H I J K L R  (indices 0-11 and 17)
  Hidden  : M N O P Q                  (indices 12-16, OLD year cols)

Font encoding:
  Unicode Devanagari  - column headers + title rows (contains Unicode U+0900+)
  KrutiDev010 legacy  - work names, remarks, signatures; converted to Unicode
                        Devanagari via krutidev_to_unicode() before rendering
  Numbers / Latin     - numeric columns, work-IDs, English sub-headers

One PDF page per signature block. PageBreak after each 3-line cluster:
  SIG1 vfuy (Anil)  SIG2 vf/k (Exec.Eng)  SIG3 lk-fu-fo- (PWD)

Usage
-----
    python apply.py                   # render OUTPUT/budget_paged.pdf
    python apply.py --report          # dry-run: print segment list
    python apply.py --sheet Details   # pick worksheet (default: Details)
    python apply.py --keep-hidden     # include hidden cols M-Q
    python apply.py -i file -o out.pdf
"""
from __future__ import annotations
import argparse, datetime as dt, re, sys
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit('pip install openpyxl')

# xlrd is needed for legacy .xls files
try:
    import xlrd as _xlrd          # type: ignore
    _HAVE_XLRD = True
except ImportError:
    _HAVE_XLRD = False

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.pdfgen import canvas as _rl_canvas
    from reportlab.platypus import (
        BaseDocTemplate, Frame, LongTable, PageBreak,
        PageTemplate, Paragraph, Table, TableStyle,
    )
except ImportError:
    sys.exit('pip install reportlab')

# ---------------------------------------------------------------------------
# Regex classifiers
# ---------------------------------------------------------------------------
TITLE_RE   = re.compile(r'\u092c\u091c\u091f \u0905\u0928\u0941\u092e\u093e\u0928')  # बजट अनुमान
HEAD_RE    = re.compile(r'\u00d8- la')           # Ø- la column-header row
SIG1_RE    = re.compile(r'vfuy')                 # vfuy f[kaph (Anil Khinchi)
SIG2_RE    = re.compile(r'vf/k')                 # vf/k'kk vfHk;Urk (Exec Eng)
SIG3_RE    = re.compile(r'lk-fu-fo-')            # lk-fu-fo- (PWD district)
TOTAL_RE   = re.compile(r'^\s*(?:;ksx|\u092f\u094b\u0917|dqy|\u0915\u0941\u0932)')
NUMERIC_RE = re.compile(r'^-?[\d,]+(\.\d+)?%?$')
DEVA_RE    = re.compile(r'[\u0900-\u097F]')     # Unicode Devanagari
LOWER_RE   = re.compile(r'[a-z]')               # lowercase Latin = KrutiDev signal
SUBHD_RE   = re.compile(r'^(?:Budget Head|eq\[;|mi en|y\?kq|foRrh;|lM\+d)')  # sub-headers

# ---------------------------------------------------------------------------
# Colour palette
# ---------------------------------------------------------------------------
C_TITLE_BG = colors.HexColor("#E8E2D5")
C_HEAD_BG  = colors.HexColor("#F2EDE4")
C_ALT_BG   = colors.HexColor("#FAFAF7")
C_TOTAL_BG = colors.HexColor("#EAE4D6")
C_RULE     = colors.HexColor("#8A8377")
C_GRID     = colors.HexColor("#C9C3B6")
C_INK      = colors.HexColor("#111111")
C_MUTE     = colors.HexColor("#6B655B")

F_DEVA  = "Deva"
F_LATIN = "Helvetica"

# Visible: A-L (0-11) + R (17) = 13 cols.  Hidden: M-Q (12-16).
VISIBLE = list(range(12)) + [17]
ALL18   = list(range(18))

#               A     B      C     D      E     F     G     H     I     J      K      L      R
PROPS13 = [2.5, 26.0,  5.5,  5.5, 14.0,  6.5,  7.0,  6.0,  6.0,  7.0,  7.0,  7.0, 20.0]
#               A     B      C     D      E     F     G     H     I     J      K      L    M    N    O    P    Q     R
PROPS18 = [2.5, 22.0,  5.0,  5.0, 12.0,  5.5,  6.0,  5.0,  5.0,  6.0,  6.0,  6.0, 4.5, 4.5, 4.5, 4.5, 4.5, 15.0]


def is_kruti(text: str) -> bool:
    '''KrutiDev text: has lowercase Latin letters, no Unicode Devanagari.'''
    s = str(text).strip()
    if not s or DEVA_RE.search(s):
        return False
    return bool(LOWER_RE.search(s))


# Words/labels that are genuine English (not KrutiDev) and must not be mangled.
ENGLISH_WORDS = {
    # common function words
    'a', 'an', 'the', 'and', 'or', 'of', 'to', 'in', 'on', 'at', 'by', 'for',
    'from', 'with', 'as', 'is', 'are', 'be', 'was', 'were', 'this', 'that',
    'these', 'those', 'it', 'its', 'not', 'no', 'up', 'out', 'off', 'per',
    'via', 'new', 'old', 'near', 'upto', 'under', 'over', 'left', 'right',
    # civic / budget / works vocabulary
    'budget', 'head', 'sub', 'subhead', 'major', 'minor', 'total', 'grand',
    'name', 'work', 'works', 'id', 'sr', 'sl', 'scheme', 'amount', 'year',
    'yr', 'yrs', 'date', 'dated', 'dt', 'page', 'note', 'remarks', 'misc',
    'road', 'roads', 'bridge', 'bridges', 'building', 'buildings',
    'construction', 'repair', 'repairs', 'restoration', 'restoriton',
    'modernisation', 'modernization', 'development', 'devlopment',
    'improvement', 'widening', 'strengthening', 'link', 'links', 'district',
    'state', 'national', 'highway', 'special', 'assistance', 'permanent',
    'announced', 'bindu', 'km', 'cr', 'crore', 'lakh', 'file', 'part',
    'section', 'internal', 'path', 'atal', 'pragati', 'sangath', 'sayra',
    'singhada', 'pipad', 'jaipur', 'thal', 'gogunda', 'udaipur', 'khel',
    'stadium', 'jaswantgarh', 'sec', 'gen', 'sc', 'st', 'vp', 'pwd', 'rcc',
    'np', 'ml', 'fd',
}

# Characters that occur in KrutiDev-encoded text but essentially never in
# ordinary English prose.  Their presence is a reliable KrutiDev signal and
# prevents English cells from being converted (and mangled) by mistake.
KRUTI_MARKERS = set(
    '[]{}|^~`@;'
    '\u00a1\u00ab\u00ba\u00bc\u00bd\u00be\u00bf'
    '\u00c2\u00c3\u00c4\u00c5\u00c6\u00c7\u00c9\u00cb\u00cc\u00cd\u00ce'
    '\u00cf\u00d1\u00d2\u00d3\u00d4\u00d6\u00d8\u00d9\u00da\u00db\u00dc'
    '\u00dd\u00de\u00df\u00e0\u00e1\u00e2\u00e3\u00e5\u00e6\u00e7\u00e8'
    '\u00e9\u00ea\u00eb\u00ec\u00ed\u00ee\u00ef\u00f1\u00f2\u00f4\u00f5'
    '\u00f6\u00f7\u00f9\u00fa\u00fb\u00fc\u00fd\u00fe\u00ff'
    '\u00b2\u00b3\u00b4\u00b5\u00b6\u00b7\u00b8\u00b9'
    '\u201c\u201d'
)


def is_english(s: str) -> bool:
    '''Heuristic: does this ASCII string read as ordinary English text?

    A string is treated as English when it contains none of the KrutiDev-only
    marker characters AND either most of its words are known English words /
    short acronyms, or it has an English-like vowel density.  KrutiDev runs
    such as ``lM+d`` / ``eq[;`` / ``Hkou`` fail all three tests and are
    therefore converted to Unicode Devanagari.
    '''
    s = str(s)
    if any(ch in KRUTI_MARKERS for ch in s):
        return False
    tokens = re.findall(r'[A-Za-z]{2,}', s)
    if tokens:
        known = sum(1 for t in tokens
                    if t.lower() in ENGLISH_WORDS
                    or (t.isupper() and len(t) <= 5))
        if known / len(tokens) >= 0.5:
            return True
    letters = [c for c in s if c.isalpha()]
    if letters:
        vowels = sum(1 for c in letters if c.lower() in 'aeiou')
        if vowels / len(letters) >= 0.32:
            return True
    return False


# ---------------------------------------------------------------------------
# KrutiDev (legacy Remington/Walkman-Chanakya) -> Unicode Devanagari
# Ported from the public Walkman Chanakya905 converter (MIT).
# ---------------------------------------------------------------------------
def krutidev_to_unicode(src: str) -> str:
    '''Convert KrutiDev legacy-encoded text to Unicode Devanagari.'''
    if not src:
        return ''
    out = str(src)
    array_one = [
        "ñ","Q+Z","sas","aa",")Z","ZZ","\u2018","\u2019","\u201c","\u201d",
        "å",  "ƒ",  "„",   "…",   "†",   "‡",   "ˆ",   "‰",   "Š",   "‹",
        "¶+",   "d+", "[+k","[+", "x+",  "T+",  "t+", "M+", "<+", "Q+", ";+", "j+", "u+",
        "Ùk", "Ù", "Dr", "–", "—","é","™","=kk","f=k",
        "à",   "á",    "â",   "ã",   "ºz",  "º",   "í", "{k", "{", "=",  "«",
        "Nî",   "Vî",    "Bî",   "Mî",   "<î", "|", "K", "}",
        "J",   "Vª",   "Mª",  "<ªª",  "Nª",   "Ø",  "Ý", "nzZ",  "æ", "ç", "Á", "xz", "#", ":",
        "v‚","vks",  "vkS",  "vk",    "v",  "b±", "Ã",  "bZ",  "b",  "m",  "Å",  ",s",  ",",   "_",
        "ô",  "d", "Dk", "D", "[k", "[", "x","Xk", "X", "Ä", "?k", "?",   "³",
        "pkS",  "p", "Pk", "P",  "N",  "t", "Tk", "T",  ">", "÷", "¥",
        "ê",  "ë",   "V",  "B",   "ì",   "ï", "M+", "<+", "M",  "<", ".k", ".",
        "r",  "Rk", "R",   "Fk", "F",  ")", "n", "/k", "èk",  "/", "Ë", "è", "u", "Uk", "U",
        "i",  "Ik", "I",   "Q",    "¶",  "c", "Ck",  "C",  "Hk",  "H", "e", "Ek",  "E",
        ";",  "¸",   "j",    "y", "Yk",  "Y",  "G",  "o", "Ok", "O",
        "'k", "'",   "\"k",  "\"",  "l", "Lk",  "L",   "g",
        "È", "z",
        "Ì", "Í", "Î",  "Ï",  "Ñ",  "Ò",  "Ó",  "Ô",   "Ö",  "Ø",  "Ù","Ük", "Ü",
        "‚",    "ks",   "kS",   "k",  "h",    "q",   "w",   "`",    "s",    "S",
        "a",    "¡",    "%",     "W",  "•", "·", "∙", "·", "~j",  "~", "\\","+"," ः",
        "^", "*",  "Þ", "ß", "(", "¼", "½", "¿", "À", "¾", "A", "-", "&", "&", "Œ", "]","~ ","@"
    ]
    array_two = [
        "॰","QZ+","sa","a","र्द्ध","Z","\"","\"","'","'",
        "०",  "१",  "२",  "३",     "४",   "५",  "६",   "७",   "८",   "९",
        "फ़्",  "क़",  "ख़", "ख़्",  "ग़", "ज़्", "ज़",  "ड़",  "ढ़",   "फ़",  "य़",  "ऱ",  "ऩ",
        "त्त", "त्त्", "क्त",  "दृ",  "कृ","nn","nnd","=k","f=",
        "ह्न",  "ह्य",  "हृ",  "ह्म",  "ह्र",  "ह्",   "द्द",  "क्ष", "क्ष्", "त्र", "त्र्",
        "छ्य",  "ट्य",  "ठ्य",  "ड्य",  "ढ्य", "द्य", "ज्ञ", "द्व",
        "श्र",  "ट्र",    "ड्र",    "ढ्र",    "छ्र",   "क्र",  "फ्र", "र्द्र",  "द्र",   "प्र", "प्र",  "ग्र", "रु",  "रू",
        "ऑ",   "ओ",  "औ",  "आ",   "अ", "ईं", "ई",  "ई",   "इ",  "उ",   "ऊ",  "ऐ",  "ए", "ऋ",
        "क्क", "क", "क", "क्", "ख", "ख्", "ग", "ग", "ग्", "घ", "घ", "घ्", "ङ",
        "चै",  "च", "च", "च्", "छ", "ज", "ज", "ज्",  "झ",  "झ्", "ञ",
        "ट्ट",   "ट्ठ",   "ट",   "ठ",   "ड्ड",   "ड्ढ",  "ड़", "ढ़", "ड",   "ढ", "ण", "ण्",
        "त", "त", "त्", "थ", "थ्",  "द्ध",  "द", "ध", "ध", "ध्", "ध्", "ध्", "न", "न", "न्",
        "प", "प", "प्",  "फ", "फ्",  "ब", "ब", "ब्",  "भ", "भ्",  "म",  "म", "म्",
        "य", "य्",  "र", "ल", "ल", "ल्",  "ळ",  "व", "व", "व्",
        "श", "श्",  "ष", "ष्", "स", "स", "स्", "ह",
        "ीं", "्र",
        "द्द", "ट्ट","ट्ठ","ड्ड","कृ","भ","्य","ड्ढ","झ्","क्र","त्त्","श","श्",
        "ॉ",  "ो",   "ौ",   "ा",   "ी",   "ु",   "ू",   "ृ",   "े",   "ै",
        "ं",   "ँ",   "ः",   "ॅ",  "ऽ", "ऽ", "ऽ", "ऽ", "्र",  "्", "?", "़",":",
        "\u2018",   "\u2019",   "\u201c",   "\u201d",  ";",  "(",    ")",   "{",    "}",   "=", "।", ".", "-",  "µ", "॰", ",","् ","/"
    ]
    # Move 'f' (short-i placeholder) to its correct position, then convert to 'ि'.
    # KrutiDev types the i-matra *before* its consonant, Unicode stores it after,
    # so swap 'f' with the consonant that follows.  A consonant may be a two-char
    # key whose first char alone is the halant form ("Hk"=भ vs "H"=भ्,
    # "Dk"=क vs "D"=क्, ...); swap past both so the pair stays intact.
    HALANT_FIRST = set("D[X?PT÷RF/UICHE¸YO'\"L.Ëè")
    out = '  ' + out + '  '
    position_of_f = out.rfind('f')
    while position_of_f != -1:
        n = 1
        if (position_of_f + 2 < len(out)
                and out[position_of_f + 1] in HALANT_FIRST
                and out[position_of_f + 2] == 'k'):
            n = 2
        out = (out[:position_of_f]
               + out[position_of_f + 1:position_of_f + 1 + n]
               + 'f' + out[position_of_f + 1 + n:])
        position_of_f = out.rfind('f', 0, position_of_f)
    out = out.replace('f', 'ि').strip()

    # Move half-R ('Z') correctly relative to matras
    out = '  ' + out + '  '
    set_of_matras = [
        "‚", "ks", "kS", "k", "h", "q", "w", "`", "s", "S",
        "a", "¡", "%", "W", "·", "∙", "·", "~j", "~", "\\", "+", " ः",
    ]
    position_of_r = out.find('Z')
    while position_of_r != -1:
        out = out.replace('Z', '', 1)
        prev_char = out[position_of_r - 1]
        if prev_char and prev_char in set_of_matras:
            out = out[:position_of_r - 2] + 'र्' + out[position_of_r - 2:]
        else:
            out = out[:position_of_r - 1] + 'र्' + out[position_of_r - 1:]
        position_of_r = out.find('Z')
    out = out.strip()

    # ASCII -> Unicode replacements (order matters)
    for i, key in enumerate(array_one):
        if key:
            out = out.replace(key, array_two[i])

    # Final cleanups for 'nn' / 'nnd'
    out = out.replace('nn', 'न्न').replace('nnd', 'न्न्')
    return out


def register_fonts(font_dir: Path) -> dict:
    """Register the Devanagari TTF.

    KrutiDev text is converted to Unicode by ``krutidev_to_unicode()`` before
    rendering, so no legacy KrutiDev font is required.
    """
    mapping = {F_DEVA: F_LATIN}
    for name in ('NotoSansDevanagari-Regular.ttf', 'NotoSansDevanagari.ttf'):
        p = font_dir / name
        if not p.exists():
            continue
        try:
            pdfmetrics.registerFont(TTFont(F_DEVA, str(p)))
            mapping[F_DEVA] = F_DEVA
            break
        except Exception as exc:
            print(f'[warn] {p}: {exc}', file=sys.stderr)
    if mapping[F_DEVA] == F_LATIN:
        print(
            f'[info] No NotoSansDevanagari TTF in {font_dir} -- '
            'place NotoSansDevanagari-Regular.ttf in fonts/',
            file=sys.stderr,
        )
    return mapping


def make_styles(mapping: dict) -> dict:
    base = dict(textColor=C_INK, splitLongWords=1, spaceBefore=0, spaceAfter=0)
    d = mapping[F_DEVA]
    return {
        # 'lat' = plain Latin fallback (numbers, work-IDs, English sub-headers)
        'lat':    ParagraphStyle('lt', fontSize=5.8, leading=7.6,
                                  alignment=TA_LEFT,   fontName=F_LATIN, **base),
        'num':    ParagraphStyle('n',  fontSize=5.8, leading=7.6,
                                  alignment=TA_RIGHT,  fontName=F_LATIN, **base),
        'no':     ParagraphStyle('no', fontSize=5.8, leading=7.6,
                                  alignment=TA_CENTER, fontName=F_LATIN,
                                  textColor=C_MUTE, spaceBefore=0, spaceAfter=0),
        # 'deva' = Unicode Devanagari (headers, titles, converted KrutiDev text)
        'deva':   ParagraphStyle('dv', fontSize=5.8, leading=7.6,
                                  alignment=TA_LEFT,   fontName=d, **base),
        'title':  ParagraphStyle('t',  fontSize=9.5, leading=12.0,
                                  alignment=TA_CENTER, fontName=d,
                                  textColor=C_INK, spaceBefore=4, spaceAfter=3),
        'colhdr': ParagraphStyle('h',  fontSize=5.5, leading=7.0,
                                  alignment=TA_LEFT,   fontName=d,
                                  textColor=C_MUTE, spaceBefore=0, spaceAfter=0),
        'sig':    ParagraphStyle('sg', fontSize=6.2, leading=8.2,
                                  alignment=TA_CENTER, fontName=F_LATIN, **base),
        'sigdeva':ParagraphStyle('sgd', fontSize=6.2, leading=8.2,
                                  alignment=TA_CENTER, fontName=d, **base),
    }


def pick_style(text: str, styles: dict):
    """Return ``(text_to_render, ParagraphStyle)``.

    KrutiDev legacy cells are converted to Unicode Devanagari so they render
    with the Devanagari font; genuine English/numeric cells pass through.
    """
    s = str(text).strip()
    if not s:
        return s, styles['lat']
    if DEVA_RE.search(s):
        return s, styles['deva']
    if NUMERIC_RE.match(s):
        return s, styles['num']
    if is_kruti(s) and not is_english(s):
        return krutidev_to_unicode(s), styles['deva']
    return s, styles['lat']


def fmt(v) -> str:
    if v is None: return ''
    s = str(v).replace('\n', ' ').strip()
    if isinstance(v, float):
        if abs(v) < 1e-9: return '0'
        return f'{v:,.0f}' if abs(v - round(v)) < 1e-6 else f'{v:,.2f}'
    try:
        fv = float(s.replace(',', ''))
        if abs(fv) < 1e-9: return '0'
        return f'{fv:,.0f}' if abs(fv - round(fv)) < 1e-6 else f'{fv:,.2f}'
    except ValueError:
        pass
    return s


def esc(t: str) -> str:
    return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def blank(row) -> bool:
    return not any(str(c).strip() for c in row)


def list_sheets(path: Path) -> list:
    """Return the worksheet names of *path* (xls or xlsx)."""
    if path.suffix.lower() == '.xls':
        if not _HAVE_XLRD:
            sys.exit('[error] Reading .xls requires xlrd (pip install xlrd==1.2.0)')
        return list(_xlrd.open_workbook(str(path)).sheet_names())
    wb = load_workbook(str(path), read_only=True)
    names = list(wb.sheetnames)
    wb.close()
    return names


def load_rows(path: Path, sheet: str | None = None,
              keep_hidden: bool = False, native: bool = False) -> list:
    """Read a worksheet as a list of rows.

    ``native=False`` (default) applies the *Details* layout: pad/trim to 18
    columns and keep only the visible ones (A-L + R).  ``native=True`` keeps
    every column the sheet actually uses, trimming all-empty edge columns --
    used for the auxiliary sheets (Cover, Saransh, Letter, ...).
    """
    suffix = path.suffix.lower()

    if suffix == '.xls':
        # Legacy Excel 97-2003 format — needs xlrd
        if not _HAVE_XLRD:
            sys.exit(
                '[error] Reading .xls requires xlrd.\n'
                '        Run:  pip install xlrd==1.2.0\n'
                '        Or open the file in Excel and Save As .xlsx'
            )
        wb = _xlrd.open_workbook(str(path))
        names = wb.sheet_names()
        if sheet:
            if sheet not in names:
                sys.exit(f'[error] Sheet "{sheet}" not found. '
                         f'Available: {", ".join(names)}')
            ws = wb.sheet_by_name(sheet)
        else:
            ws = wb.sheet_by_index(0)
        raw = []
        for ri in range(ws.nrows):
            row = []
            for ci in range(ws.ncols):
                cell = ws.cell(ri, ci)
                # xlrd type 0=empty 1=text 2=number 3=date 4=bool 5=error
                if cell.ctype == 0:
                    row.append('')
                elif cell.ctype == 2:
                    v = cell.value
                    row.append(int(v) if v == int(v) else v)
                else:
                    row.append(cell.value)
            raw.append(row)
    else:
        # xlsx / xlsm / xlam
        wb = load_workbook(str(path), data_only=True, read_only=True)
        names = wb.sheetnames
        if sheet:
            if sheet not in names:
                wb.close()
                sys.exit(f'[error] Sheet "{sheet}" not found. '
                         f'Available: {", ".join(names)}')
            ws = wb[sheet]
        else:
            ws = wb[names[0]]
        raw = [
            [('' if v is None else v) for v in row]
            for row in ws.iter_rows(values_only=True)
        ]
        wb.close()

    if native:
        ncols = max((len(r) for r in raw), default=0)
        rows = [(r + [''] * (ncols - len(r)))[:ncols] for r in raw]
        used = [i for i in range(ncols) if any(str(r[i]).strip() for r in rows)]
        if used:
            lo, hi = used[0], used[-1]
            rows = [r[lo:hi + 1] for r in rows]
        return rows

    rows = [(r + [''] * max(0, 18 - len(r)))[:18] for r in raw]
    cols = ALL18 if keep_hidden else VISIBLE
    return [[r[i] for i in cols] for r in rows]


def segment(rows: list) -> list:
    """
    Split rows into one segment per budget block.

    A segment OPENS  at a title row (bजट अनुमान).
    A segment CLOSES immediately after the 3rd signature line (lk-fu-fo-).
    Blank rows between SIG3 and the next title are discarded.
    Double-blank runs inside a segment are collapsed to a single blank.
    """
    SIG3 = re.compile(r'lk-fu-fo-')

    segs: list = []
    cur:  list = []
    after_sig3 = False          # flag: we have seen SIG3, skip blanks

    for row in rows:
        joined = ' '.join(str(c) for c in row if c)
        is_title = TITLE_RE.search(joined) and not HEAD_RE.search(joined)
        is_blank = not any(str(c).strip() for c in row)

        # A new title row starts a new segment
        if is_title:
            if cur:
                segs.append(cur)
            cur = [row]
            after_sig3 = False
            continue

        # After SIG3 we skip all blank rows until the next title
        if after_sig3:
            if is_blank:
                continue      # discard blank between SIG3 and next title
            # Non-blank, non-title content after SIG3 - shouldn't normally
            # happen, but include it in current segment
            after_sig3 = False

        cur.append(row)

        # Detect SIG3 and set the flag
        if SIG3.search(joined):
            after_sig3 = True

    if cur:
        segs.append(cur)

    # Post-process: trim leading blanks, collapse >1 consecutive blank rows
    out = []
    for seg in segs:
        while seg and not any(str(c).strip() for c in seg[0]):
            seg.pop(0)
        if not seg:
            continue
        squeezed, run = [], 0
        for row in seg:
            if not any(str(c).strip() for c in row):
                run += 1
                if run > 1:
                    continue    # drop extra blank rows
            else:
                run = 0
            squeezed.append(row)
        # Also trim trailing blanks (between last data and SIG3 there may be none,
        # but be safe)
        while squeezed and not any(str(c).strip() for c in squeezed[-1]):
            squeezed.pop()
        if squeezed:
            out.append(squeezed)

    return out

def col_widths(ncols: int, usable: float) -> list:
    props = PROPS13 if ncols == 13 else PROPS18
    if len(props) != ncols: props = [1.0] * ncols
    t = sum(props)
    return [p / t * usable for p in props]


def build_table(seg, cw, styles, mapping) -> Table:
    ncols = len(cw)
    data, cmds, header_end = [], [], 0
    first_sig = next(
        (i for i, row in enumerate(seg)
         if SIG1_RE.search(' '.join(str(c) for c in row if c))),
        -1,
    )
    L0, L1 = 1, max(1, ncols // 2 - 1)
    R0, R1 = ncols // 2, ncols - 1

    for r, row in enumerate(seg):
        joined = ' '.join(str(c) for c in row if c)

        # Signature row
        if ncols >= 2 and (SIG1_RE.search(joined) or SIG2_RE.search(joined)
                           or SIG3_RE.search(joined)):
            parts = [str(c).strip() for c in row if str(c).strip()]
            cells = [''] * ncols
            if parts:
                cells[L0] = Paragraph(esc(krutidev_to_unicode(parts[0])),
                                      styles['sigdeva'])
            if len(parts) > 1:
                cells[R0] = Paragraph(esc(krutidev_to_unicode(parts[1])),
                                      styles['sigdeva'])
            data.append(cells)
            cmds += [('SPAN', (L0, r), (L1, r)), ('SPAN', (R0, r), (R1, r))]
            if r == first_sig:
                cmds += [
                    ('LINEABOVE', (L0, r), (L1, r), 0.5, C_RULE),
                    ('LINEABOVE', (R0, r), (R1, r), 0.5, C_RULE),
                    ('TOPPADDING', (0, r), (-1, r), 18),
                    ('BOTTOMPADDING', (0, r), (-1, r), 4),
                ]
            else:
                cmds += [
                    ('TOPPADDING', (0, r), (-1, r), 1),
                    ('BOTTOMPADDING', (0, r), (-1, r), 1),
                ]
            continue

        # Normal / header / total rows
        cells = []
        for ci, v in enumerate(row):
            s = fmt(v)
            if ci == 0 and NUMERIC_RE.match(s):
                txt, st = s, styles['no']   # serial numbers
            else:
                txt, st = pick_style(s, styles)
            cells.append(Paragraph(esc(txt), st))
        data.append(cells)

        if TITLE_RE.search(joined) and not HEAD_RE.search(joined):
            cmds += [
                ('SPAN',          (0, r), (-1, r)),
                ('BACKGROUND',    (0, r), (-1, r), C_TITLE_BG),
                ('LINEABOVE',     (0, r), (-1, r), 1.8, C_RULE),
                ('LINEBELOW',     (0, r), (-1, r), 0.5, C_RULE),
                ('TOPPADDING',    (0, r), (-1, r), 7),
                ('BOTTOMPADDING', (0, r), (-1, r), 5),
            ]
            # Rebuild the title cell with the title style.  Assigning .style
            # does not re-run Paragraph._setup(), so the cell would keep the
            # Helvetica font it was built with and render Devanagari as boxes.
            for ci in range(ncols):
                if str(row[ci]).strip():
                    cells[ci] = Paragraph(esc(fmt(row[ci])), styles['title'])
                    break

        elif HEAD_RE.search(joined):
            # Column header row: rebuild cells with text truncated to ~20 chars
            # Long Devanagari headers in narrow columns (5mm) would wrap to 20+
            # lines, making rows 200pt+ tall and the table too tall for the page.
            cells = [Paragraph(esc(str(fmt(v))[:20].rstrip()), styles['colhdr'])
                     for v in row]
            data[-1] = cells   # replace the cells already appended
            cmds += [
                ('BACKGROUND',    (0, r), (-1, r), C_HEAD_BG),
                ('TOPPADDING',    (0, r), (-1, r), 3),
                ('BOTTOMPADDING', (0, r), (-1, r), 3),
                ('LINEBELOW',     (0, r), (-1, r), 0.8, C_RULE),
            ]
            # Keep only the FIRST header block: some sheets repeat the header
            # row on every logical page.  Overwriting header_end with the last
            # match pushed repeatRows to ~1000, which made LongTable refuse to
            # split at all (LayoutError "too large on page").
            if not header_end:
                header_end = r + 1

        elif SUBHD_RE.search(str(row[0])) or (
            # A row with content ONLY in col 0 is a full-span sub-header
            sum(1 for c in row if str(c).strip()) == 1 and str(row[0]).strip()
        ):
            cmds += [
                ('SPAN',          (0, r), (-1, r)),
                ('BACKGROUND',    (0, r), (-1, r), C_HEAD_BG),
                ('TOPPADDING',    (0, r), (-1, r), 2),
                ('BOTTOMPADDING', (0, r), (-1, r), 2),
            ]

        elif TOTAL_RE.search(str(row[0])):
            cmds += [
                ('LINEABOVE',  (0, r), (-1, r), 1.0, C_INK),
                ('BACKGROUND', (0, r), (-1, r), C_TOTAL_BG),
            ]

        elif r % 2 == 0 and r > (header_end or 0):
            cmds.append(('BACKGROUND', (0, r), (-1, r), C_ALT_BG))

    # LongTable (not Table): a sheet rendered as ONE tall table (e.g. a
    # 1000-row auxiliary sheet) spans many pages; Table re-wraps its Paragraph
    # cells on every split and blows up, LongTable is built for this case.
    table = LongTable(data, colWidths=cw, repeatRows=header_end or 1)
    table.setStyle(TableStyle([
        ('BOX',           (0, 0), (-1, -1), 0.8,  C_RULE),
        ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING',   (0, 0), (-1, -1), 2.5),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 2.5),
        ('TOPPADDING',    (0, 0), (-1, -1), 2.0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2.0),
        ('LINEBELOW',     (0, 0), (-1, -2), 0.25, C_GRID),
        ('LINEAFTER',     (0, 0), (0,  -1), 0.4,  C_GRID),
        ('LINEAFTER',     (1, 0), (1,  -1), 0.4,  C_GRID),
        *cmds,
    ]))
    return table


class _PagedCanvas(_rl_canvas.Canvas):
    def __init__(self, *a, label='', font='Helvetica', **kw):
        super().__init__(*a, **kw)
        self._lbl, self._font, self._saved = label, font, []

    def showPage(self):
        self._saved.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._saved)
        for state in self._saved:
            self.__dict__.update(state)
            self._footer(total)
            super().showPage()
        super().save()

    def _footer(self, total: int):
        w, h = landscape(A4)
        self.saveState()
        self.setStrokeColor(C_RULE)
        self.setLineWidth(0.4)
        self.line(12 * mm, h - 5.5 * mm, w - 12 * mm, h - 5.5 * mm)
        self.setFont(self._font, 6)
        self.setFillColor(C_MUTE)
        self.drawString(12 * mm, 4 * mm, self._lbl)
        self.drawRightString(w - 12 * mm, 4 * mm,
                             f'Page {self._pageNumber} / {total}')
        self.restoreState()


def render(segs, out: Path, src: Path, mapping: dict, margin: float = 12.0):
    page = landscape(A4)
    doc  = BaseDocTemplate(
        str(out), pagesize=page,
        leftMargin=margin * mm,  rightMargin=margin * mm,
        topMargin=10 * mm,       bottomMargin=12 * mm,
        title=f'{src.stem} - budget estimates',
        author='PDFExcelExtract / apply.py',
        subject='A4 landscape - one page per signature block',
    )
    # Zero frame padding so table width == doc.width (prevents LayoutError)
    frame = Frame(
        doc.leftMargin, doc.bottomMargin, doc.width, doc.height,
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
        id='main',
    )
    doc.addPageTemplates([PageTemplate(id='p', frames=[frame])])

    styles = make_styles(mapping)
    ncols  = len(segs[0][0]) if segs else 13
    cw     = col_widths(ncols, doc.width)

    story: list = []
    for i, seg in enumerate(segs):
        # No KeepTogether: causes LayoutError when table > frame height.
        # repeatRows keeps column headers on continuation pages.
        # Explicit PageBreak guarantees each signed block starts on its own page.
        story.append(build_table(seg, cw, styles, mapping))
        if i < len(segs) - 1:
            story.append(PageBreak())

    label = (
        f'{src.name}  |  {dt.datetime.now():%d-%b-%Y %H:%M}'
        f'  |  {len(segs)} sections'
    )
    doc.build(
        story,
        canvasmaker=lambda *a, **k: _PagedCanvas(
            *a, label=label, font=F_LATIN, **k),
    )


def main() -> int:
    ap = argparse.ArgumentParser(
        description='Render Excel budget estimates as A4 landscape PDF')
    ap.add_argument('-i', '--input',  type=Path,
                    default=Path('INPUT/FINAL Budget Anuman 31 October 2026.xlsx'))
    ap.add_argument('-o', '--output', type=Path,
                    default=Path('OUTPUT/budget_paged.pdf'))
    ap.add_argument('-s', '--sheet',  default='Details',
                    help='Worksheet name to read (default: Details)')
    ap.add_argument('--all-sheets', action='store_true',
                    help='Render every worksheet to OUTPUT/<sheet>_paged.pdf')
    ap.add_argument('--font-dir',    type=Path, default=Path('fonts'),
                    help='Folder with the NotoSansDevanagari TTF font')
    ap.add_argument('--keep-hidden', action='store_true',
                    help='Include hidden columns M N O P Q')
    ap.add_argument('--margin',      type=float, default=12.0,
                    help='Page margin in mm (default 12)')
    ap.add_argument('--report',      action='store_true',
                    help='Print segment summary without rendering')
    args = ap.parse_args()

    if not args.input.exists():
        alt = Path(__file__).resolve().parent / args.input
        if alt.exists():
            args.input = alt

    if not args.input.exists():
        print(f'[error] Input not found: {args.input}', file=sys.stderr)
        return 2

    if args.all_sheets:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        mapping = register_fonts(args.font_dir)
        made = 0
        for name in list_sheets(args.input):
            rows = load_rows(args.input, sheet=name,
                             keep_hidden=args.keep_hidden,
                             native=(name != args.sheet))
            if not rows:
                print(f'[skip] {name}: empty sheet')
                continue
            segs = segment(rows)
            if not segs:
                print(f'[skip] {name}: no printable rows')
                continue
            if name == args.sheet:
                out = args.output
            else:
                safe = re.sub(r'[^0-9A-Za-z._-]+', '_', name).strip('_')
                out = args.output.parent / f'{safe}_paged.pdf'
            render(segs, out, args.input, mapping, args.margin)
            print(f'Saved  : {out}  ({len(segs)} sections)')
            made += 1
        if not made:
            print('[error] Nothing rendered.', file=sys.stderr)
            return 1
        return 0

    rows = load_rows(args.input, sheet=args.sheet,
                     keep_hidden=args.keep_hidden)
    if not rows:
        print('[error] Empty sheet.', file=sys.stderr)
        return 1

    segs = segment(rows)
    if not segs:
        print('[error] No segments found.', file=sys.stderr)
        return 1

    if args.report:
        col_desc = ('all 18 (A-R)' if args.keep_hidden
                    else '13 visible (A-L + R;  M-Q hidden)')
        print(f'Input    : {args.input}')
        print(f'Sheet    : {args.sheet}')
        print(f'Columns  : {len(rows[0])}  ({col_desc})')
        print(f'Segments : {len(segs)}  (= PDF pages)')
        for i, seg in enumerate(segs):
            title = next(
                (' '.join(str(c) for c in row if c)[:70]
                 for row in seg
                 if TITLE_RE.search(' '.join(str(c) for c in row if c))
                 and not HEAD_RE.search(' '.join(str(c) for c in row if c))),
                '(untitled)')
            print(f'  [{i+1:02d}]  {len(seg):3d} rows   {title}')
        return 0

    args.output.parent.mkdir(parents=True, exist_ok=True)
    mapping = register_fonts(args.font_dir)
    render(segs, args.output, args.input, mapping, args.margin)
    print(f'Saved  : {args.output}')
    print(f'Pages  : {len(segs)}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
