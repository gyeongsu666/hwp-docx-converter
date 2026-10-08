"""HWP5 레코드 태그 상수 (한글 문서 파일 형식 5.0 공개 문서 기준)."""

BEGIN = 0x010  # 16

# --- DocInfo ---------------------------------------------------------------
DOCUMENT_PROPERTIES = BEGIN + 0     # 16
ID_MAPPINGS = BEGIN + 1             # 17
BIN_DATA = BEGIN + 2                # 18
FACE_NAME = BEGIN + 3               # 19
BORDER_FILL = BEGIN + 4             # 20
CHAR_SHAPE = BEGIN + 5              # 21
TAB_DEF = BEGIN + 6                 # 22
NUMBERING = BEGIN + 7               # 23
BULLET = BEGIN + 8                  # 24
PARA_SHAPE = BEGIN + 9              # 25
STYLE = BEGIN + 10                  # 26
DOC_DATA = BEGIN + 11               # 27
DISTRIBUTE_DOC_DATA = BEGIN + 12    # 28
COMPATIBLE_DOCUMENT = BEGIN + 14    # 30
LAYOUT_COMPATIBILITY = BEGIN + 15   # 31
TRACKCHANGE = BEGIN + 16            # 32
MEMO_SHAPE = BEGIN + 60             # 76
FORBIDDEN_CHAR = BEGIN + 62         # 78
TRACK_CHANGE = BEGIN + 64           # 80
TRACK_CHANGE_AUTHOR = BEGIN + 65    # 81

# --- BodyText --------------------------------------------------------------
PARA_HEADER = BEGIN + 50            # 66
PARA_TEXT = BEGIN + 51              # 67
PARA_CHAR_SHAPE = BEGIN + 52        # 68
PARA_LINE_SEG = BEGIN + 53          # 69
PARA_RANGE_TAG = BEGIN + 54         # 70
CTRL_HEADER = BEGIN + 55            # 71
LIST_HEADER = BEGIN + 56            # 72
PAGE_DEF = BEGIN + 57               # 73
FOOTNOTE_SHAPE = BEGIN + 58         # 74
PAGE_BORDER_FILL = BEGIN + 59       # 75
SHAPE_COMPONENT = BEGIN + 60        # 76
TABLE = BEGIN + 61                  # 77
SHAPE_COMPONENT_LINE = BEGIN + 62   # 78
SHAPE_COMPONENT_RECTANGLE = BEGIN + 63   # 79
SHAPE_COMPONENT_ELLIPSE = BEGIN + 64     # 80
SHAPE_COMPONENT_ARC = BEGIN + 65         # 81
SHAPE_COMPONENT_POLYGON = BEGIN + 66     # 82
SHAPE_COMPONENT_CURVE = BEGIN + 67       # 83
SHAPE_COMPONENT_OLE = BEGIN + 68         # 84
SHAPE_COMPONENT_PICTURE = BEGIN + 69     # 85
SHAPE_COMPONENT_CONTAINER = BEGIN + 70   # 86
CTRL_DATA = BEGIN + 71              # 87
EQEDIT = BEGIN + 72                 # 88
SHAPE_COMPONENT_TEXTART = BEGIN + 74     # 90
FORM_OBJECT = BEGIN + 75            # 91

# --- 문단 내 제어 문자 ------------------------------------------------------
#: 2바이트만 차지하는 제어 문자
CHAR_CONTROLS = {0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31}
#: 16바이트(8 WCHAR)를 차지하는 인라인 제어 문자
INLINE_CONTROLS = {4, 5, 6, 7, 8, 9, 19, 20}
#: 16바이트(8 WCHAR)를 차지하는 확장 제어 문자 — 표·그림 등 개체가 여기 달린다
EXTENDED_CONTROLS = {1, 2, 3, 11, 12, 14, 15, 16, 17, 18, 21, 22, 23}

LINE_BREAK = 10
PARA_BREAK = 13
HYPHEN = 24
NBSP = 30
FIXED_SPACE = 31
TAB = 9

#: CTRL_HEADER의 ctrl id (4글자, 리틀엔디언이라 뒤집어 읽는다)
CTRL_TABLE = "tbl "
CTRL_SHAPE = "gso "
CTRL_SECTION_DEF = "secd"
CTRL_COLUMN_DEF = "cold"
CTRL_HEADER_AREA = "head"
CTRL_FOOTER_AREA = "foot"
CTRL_FOOTNOTE = "fn  "
CTRL_ENDNOTE = "en  "
CTRL_EQUATION = "eqed"
CTRL_PAGE_NUM_POS = "pgnp"
CTRL_AUTO_NUM = "atno"
CTRL_NEW_NUM = "nwno"
CTRL_PAGE_HIDE = "pghd"
CTRL_PAGE_ODD_EVEN = "pgct"
CTRL_INDEX_MARK = "idxm"
CTRL_BOOKMARK = "bokm"
CTRL_DUTMAL = "tdut"
CTRL_HIDDEN_COMMENT = "tcmt"


def ctrl_id_to_str(value: int) -> str:
    """CTRL_HEADER 앞 4바이트 UINT32를 사람이 읽는 4글자로."""
    return "".join(chr((value >> shift) & 0xFF) for shift in (24, 16, 8, 0))
