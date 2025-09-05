from arama23 import _build_header_map, parse_gp_template_workbook, HEADERS


class FakeCell:
    def __init__(self, value):
        self.value = value


class FakeWorksheet:
    def __init__(self, rows, title="Sheet1"):
        self._rows = rows
        self.title = title
        self.max_row = len(rows)
        self.max_column = max(len(r) for r in rows)

    def cell(self, row, column):
        try:
            return FakeCell(self._rows[row - 1][column - 1])
        except IndexError:
            return FakeCell(None)


class FakeWorkbook:
    def __init__(self, worksheets):
        self.worksheets = worksheets


def test_build_header_map_detects_headers():
    rows = [
        ["IMEI", "Belge Türü", "MODEL", "SATIŞ TARİHi", "ALICI ADI SOYADI"],
        [None, None, None, None, None],
    ]
    ws = FakeWorksheet(rows)
    hm = _build_header_map(ws)
    assert hm == {
        "row": 1,
        "cols": {
            0: "imei",
            1: "Belge Türü",
            2: "MODEL",
            3: "SATIŞ TARİHi",
            4: "ALICI ADI SOYADI",
        },
    }


def test_parse_gp_template_workbook_extracts_rows():
    rows = [
        ["IMEI", "Belge Türü", "MODEL", "SATIŞ TARİHi", "ALICI ADI SOYADI"],
        ["490154203237518", "GIDER PUSULASI", "iPhone 11", "2024-01-01", "Alice"],
        ["invalid", "", "IMEI: 352099001761481", "", ""],
    ]
    ws = FakeWorksheet(rows)
    wb = FakeWorkbook([ws])
    logs = []
    out = parse_gp_template_workbook(wb, logs.append)

    assert logs == ["[GP] Sayfa 'Sheet1': 2 satır alındı."]
    assert len(out) == 2

    i_imei = HEADERS.index("imei")
    i_brand = HEADERS.index("Marka")
    i_status = HEADERS.index("DURUMU")
    i_loc = HEADERS.index("Bulunma")
    i_doc_type = HEADERS.index("ALIŞ BELGELERİ TÜRÜ")

    row1, row2 = out
    assert row1[i_imei] == "490154203237518"
    assert row1[i_brand] == "APPLE"
    assert row1[i_status] == "Satılmış"
    assert row1[i_loc] == "GP"
    assert row1[i_doc_type] == "Gider Pusulası"

    assert row2[i_imei] == "352099001761481"
    assert row2[i_brand] == "Bilinmeyen"
    assert row2[i_status] == "Satılabilir"
    assert row2[i_loc] == "GP"
