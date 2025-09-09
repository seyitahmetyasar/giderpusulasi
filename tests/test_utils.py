import pytest
from arama23 import extract_imeis, _luhn_ok_imei, brand_from_text


@pytest.mark.parametrize('imei,expected', [
    ("490154203237518", True),
    ("490154203237517", False),
    ("12345", False),
])
def test_luhn_ok_imei(imei, expected):
    assert _luhn_ok_imei(imei) is expected


def test_extract_imeis_deduplicates_and_validates():
    text = (
        "IMEI 490154203237518 and 352099001761481 "
        "and invalid 123456789012345 and again 490154203237518"
    )
    assert extract_imeis(text) == ["352099001761481", "490154203237518"]


@pytest.mark.parametrize('text,brand', [
    ("My new iPhone 11 is great", "APPLE"),
    ("Samsung Galaxy S22", "SAMSUNG"),
    ("Huawei Honor Magic", "HONOR"),
    ("Unknown device", "Bilinmeyen"),
])
def test_brand_from_text(text, brand):
    assert brand_from_text(text) == brand
