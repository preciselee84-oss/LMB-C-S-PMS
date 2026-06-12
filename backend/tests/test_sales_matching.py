from app.services.sales_matching import amount_matches, normalize_name


def test_normalize_name_removes_spaces_and_case() -> None:
    assert normalize_name(" OO Company ") == "oocompany"


def test_amount_matches_exact_amount() -> None:
    assert amount_matches(5_000_000, 5_000_000) == (True, "exact")


def test_amount_matches_vat_included_amount() -> None:
    assert amount_matches(5_000_000, 5_500_000) == (True, "vat_included")


def test_amount_does_not_match_unknown_amount() -> None:
    assert amount_matches(5_000_000, 4_900_000) == (False, None)
