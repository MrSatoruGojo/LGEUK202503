# tests/test_date_utils.py

import pytest
from src.utils.date_utils import date_to_week, date_to_year, generate_weeks_range_monday

def test_date_to_week_basic():
    # January 1, 2021 was a Friday in ISO week 53 of 2020
    assert date_to_week("20210101") == 53

    # A mid‐year date
    assert date_to_week("20210415") == 15  # 15th week of the year

def test_date_to_year_basic():
    assert date_to_year("20210101") == 2021
    assert date_to_year("19991231") == 1999

def test_generate_weeks_range_monday():
    # From Monday Jan 4 2021 to Sunday Jan 17 2021 covers weeks 1 and 2
    weeks = generate_weeks_range_monday("20210104", "20210117")
    # Expect labels like "21-01-04\n(W1)" and "21-01-11\n(W2)"
    assert any("W1" in w for w in weeks)
    assert any("W2" in w for w in weeks)
    # And exactly two entries
    assert len(weeks) == 2
