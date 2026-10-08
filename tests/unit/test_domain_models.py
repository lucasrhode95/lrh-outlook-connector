from __future__ import annotations

import pytest
from pydantic import ValidationError

from outlook_connector.domain.models import Coverage


def test_coverage_accepts_only_known_exclusion_reasons() -> None:
    result = Coverage(complete=True, excluded={"hidden": 2})
    assert result.excluded == {"hidden": 2}

    with pytest.raises(ValidationError):
        Coverage(complete=True, excluded={"unknown": 1})
