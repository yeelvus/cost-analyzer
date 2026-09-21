from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from budget.glodon.pipeline import extract_project_code_and_name


def test_extract_project_name_from_prefixed_export():
    assert extract_project_code_and_name("标段人材机汇总-tomax.xls") == (None, "tomax")
    assert extract_project_code_and_name("标段人材机汇总-伟隆阀门.xls") == (None, "伟隆阀门")


def test_extract_project_code_and_name_from_standard_export():
    assert extract_project_code_and_name(
        "THCC-2025-015_腾洋项目室外_标段人材机汇总.xls"
    ) == ("THCC-2025-015", "腾洋项目室外")