"""Testeaza aplicatia Streamlit pe datele simulate (ruleaza pe GitHub)."""
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))


def test_app():
    import tests.test_run as tr
    from betpro import config
    tr.test_full_run()
    src = config.DATA
    dst = ROOT / "data"
    dst.mkdir(exist_ok=True)
    for f in ("today.json", "picks.csv", "tickets.csv"):
        if (src / f).exists():
            shutil.copy(src / f, dst / f)
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=60).run()
    assert not at.exception, at.exception
    print("titlu:", [t.value for t in at.title])
    print("tabele:", len(at.dataframe))
    ms = at.multiselect(key="sel")
    ms.select(ms.options[0]).select(ms.options[5]).run()
    assert not at.exception, at.exception
    print("metrici:", [(m.label, m.value) for m in at.metric])
    at.button[0].click().run()
    assert not at.exception, at.exception
    print("propus:", at.multiselect(key="sel").value)
    print("OK")


if __name__ == "__main__":
    test_app()
