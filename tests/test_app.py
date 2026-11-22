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
    at.checkbox(key="mari").uncheck().run()      # datele de test nu sunt din ligi "cunoscute"
    assert not at.exception, at.exception
    at.button(key="propune").click().run()
    assert not at.exception, at.exception
    first = list(at.session_state["bilet"])
    print("propus:", first, "metrici:", [(m.label, m.value) for m in at.sidebar.metric])
    assert first, "nu a propus nimic"
    gone = first[0]
    at.session_state["bilet"] = first[1:]
    at.run()
    at.button(key="propune").click().run()
    assert not at.exception, at.exception
    second = list(at.session_state["bilet"])
    print("dupa ce am scos", gone, "->", second)
    assert not any(x.split(":")[0] == gone.split(":")[0] for x in second), "meciul scos a revenit"
    print("OK")


if __name__ == "__main__":
    test_app()
