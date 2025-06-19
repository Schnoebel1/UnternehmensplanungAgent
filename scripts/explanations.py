"""
explanations.py – Llama-3-8B via langchain-ollama & JSON-Output
==============================================================
Lieferte bislang bei fehlerhaftem LLM-Output einen Crash, wenn `forecast=[]`.
Jetzt: robuster Fallback und optionale Übergabe von `forecast`.
"""

from __future__ import annotations
import os, csv, json, warnings, re
from pathlib import Path
from typing import List, Optional

from langchain_ollama import OllamaLLM

# ---------------- Paths & ENV ----------------
BASE         = Path(__file__).resolve().parent.parent
CONTEXT_PATH = BASE / "data" / "cases.csv"
LLM_LOG_PATH = BASE / "outputs" / "llm_debug.txt"

OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "llama3:8b")
OLLAMA_URL   = os.getenv("OLLAMA_URL",   "http://localhost:11434")
TEMPERATURE  = float(os.getenv("OLLAMA_TEMP", "0.4"))

# ---------------- Kontexte laden ----------------
def load_contexts(path: Path) -> List[str]:
    with path.open(encoding="utf-8") as f:
        return [row["description"].strip()
                for row in csv.DictReader(f)
                if row.get("description", "").strip()]

_contexts = load_contexts(CONTEXT_PATH)

# ---------------- LLM-Client (Singleton) ----------------
_llm: OllamaLLM | None = None
def _get_llm() -> OllamaLLM:
    global _llm
    if _llm is None:
        _llm = OllamaLLM(
            model       = OLLAMA_MODEL,
            base_url    = OLLAMA_URL,
            temperature = TEMPERATURE,
        )
    return _llm

# ---------------- Prompt-Templates ----------------
_SYSTEM_PROMPT = (
    "Du bist ein deutschsprachiger Finanzcontroller. "
    "Du erhältst die historischen Werte einer **einzelnen** Bilanzposition "
    "und allgemeine externe Sachverhalte. Deine Aufgabe ist, daraus einen "
    "t1–t3-Forecast zu erstellen und **konkret zu begründen**, warum sich "
    "genau diese Sachverhalte auf **diese** Position auswirken. "
    "Antworte **nur** im JSON-Format mit den Schlüsseln: "
    "`t1`, `t2`, `t3`, `reason`."
)

_HUMAN_TEMPLATE = """\
Hier die allgemeinen externen Sachverhalte:
{contexts}

Bilanzposition: **{account}**
Historische Werte:
- t-2: {t2:.2f}
- t-1: {t1:.2f}
- t0 : {t0:.2f}

Bitte liefere:
1) Prognosewerte für t1, t2, t3 (nur Zahlen)  
2) Eine kurze **account-spezifische** Begründung (`reason`, max. 20 Wörter).

Antworte in diesem JSON-Format:
{{"t1": <Zahl>, "t2": <Zahl>, "t3": <Zahl>, "reason": "<Kurztext>"}}
"""

_JSON_CLEAN_RE = re.compile(r".*?(\{.*\})", re.DOTALL)

# ---------------- Helper: Baseline-Forecast ---------------------------------
def _baseline_from_history(history: List[float]) -> List[float]:
    """Ein sehr einfacher CAGR-Baseline-Forecast aus t-2→t0 (falls machbar)."""
    try:
        t2, _, t0 = history
        if t2 not in (None, 0) and t0 not in (None,):
            years = 2
            growth = (t0 / t2) ** (1 / years) - 1
        else:
            growth = 0.0
    except Exception:
        growth = 0.0
    t0_val = history[-1] if history else 0.0
    return [round(t0_val * (1 + growth) ** i, 2) for i in range(1, 4)]

# ---------------- Öffentliche Funktion --------------------------------------
def explain(account: str,
            history: List[float],
            forecast: Optional[List[float]] = None) -> str:
    """
    Holt Forecast & Reason vom LLM.  Auf Fehler → eigenes JSON mit Baseline-Forecast.
    """
    t2, t1, t0 = history
    contexts_str = "\n".join(f"- {c}" for c in _contexts)

    prompt = _SYSTEM_PROMPT + "\n\n" + _HUMAN_TEMPLATE.format(
        contexts = contexts_str,
        account  = account,
        t2       = t2,
        t1       = t1,
        t0       = t0,
    )

    # ---- Debug-Log ----------------------------------------------------------
    LLM_LOG_PATH.parent.mkdir(exist_ok=True, parents=True)
    with LLM_LOG_PATH.open("a", encoding="utf-8") as lf:
        lf.write("\n" + "=" * 60 + "\n")
        lf.write(f"ACCOUNT: {account}\nPROMPT:\n{prompt}\n")

    # ---- Aufruf & Parsing ---------------------------------------------------
    try:
        raw = _get_llm().invoke(prompt).strip()

        # Log:
        with LLM_LOG_PATH.open("a", encoding="utf-8") as lf:
            lf.write("\nRAW_RESPONSE:\n" + raw + "\n")

        # JSON herausfiltern
        m = _JSON_CLEAN_RE.match(raw)
        if not m:
            raise ValueError("Kein JSON-Block gefunden")

        json_text = m.group(1)
        json.loads(json_text)           # Validierungs-Probe
        return json_text

    # ---- Fallback -----------------------------------------------------------
    except Exception as e:
        with LLM_LOG_PATH.open("a", encoding="utf-8") as lf:
            lf.write(f"\nERROR during explain(): {e}\n")

        warnings.warn(
            f"Ollama/LangChain Fehler: {e!s} – liefere Fallback-Forecast",
            stacklevel=2,
        )

        # Baseline bestimmen
        baseline = forecast if forecast and len(forecast) >= 3 else _baseline_from_history(history)
        f1, f2, f3 = baseline[:3]

        return json.dumps(
            {
                "t1": f1,
                "t2": f2,
                "t3": f3,
                "reason": f"CAGR-Baseline für {account}",
            },
            ensure_ascii=False,
        )
