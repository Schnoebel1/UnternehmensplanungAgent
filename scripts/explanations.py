"""
explanations.py – Llama-3-8B via langchain-ollama & JSON-Output
==============================================================
Hier haben wir das Prompt so umformuliert, dass die “reason” sich
unmittelbar auf die jeweilige Bilanzposition bezieht.
"""

from __future__ import annotations
import os, csv, json, warnings, re
from pathlib import Path
from typing import List

from langchain_ollama import OllamaLLM

# ---------------- Paths & ENV ----------------
BASE         = Path(__file__).resolve().parent.parent
CONTEXT_PATH = BASE / "data" / "cases.csv"
LLM_LOG_PATH = BASE / "outputs" / "llm_debug.txt"

OLLAMA_MODEL  = os.getenv("OLLAMA_MODEL", "llama3:8b")
OLLAMA_URL    = os.getenv("OLLAMA_URL",   "http://localhost:11434")
TEMPERATURE   = float(os.getenv("OLLAMA_TEMP", "0.4"))

# ---------------- Kontexte laden ----------------
def load_contexts(path: Path) -> List[str]:
    ctxs: List[str] = []
    with path.open(encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            desc = row.get("description", "").strip()
            if desc:
                ctxs.append(desc)
    return ctxs

_contexts = load_contexts(CONTEXT_PATH)

# ---------------- LLM-Client (Singleton) ----------------
_llm: OllamaLLM | None = None
def _get_llm() -> OllamaLLM:
    global _llm
    if _llm is None:
        _llm = OllamaLLM(
            model=OLLAMA_MODEL,
            base_url=OLLAMA_URL,
            temperature=TEMPERATURE,
        )
    return _llm

# ---------------- Prompt-Templates ----------------
_SYSTEM_PROMPT = (
    "Du bist ein deutschsprachiger Finanzcontroller. "
    "Du erhältst die historischen Werte einer **einzelnen** Bilanzposition"
    " und allgemeine externe Sachverhalte. Deine Aufgabe ist, daraus einen"
    " t1–t3-Forecast zu erstellen und **konkret zu begründen**, warum sich"
    " genau diese externen Sachverhalte auf **diese** Bilanzposition auswirken."
    " Antworte ausschließlich im JSON-Format mit diesen Schlüsseln: `t1`, `t2`, `t3`, `reason`."
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
2) Eine kurze **account-spezifische** Begründung (`reason`, maximal 20 Wörter),  
   die erläutert, **weshalb** gerade diese externen Sachverhalte die Kennzahl der"
   " Position **{account}** beeinflussen.

Antworte in diesem JSON-Format:
{{"t1": <Zahl>, "t2": <Zahl>, "t3": <Zahl>, "reason": "<Kurztext>"}}
"""

# Regex, um alles vor erstem '{' abzuschneiden
_JSON_CLEAN_RE = re.compile(r".*?(\{.*\})", re.DOTALL)

# ---------------- Öffentliche Funktion ----------------
def explain(account: str,
            history: List[float],
            forecast: List[float]) -> str:
    """Ruft das LLM auf, extrahiert den JSON-Block und macht Fallback bei Fehlern."""
    t2, t1, t0 = history
    contexts_str = "\n".join(f"- {c}" for c in _contexts)

    prompt = _SYSTEM_PROMPT + "\n\n" + _HUMAN_TEMPLATE.format(
        contexts=contexts_str,
        account=account,
        t2=t2, t1=t1, t0=t0
    )

    # Debug-Log initialisieren
    LLM_LOG_PATH.parent.mkdir(exist_ok=True, parents=True)
    with LLM_LOG_PATH.open("a", encoding="utf-8") as lf:
        lf.write("\n" + "="*60 + "\n")
        lf.write(f"ACCOUNT: {account}\nPROMPT:\n{prompt}\n")

    try:
        llm = _get_llm()
        raw = llm.invoke(prompt).strip()

        # raw-Log
        with LLM_LOG_PATH.open("a", encoding="utf-8") as lf:
            lf.write("\nRAW_RESPONSE:\n" + raw + "\n")

        # JSON-Extraktion
        m = _JSON_CLEAN_RE.match(raw)
        if not m:
            raise ValueError("Kein JSON-Block gefunden")
        json_text = m.group(1)
        # Validierung
        _ = json.loads(json_text)
        return json_text

    except Exception as e:
        with LLM_LOG_PATH.open("a", encoding="utf-8") as lf:
            lf.write(f"\nERROR during explain(): {e}\n")
        warnings.warn(f"Ollama/LangChain Fehler: {e!s} – Fallback-JSON", stacklevel=2)
        # Fallback
        f1, f2, f3 = (round(x,2) for x in forecast)
        return json.dumps({
            "t1": f1,
            "t2": f2,
            "t3": f3,
            "reason": f"CAGR-Baseline für {account}"
        }, ensure_ascii=False)
