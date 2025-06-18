def explain(account: str,
            history: list[float],
            forecast: list[float]) -> str:
    """Hier später dein LangChain-Aufruf."""
    h = ", ".join(f"{x:.0f}" for x in history)
    f = ", ".join(f"{x:.0f}" for x in forecast)
    return f"CAGR auf Basis ({h}) ⇒ {f}"
