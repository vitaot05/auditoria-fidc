import json
from pathlib import Path
from collections import defaultdict
from datetime import datetime

BASE = Path(__file__).parent
HIST = BASE / "historico.json"
DADOS = BASE / "dados.json"

def atualizar_dashboard(resultado_final):

    historico = []
    if HIST.exists():
        try:
            historico = json.loads(HIST.read_text(encoding="utf-8"))
        except:
            historico = []

    historico.extend(resultado_final)

    HIST.write_text(
        json.dumps(historico, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    fundos = defaultdict(list)

    for r in historico:
        fundos[r["fundo"]].append(r)

    dashboard = {
        "ultima_atualizacao": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "fundos": []
    }

    for fundo, registros in fundos.items():
        total = len(registros)
        ok = sum(1 for x in registros if x["status"] == "OK")
        erros = total - ok

        dashboard["fundos"].append({
            "nome": fundo,
            "taxa_sucesso": round((ok/total)*100,2) if total else 0,
            "total_registros": total,
            "falhas": erros,
            "ultimos_registros": registros[-20:]
        })

    DADOS.write_text(
        json.dumps(dashboard, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
