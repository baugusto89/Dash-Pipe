"""Generate sample IT ticket XLSX files for testing."""
import io
import random
from datetime import datetime, timedelta

import numpy as np
import pandas as pd

TIPOS = [
    ["Acesso ao Sistema", "Reset de Senha", "Permissão Negada", "Bloqueio de Conta"],
    ["Troca de Equipamento", "Manutenção de Hardware", "Instalação de Software", "Configuração de Periférico"],
    ["Lentidão no Sistema", "Erro de Aplicação", "Atualização de Sistema", "Implantação de Módulo"],
    ["Sem Conexão à Internet", "Lentidão na Rede", "Configuração de VPN", "Falha em Switch"],
    ["Vírus/Malware", "Acesso Indevido", "Vazamento de Dados", "Atualização de Antivírus"],
    ["Solicitação Geral", "Dúvida de Uso", "Melhoria de Processo", "Consultoria TI"],
]

CATEGORIAS = [
    ["Active Directory", "ERP", "E-mail", "Portal Interno", "MFA"],
    ["Notebook", "Desktop", "Monitor", "Impressora", "Headset"],
    ["ERP - Financeiro", "ERP - RH", "CRM", "BI/Dashboard", "Intranet"],
    ["Wi-Fi", "LAN", "WAN", "VPN", "Firewall"],
    ["Endpoint Security", "SIEM", "DLP", "Pen Test", "Conformidade"],
    ["Procedimento", "Treinamento", "Documentação", "Relatório", "Outro"],
]

FILE_NAMES = [
    "[TI] Acessos Novos Colaboradores",
    "Infraestrutura e Hardware",
    "Sistemas Internos",
    "Redes e Conectividade",
    "[TI] Sistemas",
    "Outros / Geral",
]

PRIORIDADES = ["Crítica", "Alta", "Média", "Baixa"]
PRIORIDADE_WEIGHTS = [0.05, 0.20, 0.50, 0.25]
SLA_BASE = [0.88, 0.75, 0.92, 0.80, 0.70, 0.95]


def _random_dates(n: int, start: datetime, end: datetime) -> list:
    span = int((end - start).total_seconds())
    return [start + timedelta(seconds=random.randint(0, span)) for _ in range(n)]


def generate_file(index: int, n_rows: int = 300) -> bytes:
    """Return bytes of an XLSX file with realistic IT ticket data."""
    rng = np.random.default_rng(seed=index)
    tipos = TIPOS[index]
    categorias = CATEGORIAS[index]
    sla_rate = SLA_BASE[index]

    start = datetime(2024, 1, 1)
    end = datetime(2025, 5, 31)

    datas_abertura = _random_dates(n_rows, start, end)

    sla_cumprido = rng.random(n_rows) < sla_rate
    status_vals = []
    for sla in sla_cumprido:
        if sla:
            status_vals.append("Fechado")
        else:
            status_vals.append(random.choice(["Aberto", "Em Andamento", "Fechado"]))

    datas_fechamento = []
    for i, (da, s) in enumerate(zip(datas_abertura, status_vals)):
        if s == "Fechado":
            datas_fechamento.append(da + timedelta(hours=int(rng.integers(1, 120))))
        else:
            datas_fechamento.append(None)

    df = pd.DataFrame({
        "ID Chamado": [f"CH{str(index + 1)}{str(i + 1).zfill(5)}" for i in range(n_rows)],
        "Tipo": rng.choice(tipos, n_rows),
        "Categoria": rng.choice(categorias, n_rows),
        "Prioridade": rng.choice(PRIORIDADES, n_rows, p=PRIORIDADE_WEIGHTS),
        "Data Abertura": datas_abertura,
        "Data Fechamento": datas_fechamento,
        "SLA Cumprido": ["Sim" if v else "Não" for v in sla_cumprido],
        "Status": status_vals,
        "Atendente": rng.choice(
            [f"Analista {chr(65 + j)}" for j in range(6)], n_rows
        ),
        "Solicitante": [f"Colaborador {rng.integers(1, 200)}" for _ in range(n_rows)],
    })

    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine='openpyxl')
    buf.seek(0)
    return buf.getvalue()


def generate_all() -> dict:
    """Return {name: bytes} for all 6 sample files."""
    return {FILE_NAMES[i]: generate_file(i) for i in range(6)}
