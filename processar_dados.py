"""
processar_dados.py
==================
Lê arquivos MPP nativos do Microsoft Project e retorna um DataFrame limpo
e enriquecido para uso no dashboard Streamlit.

Pré-requisitos:
  - pip install mpxj
  - Java JRE/JDK 11+ instalado e disponível no PATH
"""

import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO
# ---------------------------------------------------------------------------
# Data de referência para calcular atraso
HOJE = date.today()


# ---------------------------------------------------------------------------
# FUNÇÕES AUXILIARES
# ---------------------------------------------------------------------------

def _converter_data(valor) -> date | None:
    """
    Converte datas para datetime.date.
    Suporta:
      - Objetos datetime/date nativos (vindos do Excel via openpyxl)
      - String 'Seg 09/02/26'   (exportação CSV do MSP)
      - String '2026-01-05 ...' (Excel lido como str)
      - String 'DD/MM/YYYY' e 'DD/MM/YY'
    """
    if valor is None:
        return None
    # Já é date/datetime (leitura nativa do Excel)
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if not isinstance(valor, str) or not valor.strip():
        return None
    valor = valor.strip()
    # Remove o dia da semana abreviado (Seg, Ter, ...)
    for pt in DIAS_SEMANA:
        valor = valor.replace(pt + " ", "").replace(pt, "")
    valor = valor.strip()
    # Tenta formatos em ordem de probabilidade
    for fmt in ("%d/%m/%y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(valor, fmt).date()
        except ValueError:
            continue
    return None


def _converter_percentual(valor) -> float | None:
    """
    Converte percentual para float na escala 0-100.
    Suporta:
      - '16%'  → 16.0   (string com símbolo %)
      - '0.16' → 16.0   (decimal 0-1, comum no Excel)
      - 16     → 16.0   (número inteiro)
      - 0.16   → 16.0   (float 0-1)
    """
    if valor is None:
        return None
    # Valores numéricos nativos (Excel)
    if isinstance(valor, (int, float)):
        v = float(valor)
        # Excel armazena % como decimal (0.16 = 16%)
        return round(v * 100, 1) if v <= 1.0 else round(v, 1)
    if not isinstance(valor, str):
        return None
    tem_simbolo = "%" in valor
    m = re.search(r"(\d+(?:[.,]\d+)?)", valor)
    if not m:
        return None
    v = float(m.group(1).replace(",", "."))
    # Se não tem %, trata como decimal quando <= 1
    if not tem_simbolo and v <= 1.0:
        return round(v * 100, 1)
    return round(v, 1)


def _converter_duracao(valor) -> float | None:
    """
    Converte duração para float (dias).
    Suporta '248 dias?', '248d', e valores numéricos nativos do Excel.
    """
    if isinstance(valor, (int, float)):
        return float(valor)
    if not isinstance(valor, str):
        return None
    m = re.search(r"(\d+(?:[.,]\d+)?)", valor)
    if m:
        return float(m.group(1).replace(",", "."))
    return None


# ---------------------------------------------------------------------------
# LEITURA NATIVA DE ARQUIVOS .MPP (MS Project)
# Estratégia: converte .mpp → JSON via mpxj.jar + Java subprocess (sem JPype).
# Requisitos: pip install mpxj  |  Java JRE instalado e no PATH
# ---------------------------------------------------------------------------

def _mpxj_jar_classpath() -> str:
    """
    Retorna o classpath com todos os JARs do pacote mpxj instalado.
    Lança FileNotFoundError se o pacote não estiver disponível.
    Usa os.pathsep como separador (';' no Windows, ':' no Linux/Mac).
    """
    import os
    import site

    candidatos = site.getsitepackages() + [site.getusersitepackages()]
    for sp in candidatos:
        lib_dir = Path(sp) / "mpxj" / "lib"
        if lib_dir.exists():
            jars = sorted(lib_dir.glob("*.jar"))
            if jars:
                return os.pathsep.join(str(j) for j in jars)
    raise FileNotFoundError(
        "JARs do mpxj não encontrados. Instale com: pip install mpxj"
    )


def _mpp_para_json(caminho_mpp: Path) -> dict:
    """
    Converte arquivo .mpp para dict Python usando o mpxj.jar via subprocess Java.
    Não depende de JPype – requer apenas Java no PATH.
    """
    import json as _json
    import subprocess
    import tempfile

    # Verifica Java
    java_check = subprocess.run(
        ["java", "-version"], capture_output=True, text=True
    )
    if java_check.returncode != 0:
        raise EnvironmentError(
            "Java não encontrado no PATH.\n"
            "Baixe e instale em: https://adoptium.net/"
        )

    try:
        cp = _mpxj_jar_classpath()
    except FileNotFoundError as exc:
        raise ImportError(str(exc)) from exc

    # Arquivo temporário para saída JSON
    with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as tmp:
        json_path = Path(tmp.name)

    try:
        result = subprocess.run(
            ["java", "-cp", cp,
             "org.mpxj.sample.MpxjConvert",
             str(caminho_mpp), str(json_path)],
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"Falha ao converter {caminho_mpp.name}:\n"
                f"{(result.stderr or result.stdout)[:800]}"
            )
        with open(json_path, encoding="utf-8") as f:
            return _json.load(f)
    finally:
        try:
            json_path.unlink()
        except Exception:
            pass


def _parse_iso_date(valor: "str | None") -> "date | None":
    """Converte string ISO 'YYYY-MM-DDTHH:MM:SS.f' para datetime.date."""
    if not valor:
        return None
    try:
        return datetime.fromisoformat(str(valor)[:10]).date()
    except Exception:
        return None


def _ler_mpp(caminho: Path) -> pd.DataFrame:
    """
    Lê arquivo .mpp nativo do MS Project usando o mpxj.jar via Java subprocess.

    Pré-requisitos:
        - pip install mpxj
        - Java JRE instalado e disponível no PATH

    Retorna DataFrame com colunas:
        nome, nivel, pct_concluido, duracao_dias,
        inicio, termino, predecessoras, recursos, recursos_lista, observacao
    """
    data = _mpp_para_json(caminho)

    # Mapa resource_unique_id → nome do recurso
    recursos_map: dict[int, str] = {}
    for r in data.get("resources", []):
        uid = r.get("unique_id")
        nome_r = r.get("name")
        if uid is not None and nome_r:
            recursos_map[int(uid)] = str(nome_r).strip()

    # Mapa task_unique_id → lista de nomes de recursos (via assignments)
    assignments_map: dict[int, list[str]] = {}
    for a in data.get("assignments", []):
        t_uid = a.get("task_unique_id")
        r_uid = a.get("resource_unique_id")
        if t_uid is None:
            continue
        r_nome = recursos_map.get(int(r_uid), "") if r_uid is not None else ""
        if r_nome:
            assignments_map.setdefault(int(t_uid), []).append(r_nome)

    registros: list[dict] = []
    for task in data.get("tasks", []):
        nome_str = (task.get("name") or "").strip()
        if not nome_str:
            continue

        outline_level = task.get("outline_level")
        if outline_level is None:
            continue  # pula o resumo do projeto (nó raiz artificial)

        # outline_level=1 (primeira tarefa real) → nivel=0,
        # outline_level=2 → nivel=1, etc. — igual ao comportamento do CSV
        nivel = max(0, int(outline_level) - 1)

        # % concluído (0 – 100)
        pct_raw = task.get("percent_complete")
        try:
            pct = float(pct_raw) if pct_raw is not None else None
        except (ValueError, TypeError):
            pct = None

        # Duração: o mpxj JSON armazena em segundos (8 h/dia = 28 800 s)
        dur_sec = task.get("duration")
        try:
            duracao_dias = round(float(dur_sec) / 28800.0, 2) if dur_sec is not None else None
        except (ValueError, TypeError):
            duracao_dias = None

        inicio = _parse_iso_date(task.get("start"))
        termino = _parse_iso_date(task.get("finish"))

        # Recursos via mapa de assignments
        uid = task.get("unique_id")
        recursos_lista = list(dict.fromkeys(  # remove duplicatas mantendo ordem
            assignments_map.get(int(uid), [])
        )) if uid is not None else []
        recursos_str = ";".join(recursos_lista)

        # Observações (notas não são exportadas no JSON do mpxj)
        obs = ""

        registros.append({
            "nome": nome_str,
            "nivel": nivel,
            "pct_concluido": pct,
            "duracao_dias": duracao_dias,
            "inicio": inicio,
            "termino": termino,
            "predecessoras": "",
            "recursos": recursos_str,
            "recursos_lista": recursos_lista,
            "observacao": obs,
        })

    if not registros:
        raise ValueError(f"Nenhuma tarefa encontrada no arquivo: {caminho.name}")

    return pd.DataFrame(registros)


# ---------------------------------------------------------------------------
# FUNÇÃO PRINCIPAL
# ---------------------------------------------------------------------------

def carregar_dados(caminho: Path | str) -> pd.DataFrame:
    """
    Carrega e processa arquivo .mpp nativo do MS Project.

    Retorna um DataFrame com as colunas:
        nome, nivel, fase, pct_concluido, duracao_dias,
        inicio, termino, predecessoras, recursos, recursos_lista,
        observacao, status, atrasada

    Pré-requisitos:
        - pip install mpxj
        - Java JRE/JDK 11+ instalado e no PATH
    """
    caminho = Path(caminho)
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    if caminho.suffix.lower() != ".mpp":
        raise ValueError(
            f"Formato não suportado: '{caminho.suffix}'. "
            "Envie um arquivo .mpp nativo do MS Project."
        )

    # ── Leitura MPP ────────────────────────────────────────────────────────────────
    df_mpp = _ler_mpp(caminho)

    # Determina a fase (nível 1 da hierarquia)
    fases_stack_m: list[str] = []
    faselist_m: list[str] = []
    for _, row in df_mpp.iterrows():
        nivel_m = row["nivel"]
        nome_m = row["nome"]
        while len(fases_stack_m) > nivel_m:
            fases_stack_m.pop()
        fases_stack_m.append(nome_m)
        faselist_m.append(
            fases_stack_m[1] if len(fases_stack_m) > 1 else nome_m
        )
    df_mpp["fase"] = faselist_m

    # Status calculado
    def _status_mpp(row):
        pct = row["pct_concluido"]
        termino_r = row["termino"]
        if pct is None:
            return "Indefinido"
        if pct >= 100:
            return "Concluída"
        if termino_r and termino_r < HOJE and pct < 100:
            return "Atrasada"
        if pct > 0:
            return "Em andamento"
        return "Não iniciada"

    df_mpp["status"] = df_mpp.apply(_status_mpp, axis=1)
    df_mpp["atrasada"] = df_mpp["status"] == "Atrasada"

    return df_mpp[[
        "nome", "nivel", "fase", "pct_concluido", "duracao_dias",
        "inicio", "termino", "predecessoras", "recursos",
        "recursos_lista", "observacao", "status", "atrasada"
    ]].reset_index(drop=True)


# ---------------------------------------------------------------------------
# EXECUÇÃO DIRETA (teste)
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print("Uso: python processar_dados.py <arquivo.mpp>")
        sys.exit(1)
    df = carregar_dados(sys.argv[1])
    print(f"\n[OK] {len(df)} tarefas carregadas\n")
    print(df[["nome", "nivel", "fase", "pct_concluido", "status"]].to_string(index=False))
    print("\nColunas disponíveis:", list(df.columns))
