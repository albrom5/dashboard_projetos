"""
processar_dados.py
==================
Lê arquivos do Microsoft Project e retorna um DataFrame limpo
e enriquecido para uso no dashboard Streamlit.

Formatos suportados:
  - MPP  (arquivo nativo do MS Project)         -- requer Java 11+ e `mpxj`
  - CSV  (separador ;, encoding CP1252/Latin-1) -- exportação padrão PT-BR do MSP
  - XLSX (Excel)                                -- exportação nativa do MSP ou planilha manual
"""

import io
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO
# ---------------------------------------------------------------------------
# Prefer o .mpp nativo; cai para o CSV exportado se não encontrar
_BASE = Path(__file__).parent
ARQUIVO_PADRAO = (
    _BASE / "DataMesh_Fase2_v6_TESTE.mpp"
    if (_BASE / "DataMesh_Fase2_v6_TESTE.mpp").exists()
    else _BASE / "3__DataMesh_Fase2.csv"
)

# Mapa de dias da semana PT em abreviações usadas pelo MSP
DIAS_SEMANA = {"Seg": "Mon", "Ter": "Tue", "Qua": "Wed",
               "Qui": "Thu", "Sex": "Fri", "Sáb": "Sat", "Dom": "Sun"}

# Data de referência para calcular atraso
HOJE = date.today()


# ---------------------------------------------------------------------------
# FUNÇÕES AUXILIARES
# ---------------------------------------------------------------------------

def _normalizar_encoding(caminho: Path) -> str:
    """Tenta ler o arquivo com diferentes encodings comuns do Windows/MSP."""
    for enc in ("cp1252", "latin-1", "utf-8-sig", "utf-8"):
        try:
            texto = caminho.read_text(encoding=enc)
            return texto
        except (UnicodeDecodeError, UnicodeError):
            continue
    raise ValueError(f"Não foi possível decodificar o arquivo: {caminho}")


def _limpar_header(linhas: list[str]) -> list[str]:
    """
    O MSP às vezes quebra o cabeçalho em múltiplas linhas físicas.
    Esta função garante que a primeira linha seja o cabeçalho completo.
    """
    # Conta campos no cabeçalho (linha 0); se a linha 1 tem menos campos
    # do que o esperado, é continuação do cabeçalho
    cabecalho = linhas[0].rstrip("\n").rstrip("\r")
    ncols = cabecalho.count(";") + 1
    dados = []
    i = 1
    while i < len(linhas):
        linha = linhas[i].rstrip("\n").rstrip("\r")
        if linha.count(";") + 1 < ncols and not dados:
            # ainda é parte do cabeçalho
            cabecalho += " " + linha.strip()
            i += 1
        else:
            dados.append(linha)
            i += 1
    return [cabecalho] + dados


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


def _nivel_hierarquia(nome: str) -> int:
    """Conta espaços à esquerda e converte em nível (0 = raiz)."""
    espacos = len(nome) - len(nome.lstrip(" "))
    return espacos // 3  # MSP usa ~3 espaços por nível


def _extrair_fase(nome_limpo: str, nivel: int, fases: list[str]) -> str:
    """Heurística simples: fase é o item de nível 1."""
    return fases[-1] if fases else "Geral"


def _ler_excel(caminho: Path) -> pd.DataFrame:
    """
    Lê planilha Excel exportada do MS Project.
    Retorna DataFrame bruto com colunas como strings para alimentar
    o mesmo pipeline de normalização do CSV.
    """
    # engine='openpyxl' para .xlsx; tenta auto-detectar para .xls
    engine = "openpyxl" if caminho.suffix.lower() == ".xlsx" else None
    df = pd.read_excel(
        caminho,
        engine=engine,
        keep_default_na=False,
        # Não forçamos dtype=str para preservar datas e números nativos;
        # os converters individuais sabem lidar com ambos os tipos.
    )
    return df


# ---------------------------------------------------------------------------
# LEITURA NATIVA DE ARQUIVOS .MPP (MS Project)
# Estratégia: converte .mpp → JSON via mpxj.jar + Java subprocess (sem JPype).
# Requisitos: pip install mpxj  |  Java JRE instalado e no PATH
# ---------------------------------------------------------------------------

def _mpxj_jar_classpath() -> str:
    """
    Retorna o classpath com todos os JARs do pacote mpxj instalado.
    Lança FileNotFoundError se o pacote não estiver disponível.
    """
    import site

    candidatos = site.getsitepackages() + [site.getusersitepackages()]
    for sp in candidatos:
        lib_dir = Path(sp) / "mpxj" / "lib"
        if lib_dir.exists():
            jars = sorted(lib_dir.glob("*.jar"))
            if jars:
                return ";".join(str(j) for j in jars)  # ';' no Windows
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

def carregar_dados(caminho: Path | str = ARQUIVO_PADRAO) -> pd.DataFrame:
    """
    Carrega e processa arquivo do MS Project: nativo .mpp, CSV ou Excel.

    Retorna um DataFrame com as colunas:
        nome, nivel, fase, pct_concluido, duracao_dias,
        inicio, termino, predecessoras, recursos, recursos_lista,
        observacao, status, atrasada

    Para arquivos .mpp é necessário:
        - pip install mpxj
        - Java JRE/JDK 11+ instalado e no PATH
    """
    caminho = Path(caminho)
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    extensao = caminho.suffix.lower()

    # ── Caminho MPP (nativo MS Project) ─────────────────────────────────────
    if extensao == ".mpp":
        df_mpp = _ler_mpp(caminho)

        # Determina a fase (mesma lógica do caminho CSV/Excel)
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

    if extensao in (".xlsx", ".xls"):
        # ── Caminho Excel ────────────────────────────────────────────────
        df_raw = _ler_excel(caminho)

        # Converte todas as colunas do Excel para string, exceto as que
        # já são date/datetime (tratadas individualmente mais adiante)
        for col in df_raw.columns:
            if df_raw[col].dtype == object:
                df_raw[col] = df_raw[col].astype(str).str.strip()
    else:
        # ── Caminho CSV ──────────────────────────────────────────────────
        # 1. Leitura bruta
        texto = _normalizar_encoding(caminho)
        linhas = texto.splitlines()

        # 2. Normaliza header multi-linha
        linhas = _limpar_header(linhas)

        # 3. Lê com pandas
        conteudo = "\n".join(linhas)
        df_raw = pd.read_csv(
            io.StringIO(conteudo),
            sep=";",
            dtype=str,
            keep_default_na=False,
            on_bad_lines="skip",
        )

    # 4. Normaliza nomes de colunas
    df_raw.columns = (
        df_raw.columns
        .str.strip()
        .str.lower()
        .str.replace(r"\s+", "_", regex=True)
        .str.replace(r"[^a-z0-9_]", "", regex=True)
    )

    # Tenta mapear colunas conhecidas independente do encoding
    mapa_colunas = {}
    for col in df_raw.columns:
        col_lower = col
        if "nome" in col_lower and "tarefa" in col_lower:
            mapa_colunas[col] = "nome_raw"
        elif "conclu" in col_lower or "pct" in col_lower or "percent" in col_lower:
            mapa_colunas[col] = "pct_str"
        elif "dura" in col_lower:
            mapa_colunas[col] = "duracao_str"
        elif col_lower.startswith("in") and ("cio" in col_lower or "ic" in col_lower):
            mapa_colunas[col] = "inicio_str"
        elif "rmino" in col_lower or "rmino" in col_lower or col_lower.startswith("t"):
            mapa_colunas[col] = "termino_str"
        elif "predecessor" in col_lower:
            mapa_colunas[col] = "predecessoras"
        elif "recurso" in col_lower or "resource" in col_lower:
            mapa_colunas[col] = "recursos"
        elif "observa" in col_lower or "obs" in col_lower or "note" in col_lower:
            mapa_colunas[col] = "observacao"

    # Fallback: assume ordem padrão de exportação do MSP
    colunas_originais = list(df_raw.columns)
    if len(colunas_originais) >= 8 and "nome_raw" not in mapa_colunas.values():
        fallback = ["nome_raw", "pct_str", "duracao_str", "inicio_str",
                    "termino_str", "predecessoras", "recursos", "observacao"]
        mapa_colunas = {colunas_originais[i]: fallback[i]
                        for i in range(min(8, len(colunas_originais)))}

    df_raw = df_raw.rename(columns=mapa_colunas)

    # Garante que as colunas existem
    for col in ["nome_raw", "pct_str", "duracao_str", "inicio_str",
                "termino_str", "predecessoras", "recursos", "observacao"]:
        if col not in df_raw.columns:
            df_raw[col] = ""

    # 5. Remove linhas sem nome de tarefa
    df_raw = df_raw[df_raw["nome_raw"].str.strip() != ""].copy()

    # 6. Calcula nível hierárquico e nome limpo
    df_raw["nivel"] = df_raw["nome_raw"].apply(_nivel_hierarquia)
    df_raw["nome"] = df_raw["nome_raw"].str.strip()

    # 7. Determina a fase (nível 1 da hierarquia)
    fases_stack: list[str] = []
    faselist: list[str] = []
    for _, row in df_raw.iterrows():
        nivel = row["nivel"]
        nome = row["nome"]
        # Mantém a pilha de fases até o nível atual
        while len(fases_stack) > nivel:
            fases_stack.pop()
        fases_stack.append(nome)
        # A fase é o item de nível 1 (índice 1 na pilha)
        if len(fases_stack) > 1:
            faselist.append(fases_stack[1])
        else:
            faselist.append(nome)
    df_raw["fase"] = faselist

    # 8. Conversões
    df_raw["pct_concluido"] = df_raw["pct_str"].apply(_converter_percentual)
    df_raw["duracao_dias"] = df_raw["duracao_str"].apply(_converter_duracao)
    df_raw["inicio"] = df_raw["inicio_str"].apply(_converter_data)
    df_raw["termino"] = df_raw["termino_str"].apply(_converter_data)

    # 9. Status calculado
    def _status(row):
        pct = row["pct_concluido"]
        termino = row["termino"]
        if pct is None:
            return "Indefinido"
        if pct >= 100:
            return "Concluída"
        if termino and termino < HOJE and pct < 100:
            return "Atrasada"
        if pct > 0:
            return "Em andamento"
        return "Não iniciada"

    df_raw["status"] = df_raw.apply(_status, axis=1)
    df_raw["atrasada"] = df_raw["status"] == "Atrasada"

    # 10. Recursos como lista
    df_raw["recursos_lista"] = (
        df_raw["recursos"]
        .str.replace(r'["\']', "", regex=True)
        .str.split(r"[;,]")
        .apply(lambda lst: [r.strip() for r in lst if r.strip()] if isinstance(lst, list) else [])
    )

    # 11. Seleciona e ordena colunas finais
    df_final = df_raw[[
        "nome", "nivel", "fase", "pct_concluido", "duracao_dias",
        "inicio", "termino", "predecessoras", "recursos",
        "recursos_lista", "observacao", "status", "atrasada"
    ]].reset_index(drop=True)

    return df_final


# ---------------------------------------------------------------------------
# EXECUÇÃO DIRETA (teste)
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    df = carregar_dados()
    print(f"\n[OK] {len(df)} tarefas carregadas\n")
    print(df[["nome", "nivel", "fase", "pct_concluido", "status"]].to_string(index=False))
    print("\nColunas disponíveis:", list(df.columns))
