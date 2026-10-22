import argparse
import re
from pathlib import Path
import pandas as pd

# ==============================================================================
# 1. CONFIGURAÇÃO DE CAMINHOS PADRÃO (Altere para os caminhos reais do seu PC)
# ==============================================================================
PASTA_NOTAS_PADRAO = Path(r"C:\automacao_opee\pasta_notas_tecnicas")
PLANILHA_EXCEL_PADRAO = Path(r"C:\automacao_opee\planilha_empresas.xlsx")
PASTA_SAIDA_PADRAO = Path(r"C:\automacao_opee\local\05_Matching_Notas")

ANO_REFERENCIA = "2026"

# Nomes exatos das colunas na planilha Excel
COLUNAS_OBRIGATORIAS = [
    "Distribuidora",
    "Codigo do Projeto",
    "Nota Técnica",
    "Valor Aprovado",
]

# Captura o PRIMEIRO número (de 1 ou mais algarismos) que estiver entre "_" e "_"
REGEX_NUMERO_ENTRE_UNDERSCORES = re.compile(r"_\s*(\d+)\s*_")


def normalizar_chave_nt(numero_str: str, ano: str = ANO_REFERENCIA) -> str:
    """
    Padroniza '0045' ou '45/2026' para '45/2026', garantindo o match
    mesmo se houver zeros à esquerda no arquivo ou na planilha.
    """
    texto = str(numero_str).strip()
    match = re.match(r"^0*(\d+)\s*/\s*(\d{4})$", texto)
    if match:
        num = match.group(1) or "0"
        return f"{int(num)}/{match.group(2)}"

    if texto.isdigit():
        return f"{int(texto)}/{ano}"

    return texto.upper()


def escanear_pasta_notas(pasta_notas: Path, recursivo: bool = False) -> tuple[dict[str, list[str]], list[str]]:
    """
    Busca o primeiro número entre '_ _' no nome de cada arquivo (.pdf, .docx, .doc).
    Retorna:
      - notas_validas: dict mapeando 'XXX/2026' -> lista de arquivos daquela nota
      - fora_formatacao: lista com os nomes dos arquivos sem número entre '_ _'
    """
    if not pasta_notas.exists() or not pasta_notas.is_dir():
        raise FileNotFoundError(f"A pasta de Notas Técnicas não foi encontrada: {pasta_notas}")

    notas_validas: dict[str, list[str]] = {}
    fora_formatacao: list[str] = []

    iterador = pasta_notas.rglob("*") if recursivo else pasta_notas.iterdir()

    for item in sorted(iterador):
        if not item.is_file():
            continue

        nome_arquivo = item.name

        # Ignora arquivos temporários do Word (~$) e arquivos de sistema do Windows
        if nome_arquivo.startswith("~$") or nome_arquivo.lower() in ("desktop.ini", "thumbs.db"):
            continue

        # Considera apenas documentos PDF e Word
        if item.suffix.lower() not in (".pdf", ".docx", ".doc"):
            continue

        # Procura o primeiro número entre "_" e "_" no nome do arquivo (sem a extensão)
        match = REGEX_NUMERO_ENTRE_UNDERSCORES.search(item.stem)
        if match:
            numero_xxx = match.group(1)
            chave_normalizada = normalizar_chave_nt(numero_xxx, ANO_REFERENCIA)
            notas_validas.setdefault(chave_normalizada, []).append(nome_arquivo)
        else:
            fora_formatacao.append(nome_arquivo)

    return notas_validas, fora_formatacao


def executar_matching(
    pasta_notas: Path,
    caminho_planilha: Path,
    pasta_saida: Path,
    nome_aba: str | int = 0,
    recursivo: bool = False,
):
    pasta_saida.mkdir(parents=True, exist_ok=True)

    print("=" * 70)
    print(" INICIANDO MATCHING: PASTA DE NOTAS TÉCNICAS x PLANILHA EXCEL")
    print("=" * 70)

    # 1. Escanear pasta de Notas Técnicas
    print(f"[1/4] Escaneando arquivos em: {pasta_notas}")
    notas_na_pasta, fora_formatacao = escanear_pasta_notas(pasta_notas, recursivo=recursivo)
    print(f" -> Notas Técnicas únicas extraídas (_XXX_): {len(notas_na_pasta)}")
    print(f" -> Arquivos fora da formatação: {len(fora_formatacao)}")

    # 2. Salvar "fora_da_formatacao.txt"
    caminho_fora_fmt = pasta_saida / "fora_da_formatacao.txt"
    with open(caminho_fora_fmt, "w", encoding="utf-8") as f:
        if fora_formatacao:
            f.write("\n".join(fora_formatacao) + "\n")
    print(f"[2/4] Arquivo '{caminho_fora_fmt.name}' gerado.")

    # 3. Ler planilha Excel e cruzar pelo campo "Nota Técnica" == "XXX/2026"
    if not caminho_planilha.exists():
        raise FileNotFoundError(f"A planilha Excel não foi encontrada: {caminho_planilha}")

    print(f"[3/4] Lendo planilha Excel: {caminho_planilha.name}...")
    df = pd.read_excel(caminho_planilha, sheet_name=nome_aba)
    df.columns = [str(col).strip() for col in df.columns]

    colunas_faltantes = [col for col in COLUNAS_OBRIGATORIAS if col not in df.columns]
    if colunas_faltantes:
        raise ValueError(
            f"Colunas não encontradas na planilha: {colunas_faltantes}\n"
            f"Colunas existentes: {list(df.columns)}"
        )

    df["_chave_nt"] = df["Nota Técnica"].dropna().astype(str).apply(
        lambda val: normalizar_chave_nt(val, ANO_REFERENCIA)
    )

    chaves_encontradas_na_pasta = set(notas_na_pasta.keys())
    df_matched = df[df["_chave_nt"].isin(chaves_encontradas_na_pasta)].copy()
    df_final = df_matched[COLUNAS_OBRIGATORIAS].copy()

    # 4. Salvar os dados pareados em TXT ("dados_pareados_opee.txt")
    caminho_txt_saida = pasta_saida / "dados_pareados_opee.txt"
    with open(caminho_txt_saida, "w", encoding="utf-8") as f:
        # Cabeçalho
        f.write(" | ".join(COLUNAS_OBRIGATORIAS) + "\n")
        f.write("-" * 80 + "\n")

        # Linhas pareadas
        for _, linha in df_final.iterrows():
            distribuidora = str(linha["Distribuidora"]).strip()
            cod_projeto = str(linha["Codigo do Projeto"]).strip()
            nota_tecnica = str(linha["Nota Técnica"]).strip()
            valor_aprovado = str(linha["Valor Aprovado"]).strip()
            f.write(f"{distribuidora} | {cod_projeto} | {nota_tecnica} | {valor_aprovado}\n")

    print(f"[4/4] Arquivo '{caminho_txt_saida.name}' gerado com {len(df_final)} linha(s)!")

    # Resumo final no terminal
    chaves_na_planilha = set(df_matched["_chave_nt"].unique())
    print("\n" + "=" * 70)
    print(" RESUMO DO MATCHING")
    print("=" * 70)
    print(f" - Notas únicas identificadas na pasta    : {len(chaves_encontradas_na_pasta)}")
    print(f" - Notas encontradas na planilha (/2026)  : {len(chaves_na_planilha)}")
    print(f" - Total de projetos gravados no TXT      : {len(df_final)}")
    print(f" - Arquivos em 'fora_da_formatacao.txt'   : {len(fora_formatacao)}")
    print(f" - Salvo em: {pasta_saida}")
    print("=" * 70)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Extrai o número _XXX_ das Notas Técnicas e cruza com a planilha Excel gerando um TXT."
    )
    parser.add_argument("--pasta", "-p", type=Path, default=PASTA_NOTAS_PADRAO, help="Pasta das Notas Técnicas.")
    parser.add_argument("--planilha", "-x", type=Path, default=PLANILHA_EXCEL_PADRAO, help="Arquivo Excel (.xlsx).")
    parser.add_argument("--saida", "-s", type=Path, default=PASTA_SAIDA_PADRAO, help="Pasta de saída dos TXTs.")
    parser.add_argument("--aba", "-a", default=0, help="Nome ou número da aba do Excel (padrão: 0).")
    parser.add_argument("--recursivo", "-r", action="store_true", help="Busca também dentro de subpastas.")

    args = parser.parse_args()
    executar_matching(
        pasta_notas=args.pasta,
        caminho_planilha=args.planilha,
        pasta_saida=args.saida,
        nome_aba=args.aba,
        recursivo=args.recursivo,
    )
