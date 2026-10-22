import argparse
import json
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

# Regex para validar o formato: "Nota Técnica_XXX_qualquer coisa.ext"
# Grupo 1 captura apenas os algarismos (XXX) entre os underscores "_"
REGEX_NOTA_TECNICA = re.compile(
    r"^Nota\s+T[ée]cnica_(\d+)_.+\.(pdf|docx|doc)$",
    re.IGNORECASE,
)


def normalizar_chave_nt(numero_str: str, ano: str = ANO_REFERENCIA) -> str:
    """
    Converte '0045' ou '45/2026' para o formato padronizado '45/2026'
    para garantir que zeros à esquerda não impeçam o matching.
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
    Lê apenas os nomes dos arquivos na pasta (sem abri-los).
    Retorna:
      - notas_validas: dict mapeando 'XXX/2026' -> lista de arquivos correspondentes
      - fora_formatacao: lista com os nomes dos arquivos fora do padrão
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

        # Ignora arquivos temporários/ocultos do Windows e do Word
        if nome_arquivo.startswith("~$") or nome_arquivo.lower() in ("desktop.ini", "thumbs.db"):
            continue

        match = REGEX_NOTA_TECNICA.match(nome_arquivo)
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

    # 1. Escanear os nomes dos arquivos na pasta de Notas Técnicas
    print(f"[1/4] Escaneando arquivos em: {pasta_notas}")
    notas_na_pasta, fora_formatacao = escanear_pasta_notas(pasta_notas, recursivo=recursivo)
    print(f" -> Notas Técnicas únicas identificadas (padrão válido): {len(notas_na_pasta)}")
    print(f" -> Arquivos fora da formatação: {len(fora_formatacao)}")

    # 2. Salvar o arquivo "fora_da_formatacao.txt"
    caminho_fora_fmt = pasta_saida / "fora_da_formatacao.txt"
    with open(caminho_fora_fmt, "w", encoding="utf-8") as f:
        if fora_formatacao:
            f.write("\n".join(fora_formatacao) + "\n")
    print(f"[2/4] Arquivo '{caminho_fora_fmt.name}' gerado ({len(fora_formatacao)} registro(s)).")

    # 3. Carregar e validar a planilha do Excel
    if not caminho_planilha.exists():
        raise FileNotFoundError(f"A planilha Excel não foi encontrada: {caminho_planilha}")

    print(f"[3/4] Lendo planilha Excel: {caminho_planilha.name}...")
    df = pd.read_excel(caminho_planilha, sheet_name=nome_aba)

    # Remove espaços extras invisíveis nos títulos das colunas da planilha
    df.columns = [str(col).strip() for col in df.columns]

    colunas_faltantes = [col for col in COLUNAS_OBRIGATORIAS if col not in df.columns]
    if colunas_faltantes:
        raise ValueError(
            f"As seguintes colunas não foram encontradas na planilha: {colunas_faltantes}\n"
            f"Colunas disponíveis na planilha: {list(df.columns)}"
        )

    # Cria uma coluna auxiliar normalizada para cruzar 'XXX/2026' com precisão
    df["_chave_nt"] = df["Nota Técnica"].dropna().astype(str).apply(
        lambda val: normalizar_chave_nt(val, ANO_REFERENCIA)
    )

    # Filtra apenas as linhas da planilha cuja Nota Técnica (XXX/2026) existe na pasta
    chaves_encontradas_na_pasta = set(notas_na_pasta.keys())
    df_matched = df[df["_chave_nt"].isin(chaves_encontradas_na_pasta)].copy()

    # Seleciona exatamente as 4 colunas pedidas e ordena por Nota Técnica
    df_final = df_matched[COLUNAS_OBRIGATORIAS].copy()

    # Verifica quais Notas Técnicas estavam na pasta mas NÃO foram achadas na planilha
    chaves_na_planilha = set(df_matched["_chave_nt"].unique())
    chaves_sem_match = sorted(chaves_encontradas_na_pasta - chaves_na_planilha)

    # 4. Salvar os arquivos finais formatados (Excel, CSV e JSON para o Playwright)
    caminho_excel_saida = pasta_saida / "dados_pareados_opee.xlsx"
    caminho_csv_saida = pasta_saida / "dados_pareados_opee.csv"
    caminho_json_saida = pasta_saida / "dados_pareados_opee.json"

    df_final.to_excel(caminho_excel_saida, index=False)
    df_final.to_csv(caminho_csv_saida, index=False, sep=";", encoding="utf-8-sig")

    # Salva também em JSON estruturado (ótimo para o Playwright ler na próxima etapa)
    registros_json = df_final.where(pd.notnull(df_final), None).to_dict(orient="records")
    with open(caminho_json_saida, "w", encoding="utf-8") as f:
        json.dump(registros_json, f, ensure_ascii=False, indent=4)

    # Se houver notas na pasta que não existem como XXX/2026 na planilha, salva um aviso
    if chaves_sem_match:
        caminho_sem_match = pasta_saida / "notas_nao_encontradas_na_planilha.txt"
        with open(caminho_sem_match, "w", encoding="utf-8") as f:
            for chave in chaves_sem_match:
                arquivos_orig = ", ".join(notas_na_pasta[chave])
                f.write(f"{chave} -> Arquivos: {arquivos_orig}\n")

    print("\n" + "=" * 70)
    print(" RESUMO DO MATCHING CONCLUÍDO")
    print("=" * 70)
    print(f" - Total de registros pareados com sucesso : {len(df_final)}")
    print(f" - Notas da pasta encontradas na planilha  : {len(chaves_na_planilha)}/{len(chaves_encontradas_na_pasta)}")
    print(f" - Arquivos listados em fora_da_formatacao : {len(fora_formatacao)}")
    if chaves_sem_match:
        print(f" - [AVISO] Notas na pasta sem '{ANO_REFERENCIA}' na tabela: {len(chaves_sem_match)} (veja notas_nao_encontradas_na_planilha.txt)")
    print("-" * 70)
    print(f" Arquivos gerados em: {pasta_saida}")
    print(f"   1. {caminho_excel_saida.name} (Planilha pronta para visualização)")
    print(f"   2. {caminho_csv_saida.name}   (CSV separado por ';' compatível com Excel PT-BR)")
    print(f"   3. {caminho_json_saida.name}  (JSON pronto para o robô do Playwright)")
    print(f"   4. {caminho_fora_fmt.name}  (Arquivos fora do padrão)")
    print("=" * 70)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Cruza os nomes dos arquivos da pasta de Notas Técnicas com a planilha Excel."
    )
    parser.add_argument(
        "--pasta", "-p",
        type=Path,
        default=PASTA_NOTAS_PADRAO,
        help="Caminho da pasta onde estão os PDFs e DOCXs das Notas Técnicas.",
    )
    parser.add_argument(
        "--planilha", "-x",
        type=Path,
        default=PLANILHA_EXCEL_PADRAO,
        help="Caminho completo do arquivo Excel (.xlsx).",
    )
    parser.add_argument(
        "--saida", "-s",
        type=Path,
        default=PASTA_SAIDA_PADRAO,
        help="Pasta onde os arquivos de resultado serão salvos.",
    )
    parser.add_argument(
        "--aba", "-a",
        default=0,
        help="Nome ou índice da aba da planilha Excel (padrão: primeira aba).",
    )
    parser.add_argument(
        "--recursivo", "-r",
        action="store_true",
        help="Ative se as Notas Técnicas estiverem divididas dentro de subpastas.",
    )

    args = parser.parse_args()
    executar_matching(
        pasta_notas=args.pasta,
        caminho_planilha=args.planilha,
        pasta_saida=args.saida,
        nome_aba=args.aba,
        recursivo=args.recursivo,
    )
