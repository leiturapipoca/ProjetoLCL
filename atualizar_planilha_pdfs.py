import argparse
import re
import unicodedata
from pathlib import Path
import pandas as pd
import openpyxl
import pdfplumber

from config import ONEDRIVE_BASE_DIR, inicializar_pastas
from controle_downloads import chave_projeto

# ==============================================================================
# 1. CONFIGURAÇÕES DE CAMINHOS PADRÃO
# ==============================================================================
PASTA_NOTAS_PADRAO = Path(r"C:\automacao_opee\pasta_notas_tecnicas")
PLANILHA_EXCEL_PADRAO = Path(r"C:\automacao_opee\planilha_empresas.xlsx")
ARQUIVO_WARNINGS_PADRAO = ONEDRIVE_BASE_DIR / "warnings_notas_tecnicas.txt"

ANO_PADRAO = "2026"

# Regex separadas: aceita apenas PE e identifica outros (PG, PD) para ignorar
REGEX_PROJETO_PE = re.compile(r"\bPE-\d+-\d+\s*/\s*\d{4}\b", re.IGNORECASE)
REGEX_PROJETO_OUTROS = re.compile(r"\b(?!PE-)[A-Z]{2,3}-\d+-\d+\s*/\s*\d{4}\b", re.IGNORECASE)

# Regex para "NOTA TÉCNICA Nº XX/2026" logo após a logo no PDF
REGEX_NT_CABECALHO = re.compile(
    r"NOTA\s+T[ÉE]CNICA\s*(?:N[º°o.]*)?\s*([0-9X]+)\s*/\s*(\d{4})",
    re.IGNORECASE,
)

# Regex para o primeiro número entre "_" e "_" no nome do arquivo (_XXX_)
REGEX_NUMERO_ARQUIVO = re.compile(r"_\s*(\d+)\s*_")


def normalizar_nome_coluna(texto: str) -> str:
    if not texto:
        return ""
    sem_acento = "".join(
        c for c in unicodedata.normalize("NFD", str(texto))
        if unicodedata.category(c) != "Mn"
    )
    return re.sub(r"[^a-zA-Z0-9]+", " ", sem_acento).strip().lower()


def encontrar_coluna(colunas_planilha: list[str], nomes_possiveis: list[str]) -> str:
    mapa = {normalizar_nome_coluna(c): c for c in colunas_planilha}
    for nome in nomes_possiveis:
        if nome in colunas_planilha:
            return nome
        norm = normalizar_nome_coluna(nome)
        if norm in mapa:
            return mapa[norm]
    raise ValueError(f"Coluna não encontrada. Esperado um de {nomes_possiveis} em {colunas_planilha}")


def campo_esta_vazio(valor) -> bool:
    if pd.isna(valor) or valor is None:
        return True
    texto = str(valor).strip()
    return not bool(texto) or texto.lower() in ("nan", "none", "-", "n/a", "não", "nao")


def limpar_celula_pdf(valor_celula: str | None) -> str:
    if not valor_celula:
        return ""
    return re.sub(r"\s+", " ", str(valor_celula)).strip()


def eh_zero_reais(texto_celula: str) -> bool:
    """Retorna True se o valor na coluna de glosa equivaler a R$ 0,00."""
    if not texto_celula or not texto_celula.strip():
        return True
    limpo = re.sub(r"\s+", "", texto_celula).upper()
    match_num = re.search(r"(\d[\d.,]*)", limpo)
    if not match_num:
        return True
    num_str = match_num.group(1).replace(".", "").replace(",", ".")
    try:
        return abs(float(num_str)) < 0.005
    except ValueError:
        return limpo in ("R$0,00", "0,00", "0", "-")


def extrair_codigo_projeto(texto_completo: str) -> tuple[str | None, str | None]:
    """
    Retorna (codigo_pe, codigo_ignorado).
    Se o projeto for PG, PD, etc., retorna (None, codigo_outro) para ser ignorado.
    """
    match_bloco = re.search(
        r"(?:I\s*[-–—.]?\s*DO\s+OBJETIVO)(.*?)(?:II\s*[-–—.]|\n\s*2\s*[-–—.]|$)",
        texto_completo,
        re.IGNORECASE | re.DOTALL,
    )
    if match_bloco:
        trecho_objetivo = match_bloco.group(1)
        codigos_pe = REGEX_PROJETO_PE.findall(trecho_objetivo)
        if codigos_pe:
            return chave_projeto(codigos_pe[-1]), None
        codigos_outros = REGEX_PROJETO_OUTROS.findall(trecho_objetivo)
        if codigos_outros:
            return None, codigos_outros[-1].upper()

    # Fallback no texto completo do PDF
    codigos_pe_geral = REGEX_PROJETO_PE.findall(texto_completo)
    if codigos_pe_geral:
        return chave_projeto(codigos_pe_geral[0]), None

    codigos_outros_geral = REGEX_PROJETO_OUTROS.findall(texto_completo)
    if codigos_outros_geral:
        return None, codigos_outros_geral[0].upper()

    return None, None


def extrair_nota_tecnica(texto_primeira_pagina: str, texto_completo: str, nome_arquivo: str) -> tuple[str | None, bool]:
    match_pdf = REGEX_NT_CABECALHO.search(texto_primeira_pagina) or REGEX_NT_CABECALHO.search(texto_completo)

    esqueceu_xx = False
    if match_pdf:
        num_bruto = match_pdf.group(1).strip().upper()
        ano_pdf = match_pdf.group(2).strip()
        if "X" not in num_bruto and num_bruto.isdigit():
            return f"{int(num_bruto)}/{ano_pdf}", False
        else:
            esqueceu_xx = True

    stem = Path(nome_arquivo).stem
    match_arq = REGEX_NUMERO_ARQUIVO.search(stem)
    if match_arq:
        num_arq = int(match_arq.group(1))
        return f"{num_arq}/{ANO_PADRAO}", esqueceu_xx

    return None, esqueceu_xx


def extrair_valor_por_total_pdf(
    tabelas_brutas: list[list[list[str | None]]],
    texto_completo: str,
) -> tuple[str | None, bool, str]:
    """
    Procura a linha contendo 'Total' na estrutura: Total | valor | glosado.
    - Coluna imediatamente após 'Total' -> Valor Aprovado.
    - Coluna seguinte (se existir e != R$ 0,00) -> Warning 'Valor glosado'.
    """
    # 1. Busca estruturada nas tabelas extraídas pelo pdfplumber
    for tab in tabelas_brutas:
        for linha in tab:
            if not linha:
                continue
            # Filtra células vazias/None geradas por mesclagem de colunas no PDF
            celulas = [limpar_celula_pdf(c) for c in linha if limpar_celula_pdf(c)]
            for idx, texto_cel in enumerate(celulas):
                if re.search(r"\bTotal\b", texto_cel, re.IGNORECASE):
                    if idx + 1 < len(celulas):
                        valor_aprovado = celulas[idx + 1]
                        if not re.search(r"\d", valor_aprovado):
                            continue

                        texto_glosa = ""
                        teve_glosa = False
                        if idx + 2 < len(celulas):
                            texto_glosa = celulas[idx + 2]
                            if re.search(r"\d", texto_glosa) and not eh_zero_reais(texto_glosa):
                                teve_glosa = True

                        return valor_aprovado, teve_glosa, texto_glosa

    # 2. Fallback por linha de texto (caso a tabela no PDF não tenha bordas visíveis)
    padrao_linha_total = re.compile(
        r"\bTotal\b[^\d\n]*((?:R\$\s*)?\d{1,3}(?:\.\d{3})*,\d{2})(?:[^\d\n]+((?:R\$\s*)?\d{1,3}(?:\.\d{3})*,\d{2}))?",
        re.IGNORECASE,
    )
    match_txt = padrao_linha_total.search(texto_completo)
    if match_txt:
        valor_aprovado = match_txt.group(1).strip()
        texto_glosa = (match_txt.group(2) or "").strip()
        teve_glosa = bool(texto_glosa) and not eh_zero_reais(texto_glosa)
        return valor_aprovado, teve_glosa, texto_glosa

    return None, False, ""


def analisar_pdf(caminho_pdf: Path) -> dict:
    textos_paginas = []
    todas_tabelas = []

    with pdfplumber.open(str(caminho_pdf)) as pdf:
        for pagina in pdf.pages:
            txt = pagina.extract_text() or ""
            if txt.strip():
                textos_paginas.append(txt)

            tabelas_pag = pagina.extract_tables() or []
            if tabelas_pag:
                todas_tabelas.extend(tabelas_pag)

    texto_primeira_pagina = textos_paginas[0] if textos_paginas else ""
    texto_completo = "\n".join(textos_paginas)

    codigo_pe, codigo_ignorado = extrair_codigo_projeto(texto_completo)
    if codigo_ignorado and not codigo_pe:
        return {
            "arquivo": caminho_pdf.name,
            "ignorado_nao_pe": True,
            "codigo_ignorado": codigo_ignorado,
        }

    nota_tecnica, esqueceu_xx = extrair_nota_tecnica(texto_primeira_pagina, texto_completo, caminho_pdf.name)
    valor_aprovado, teve_glosa, texto_glosa = extrair_valor_por_total_pdf(todas_tabelas, texto_completo)

    return {
        "arquivo": caminho_pdf.name,
        "ignorado_nao_pe": False,
        "codigo_projeto": codigo_pe,
        "nota_tecnica": nota_tecnica,
        "valor_aprovado": valor_aprovado,
        "esqueceu_xx": esqueceu_xx,
        "teve_glosa": teve_glosa,
        "texto_glosa": texto_glosa,
    }


def executar_atualizacao_pdfs(
    pasta_notas: Path,
    caminho_planilha: Path,
    caminho_warnings: Path,
    nome_aba: str | int = 0,
    recursivo: bool = False,
):
    inicializar_pastas()

    if not pasta_notas.exists():
        raise FileNotFoundError(f"Pasta de Notas Técnicas não encontrada: {pasta_notas}")
    if not caminho_planilha.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {caminho_planilha}")

    print("=" * 75)
    print(" LENDO ARQUIVOS .PDF (APENAS PE) E ATUALIZANDO PLANILHA")
    print("=" * 75)

    iterador = pasta_notas.rglob("*") if recursivo else pasta_notas.iterdir()
    arquivos_pdf = sorted(
        arq for arq in iterador
        if arq.is_file() and arq.suffix.lower() == ".pdf"
    )

    print(f"[1/4] Encontrados {len(arquivos_pdf)} arquivo(s) .pdf na pasta.")

    dados_extraidos: dict[str, dict] = {}
    warnings: list[str] = []
    ignorados_nao_pe = 0

    for idx, arq in enumerate(arquivos_pdf, start=1):
        try:
            info = analisar_pdf(arq)
        except Exception as erro:
            warnings.append(f"[ERRO AO LER PDF] Arquivo: {arq.name} | Erro: {erro}")
            continue

        if info.get("ignorado_nao_pe"):
            ignorados_nao_pe += 1
            print(f"   [{idx}/{len(arquivos_pdf)}] [IGNORADO - NÃO É PE] {info['codigo_ignorado']} ({arq.name})")
            continue

        cod = info["codigo_projeto"]
        nt = info["nota_tecnica"]
        val = info["valor_aprovado"]

        if info["esqueceu_xx"]:
            warnings.append(
                f"-Esqueceu de mudar na própria nota técnica- | Projeto: {cod or 'N/A'} | "
                f"NT atribuída pelo nome: {nt or 'N/A'} | Arquivo: {arq.name}"
            )

        if info["teve_glosa"]:
            warnings.append(
                f"Valor glosado | Projeto: {cod or 'N/A'} | NT: {nt or 'N/A'} | "
                f"Valor glosado: {info['texto_glosa']} | Valor aprovado (após Total): {val} | Arquivo: {arq.name}"
            )

        if not cod:
            warnings.append(f"[PROJETO PE NÃO IDENTIFICADO] Regex PE-XXXXX-XXXX/XXXX não encontrada em: {arq.name}")
            continue

        dados_extraidos[cod] = info
        print(f"   [{idx}/{len(arquivos_pdf)}] {cod} -> NT: {nt} | Valor: {val}")

    # 2. Salvar arquivo de warnings
    caminho_warnings.parent.mkdir(parents=True, exist_ok=True)
    with open(caminho_warnings, "w", encoding="utf-8") as f:
        if warnings:
            f.write("\n".join(warnings) + "\n")
    print(f"\n[2/4] Arquivo de warnings salvo em '{caminho_warnings.name}' ({len(warnings)} ocorrência(s)).")

    # 3. Ler planilha e verificar atualizações (apenas projetos PE)
    print(f"[3/4] Analisando planilha '{caminho_planilha.name}'...")
    df = pd.read_excel(caminho_planilha, sheet_name=nome_aba)
    colunas_originais = [str(c).strip() for c in df.columns]
    df.columns = colunas_originais

    col_cod = encontrar_coluna(colunas_originais, ["Codigo do Projeto", "Código do Projeto"])
    col_aneel = encontrar_coluna(colunas_originais, ["Avaliado Pela ANEEL"])
    col_nt = encontrar_coluna(colunas_originais, ["Nota Técnica", "Nota Tecnica"])
    col_valor = encontrar_coluna(colunas_originais, ["Valor Aprovado"])

    idx_col_aneel = colunas_originais.index(col_aneel) + 1
    idx_col_nt = colunas_originais.index(col_nt) + 1
    idx_col_valor = colunas_originais.index(col_valor) + 1

    amostra_aneel = df[col_aneel].dropna().astype(str).str.strip().tolist()
    texto_sim = "SIM" if "SIM" in amostra_aneel else "Sim"

    atualizacoes = []

    for idx_df, row in df.iterrows():
        cod_bruto = str(row[col_cod]).strip() if pd.notna(row[col_cod]) else ""
        if not cod_bruto:
            continue

        try:
            chave = chave_projeto(cod_bruto)
        except ValueError:
            continue

        # Ignora se não for PE-
        if not chave.upper().startswith("PE-"):
            continue

        if chave not in dados_extraidos:
            continue

        info_pdf = dados_extraidos[chave]
        nt_pdf = info_pdf["nota_tecnica"]
        val_pdf = info_pdf["valor_aprovado"]

        aneel_atual = str(row[col_aneel]).strip().upper() if pd.notna(row[col_aneel]) else ""
        nt_vazia = campo_esta_vazio(row[col_nt])
        valor_vazio = campo_esta_vazio(row[col_valor])

        era_nao = aneel_atual in ("NÃO", "NAO", "N", "")
        era_sim_incompleto = aneel_atual in ("SIM", "S") and (nt_vazia or valor_vazio)

        if era_nao or era_sim_incompleto:
            linha_excel = idx_df + 2
            novo_aneel = texto_sim if era_nao else None
            nova_nt = nt_pdf if (era_nao or nt_vazia) and nt_pdf else None
            novo_val = val_pdf if (era_nao or valor_vazio) and val_pdf else None
            atualizacoes.append((linha_excel, chave, novo_aneel, nova_nt, novo_val))

    # 4. Gravar alterações na planilha Excel
    print(f"[4/4] Aplicando atualizações em {len(atualizacoes)} linha(s) da planilha...")
    if atualizacoes:
        while True:
            try:
                wb = openpyxl.load_workbook(caminho_planilha)
                ws = wb[nome_aba] if isinstance(nome_aba, str) else wb.worksheets[int(nome_aba)]

                for linha_excel, chave, novo_aneel, nova_nt, novo_val in atualizacoes:
                    if novo_aneel is not None:
                        ws.cell(row=linha_excel, column=idx_col_aneel, value=novo_aneel)
                    if nova_nt is not None:
                        ws.cell(row=linha_excel, column=idx_col_nt, value=nova_nt)
                    if novo_val is not None:
                        ws.cell(row=linha_excel, column=idx_col_valor, value=novo_val)

                wb.save(caminho_planilha)
                wb.close()
                break
            except PermissionError:
                print("\n [!] ATENÇÃO: A planilha está aberta no Excel!")
                input(" >>> Feche a planilha no Excel e pressione ENTER para salvar... ")

    print("\n" + "=" * 75)
    print(" RESUMO FINAL (LEITURA DE PDFS)")
    print("=" * 75)
    print(f" - Projetos PE válidos extraídos                  : {len(dados_extraidos)}")
    print(f" - Projetos ignorados (PG, PD, etc.)              : {ignorados_nao_pe}")
    print(f" - Projetos atualizados na planilha Excel         : {len(atualizacoes)}")
    print(f" - Ocorrências salvas no arquivo de warnings      : {len(warnings)}")
    print("=" * 75)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Extrai Projeto (apenas PE), Nota Técnica e Valor (após Total) dos arquivos .pdf e atualiza a planilha."
    )
    parser.add_argument("--pasta", "-p", type=Path, default=PASTA_NOTAS_PADRAO, help="Pasta com as Notas Técnicas (.pdf).")
    parser.add_argument("--planilha", "-x", type=Path, default=PLANILHA_EXCEL_PADRAO, help="Planilha Excel (.xlsx).")
    parser.add_argument("--warnings", "-w", type=Path, default=ARQUIVO_WARNINGS_PADRAO, help="Arquivo TXT de warnings.")
    parser.add_argument("--aba", "-a", default=0, help="Nome ou número da aba da planilha (padrão: 0).")
    parser.add_argument("--recursivo", "-r", action="store_true", help="Busca .pdf também em subpastas.")

    args = parser.parse_args()
    executar_atualizacao_pdfs(
        pasta_notas=args.pasta,
        caminho_planilha=args.planilha,
        caminho_warnings=args.warnings,
        nome_aba=args.aba,
        recursivo=args.recursivo,
    )
