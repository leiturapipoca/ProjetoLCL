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
PASTA_DESPACHOS_PADRAO = Path(r"C:\automacao_opee\pasta_despachos")
PLANILHA_EXCEL_PADRAO = Path(r"C:\automacao_opee\planilha_empresas.xlsx")
ARQUIVO_WARNINGS_PADRAO = ONEDRIVE_BASE_DIR / "warning_despachos_pdf.txt"

ANO_PADRAO = "2026"

# Regex para "DESPACHO Nº XXXX" (suporta \u00ba 'º', \u00b0 '°', sublinhados e pontos no número)
REGEX_DESPACHO_TEXTO = re.compile(
    r"DESPACHO\s*(?:N[\s_º°oª\u00ba\u00b0.:º\-–—]*)?\s*(\d[\d.]*)\s*(?:/\s*(20\d{2}))?",
    re.IGNORECASE,
)

# Regex fallback para o nome do arquivo: "Despacho_XXXX_..."
REGEX_DESPACHO_ARQUIVO = re.compile(r"Despacho[\s_-]*(\d+)", re.IGNORECASE)
REGEX_NUMERO_ENTRE_UNDERSCORES = re.compile(r"_\s*(\d+)\s*_")

# Regex para Código PE (aceita barra ou traço antes do ano)
REGEX_PROJETO_PE = re.compile(r"\bPE-\d+-\d+\s*[/_–—-]\s*\d{4}\b", re.IGNORECASE)

# Regex para o valor no 2º parágrafo: "R$ XXXXX (...)"
REGEX_VALOR_COM_PARENTESES = re.compile(
    r"R\$\s*(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})\s*\(",
    re.IGNORECASE,
)

# Fallback caso o PDF tenha engolido o parêntese ou colocado espaço atípico
REGEX_VALOR_BR = re.compile(r"R\$\s*(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})", re.IGNORECASE)


def normalizar_texto_pdf(texto: str) -> str:
    """
    Normaliza acentos decompostos (NFC), converte hífens especiais do PDF para '-'
    e corrige espaços internos em números ('150. 000, 00' -> '150.000,00').
    """
    if not texto:
        return ""
    txt = unicodedata.normalize("NFC", str(texto))
    # Converte hífens invisíveis ou não-separáveis do PDF em hífen comum '-'
    txt = re.sub(r"[\u2010\u2011\u2012\u2013\u2014\u2212\u00ad]", "-", txt)
    # Une espaçamentos de fonte em números monetários
    txt = re.sub(r"(\d)\s*([.,])\s*(\d)", r"\1\2\3", txt)
    return txt


def colar_quebras_de_codigo(texto: str) -> str:
    """Remove quebras de linha no meio do código PE-XXXXX-XXXX/XXXX."""
    sem_quebra_traco = re.sub(r"([A-Z0-9])\s*[-–—_]\s*(\d)", r"\1-\2", texto, flags=re.IGNORECASE)
    sem_quebra_barra = re.sub(r"(\d)\s*[/_–—-]\s*(\d{4})\b", r"\1/\2", sem_quebra_traco)
    return sem_quebra_barra


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


def converter_para_float_br(valor_bruto) -> float | None:
    """
    Converte 'R$ 1.250.430,50', 'R$ 0,00' ou float para número float do Python (1250430.50).
    """
    if valor_bruto is None or pd.isna(valor_bruto):
        return None
    if isinstance(valor_bruto, (int, float)):
        return float(valor_bruto)

    texto = normalizar_texto_pdf(str(valor_bruto)).strip()
    match = re.search(r"(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2}|\d+(?:\.\d+)?)", texto)
    if not match:
        return None

    num_str = match.group(1)
    if "," in num_str:
        num_str = num_str.replace(".", "").replace(",", ".")
    try:
        return float(num_str)
    except ValueError:
        return None


def formatar_moeda_br(valor_float: float | None) -> str:
    """Formata um float (1250430.5) para string no padrão 'R$ 1.250.430,50'."""
    if valor_float is None:
        return "N/A"
    s = f"{valor_float:,.2f}"
    s_br = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s_br}"


def eh_zero_reais(valor) -> bool:
    """Retorna True se o valor na planilha for R$ 0,00, 0 ou estiver vazio."""
    if valor is None or pd.isna(valor):
        return True
    val_float = converter_para_float_br(valor)
    if val_float is not None:
        return abs(val_float) < 0.005
    limpo = re.sub(r"[R$\s]", "", str(valor)).upper()
    return limpo in ("", "0,00", "0.00", "0", "-", "NAN", "NONE")


def extrair_numero_despacho(texto_completo: str, nome_arquivo: str) -> str | None:
    """
    1ª tentativa: Busca 'DESPACHO Nº XXXX' nas primeiras linhas do PDF.
    2ª tentativa (Fallback): Busca no padrão 'Despacho_XXXX_...' do nome do arquivo.
    Retorna sempre no formato 'XXXX/2026'.
    """
    primeiras_linhas = "\n".join(texto_completo.splitlines()[:12])

    for trecho in (primeiras_linhas, texto_completo):
        match = REGEX_DESPACHO_TEXTO.search(trecho)
        if match:
            apenas_digitos = re.sub(r"\D", "", match.group(1))
            ano = match.group(2) if match.group(2) else ANO_PADRAO
            if apenas_digitos:
                return f"{int(apenas_digitos)}/{ano}"

    # Fallback: Nome do arquivo (Despacho_XXXX_...)
    stem = Path(nome_arquivo).stem
    match_arq = REGEX_DESPACHO_ARQUIVO.search(stem) or REGEX_NUMERO_ENTRE_UNDERSCORES.search(stem)
    if match_arq:
        apenas_digitos = re.sub(r"\D", "", match_arq.group(1))
        if apenas_digitos:
            return f"{int(apenas_digitos)}/{ANO_PADRAO}"

    return None


def isolar_bloco_decide(texto_completo: str) -> str:
    """
    Retorna o texto que vem logo após 'DECIDE:' (onde começa o 2º parágrafo).
    Se não encontrar a palavra 'DECIDE:', retorna o texto completo.
    """
    partes = re.split(r"DECIDE\s*:", texto_completo, maxsplit=1, flags=re.IGNORECASE)
    if len(partes) > 1:
        return partes[1]
    return texto_completo


def extrair_codigo_pe_despacho(bloco_decide: str, texto_completo: str) -> str | None:
    """
    Extrai o código PE-XXXXX-XXXX/XXXX priorizando o 2º parágrafo (após 'DECIDE:').
    """
    for fonte in (bloco_decide, colar_quebras_de_codigo(bloco_decide), texto_completo, colar_quebras_de_codigo(texto_completo)):
        achados = REGEX_PROJETO_PE.findall(fonte)
        if achados:
            cod_padronizado = re.sub(r"\s*[/_–—-]\s*(\d{4})$", r"/\1", achados[0])
            return chave_projeto(cod_padronizado)
    return None


def extrair_valor_despacho(bloco_decide: str, texto_completo: str) -> float | None:
    """
    Extrai o valor monetário acompanhado de parênteses 'R$ XXXXX (...)' no 2º parágrafo (após 'DECIDE:').
    """
    # 1ª tentativa: padrão exato 'R$ XXXXX (...)' no parágrafo após DECIDE:
    match_parenteses = REGEX_VALOR_COM_PARENTESES.search(bloco_decide)
    if match_parenteses:
        return converter_para_float_br(match_parenteses.group(1))

    # 2ª tentativa: padrão 'R$ XXXXX (...)' no texto completo (caso quebras de página afetem o split)
    match_parenteses_geral = REGEX_VALOR_COM_PARENTESES.search(texto_completo)
    if match_parenteses_geral:
        return converter_para_float_br(match_parenteses_geral.group(1))

    # 3ª tentativa (Fallback): primeiro valor 'R$ XXXXX' dentro do parágrafo após DECIDE:
    match_simples = REGEX_VALOR_BR.search(bloco_decide)
    if match_simples:
        return converter_para_float_br(match_simples.group(1))

    return None


def analisar_pdf_despacho(caminho_pdf: Path) -> dict:
    """Abre um PDF de Despacho e extrai Número do Despacho, Código PE e Valor em R$."""
    textos_paginas = []
    with pdfplumber.open(str(caminho_pdf)) as pdf:
        for pagina in pdf.pages:
            txt = normalizar_texto_pdf(pagina.extract_text() or "")
            if txt.strip():
                textos_paginas.append(txt)

    texto_completo = "\n".join(textos_paginas)
    bloco_decide = isolar_bloco_decide(texto_completo)

    numero_despacho = extrair_numero_despacho(texto_completo, caminho_pdf.name)
    codigo_pe = extrair_codigo_pe_despacho(bloco_decide, texto_completo)
    valor_float = extrair_valor_despacho(bloco_decide, texto_completo)

    return {
        "arquivo": caminho_pdf.name,
        "despacho": numero_despacho,
        "codigo_projeto": codigo_pe,
        "valor_float": valor_float,
    }


def executar_atualizacao_despachos(
    pasta_despachos: Path,
    caminho_planilha: Path,
    caminho_warnings: Path,
    nome_aba: str | int = 0,
    recursivo: bool = False,
):
    inicializar_pastas()

    if not pasta_despachos.exists():
        raise FileNotFoundError(f"Pasta de Despachos não encontrada: {pasta_despachos}")
    if not caminho_planilha.exists():
        raise FileNotFoundError(f"Planilha Excel não encontrada: {caminho_planilha}")

    print("=" * 80)
    print(" LENDO PDFS DE DESPACHO E ATUALIZANDO PLANILHA EXCEL")
    print("=" * 80)

    iterador = pasta_despachos.rglob("*") if recursivo else pasta_despachos.iterdir()
    arquivos_pdf = sorted(
        arq for arq in iterador
        if arq.is_file() and arq.suffix.lower() == ".pdf"
    )

    print(f"[1/4] Encontrados {len(arquivos_pdf)} arquivo(s) .pdf na pasta de Despachos.")

    despachos_validos: list[dict] = []
    warnings: list[str] = []

    for idx, arq in enumerate(arquivos_pdf, start=1):
        try:
            info = analisar_pdf_despacho(arq)
        except Exception as erro:
            warnings.append(f"[ERRO AO LER PDF] Arquivo: {arq.name} | Erro: {erro}")
            continue

        cod = info["codigo_projeto"]
        desp = info["despacho"]
        val = info["valor_float"]

        if not cod:
            warnings.append(
                f"[CÓDIGO PE NÃO IDENTIFICADO] Despacho: {desp or 'N/A'} | Arquivo: {arq.name}"
            )
            print(f"   [{idx}/{len(arquivos_pdf)}] [AVISO - SEM PE] Despacho: {desp} ({arq.name})")
            continue

        if not desp:
            warnings.append(
                f"[NÚMERO DE DESPACHO NÃO IDENTIFICADO] Projeto: {cod} | Arquivo: {arq.name}"
            )
            print(f"   [{idx}/{len(arquivos_pdf)}] [AVISO - SEM Nº DESPACHO] {cod} ({arq.name})")
            continue

        if val is None:
            warnings.append(
                f"[VALOR R$ NÃO ENCONTRADO NO PDF] Projeto: {cod} | Despacho: {desp} | Arquivo: {arq.name}"
            )
            print(f"   [{idx}/{len(arquivos_pdf)}] [AVISO - SEM VALOR R$] {cod} -> Despacho: {desp}")
            continue

        despachos_validos.append(info)
        print(f"   [{idx}/{len(arquivos_pdf)}] [LIDO] {cod} -> Despacho: {desp} | Valor: {formatar_moeda_br(val)}")

    # 2. Carregar a planilha Excel e mapear os projetos existentes
    print(f"\n[2/4] Cruzando {len(despachos_validos)} Despacho(s) com a planilha '{caminho_planilha.name}'...")
    df = pd.read_excel(caminho_planilha, sheet_name=nome_aba)
    colunas_originais = [str(c).strip() for c in df.columns]
    df.columns = colunas_originais

    col_cod = encontrar_coluna(colunas_originais, ["Código do Projeto", "Codigo do Projeto"])
    col_despacho = encontrar_coluna(colunas_originais, ["Despacho"])
    col_valor = encontrar_coluna(colunas_originais, ["Valor Aprovado"])

    idx_col_despacho = colunas_originais.index(col_despacho) + 1
    idx_col_valor = colunas_originais.index(col_valor) + 1

    # Cria um índice {chave_pe: [(linha_excel, valor_atual_na_planilha), ...]}
    mapa_planilha: dict[str, list[tuple[int, object]]] = {}
    for idx_df, row in df.iterrows():
        cod_bruto = str(row[col_cod]).strip() if pd.notna(row[col_cod]) else ""
        if not cod_bruto:
            continue
        cod_norm = re.sub(r"\s*[/_–—-]\s*(\d{4})$", r"/\1", cod_bruto)
        try:
            chave = chave_projeto(cod_norm)
        except ValueError:
            continue
        mapa_planilha.setdefault(chave, []).append((idx_df + 2, row[col_valor]))

    # Verifica se a coluna Valor Aprovado na planilha guarda textos ("R$ XX,XX") ou números float
    amostra_valores = [v for v in df[col_valor].dropna().tolist() if str(v).strip() != ""]
    usa_string_rs = any(isinstance(v, str) and "R$" in v.upper() for v in amostra_valores)

    # 3. Comparar valores e definir atualizações
    atualizacoes = []  # Lista de tuplas: (linha_excel, codigo_pe, novo_despacho, novo_valor_ou_None)

    for item in despachos_validos:
        cod = item["codigo_projeto"]
        desp = item["despacho"]
        val_pdf = item["valor_float"]
        arq_nome = item["arquivo"]

        # Caso 1: Código PE não encontrado na planilha Excel
        if cod not in mapa_planilha:
            warnings.append(
                f"[PROJETO NÃO ENCONTRADO NA PLANILHA] Projeto: {cod} | Despacho: {desp} | "
                f"Valor no Despacho: {formatar_moeda_br(val_pdf)} | Arquivo: {arq_nome}"
            )
            continue

        for linha_excel, val_planilha_bruto in mapa_planilha[cod]:
            val_planilha_float = converter_para_float_br(val_planilha_bruto)
            planilha_zerada = eh_zero_reais(val_planilha_bruto)

            # Caso 2: Valor na planilha é 0,00 (ou vazio) -> Reescreve o Valor Aprovado e preenche o Despacho
            if planilha_zerada:
                novo_valor_gravar = formatar_moeda_br(val_pdf) if usa_string_rs else val_pdf
                atualizacoes.append((linha_excel, cod, desp, novo_valor_gravar))

            # Caso 3: Valor na planilha é IGUAL ao do Despacho (tolerância de 1 centavo para arredondamento)
            elif val_planilha_float is not None and abs(val_planilha_float - val_pdf) < 0.01:
                atualizacoes.append((linha_excel, cod, desp, None))

            # Caso 4: Valor na planilha é DIFERENTE do Despacho -> Registra no warning
            else:
                warnings.append(
                    f"[VALOR DIVERGENTE] Projeto: {cod} | Despacho: {desp} | "
                    f"Valor na Planilha: {formatar_moeda_br(val_planilha_float)} ({val_planilha_bruto}) | "
                    f"Valor no Despacho: {formatar_moeda_br(val_pdf)} | Arquivo: {arq_nome}"
                )

    # 4. Salvar arquivo de warnings de auditoria
    caminho_warnings.parent.mkdir(parents=True, exist_ok=True)
    with open(caminho_warnings, "w", encoding="utf-8") as f:
        if warnings:
            f.write("\n".join(warnings) + "\n")
    print(f"[3/4] Arquivo de auditoria salvo em '{caminho_warnings.name}' ({len(warnings)} ocorrência(s)).")

    # 5. Gravar alterações na planilha Excel
    print(f"[4/4] Aplicando atualizações em {len(atualizacoes)} linha(s) da planilha...")
    if atualizacoes:
        while True:
            try:
                wb = openpyxl.load_workbook(caminho_planilha)
                ws = wb[nome_aba] if isinstance(nome_aba, str) else wb.worksheets[int(nome_aba)]

                for linha_excel, cod, novo_desp, novo_val in atualizacoes:
                    ws.cell(row=linha_excel, column=idx_col_despacho, value=novo_desp)
                    if novo_val is not None:
                        cel_val = ws.cell(row=linha_excel, column=idx_col_valor, value=novo_val)
                        if isinstance(novo_val, float):
                            cel_val.number_format = "R$ #,##0.00"

                wb.save(caminho_planilha)
                wb.close()
                break
            except PermissionError:
                print("\n [!] ATENÇÃO: A planilha está aberta no Excel!")
                input(" >>> Feche a planilha no Excel e pressione ENTER para salvar... ")

    print("\n" + "=" * 80)
    print(" RESUMO FINAL (PROCESSAMENTO DE DESPACHOS)")
    print("=" * 80)
    print(f" - PDFs de Despacho lidos com sucesso             : {len(despachos_validos)}/{len(arquivos_pdf)}")
    print(f" - Linhas atualizadas na planilha Excel           : {len(atualizacoes)}")
    print(f" - Ocorrências salvas em warning_despachos_pdf.txt: {len(warnings)}")
    print(f" - Caminho do arquivo de warnings                 : {caminho_warnings}")
    print("=" * 80)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Lê PDFs de Despacho, valida o Valor Aprovado e preenche a coluna Despacho na planilha Excel."
    )
    parser.add_argument("--pasta", "-p", type=Path, default=PASTA_DESPACHOS_PADRAO, help="Pasta com os PDFs de Despacho.")
    parser.add_argument("--planilha", "-x", type=Path, default=PLANILHA_EXCEL_PADRAO, help="Planilha Excel (.xlsx).")
    parser.add_argument("--warnings", "-w", type=Path, default=ARQUIVO_WARNINGS_PADRAO, help="Arquivo TXT de warnings.")
    parser.add_argument("--aba", "-a", default=0, help="Nome ou número da aba da planilha (padrão: 0).")
    parser.add_argument("--recursivo", "-r", action="store_true", help="Busca PDFs também dentro de subpastas.")

    args = parser.parse_args()
    executar_atualizacao_despachos(
        pasta_despachos=args.pasta,
        caminho_planilha=args.planilha,
        caminho_warnings=args.warnings,
        nome_aba=args.aba,
        recursivo=args.recursivo,
    )
