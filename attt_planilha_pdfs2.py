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

# REGEX PRINCIPAIS (Exatamente as que funcionaram no seu código base)
REGEX_PROJETO_PE = re.compile(r"\bPE-\d+-\d+\s*[/_–—-]\s*\d{4}\b", re.IGNORECASE)
REGEX_PROJETO_OUTROS = re.compile(r"\b(?!PE-)[A-Z]{2,3}-\d+-\d+\s*[/_–—-]\s*\d{4}\b", re.IGNORECASE)

# Regex principal para "NOTA TÉCNICA Nº XX/2026"
REGEX_NT_CABECALHO = re.compile(
    r"NOTA\s+T[ÉE]CNICA\s*(?:N[º°o.]*)?\s*([0-9X]+)\s*/\s*(\d{4})",
    re.IGNORECASE,
)

# Regex FALLBACK para a 1ª linha da Nota Técnica (aceita traço '_' embaixo do Nº, símbolos ou quebra de linha)
REGEX_NT_TOPO_FALLBACK = re.compile(
    r"NOTA\s+T[EÉ]CNICA[\s_º°oªN.:º\-–—]*?([0-9X]{1,5})\s*/\s*(20\d{2})",
    re.IGNORECASE,
)

# Regex para o primeiro número entre "_" e "_" no nome do arquivo (_XXX_)
REGEX_NUMERO_ARQUIVO = re.compile(r"_\s*(\d+)\s*_")

# Regex para valores monetários no padrão brasileiro (ex: R$ 1.234.567,89 ou 0,00)
REGEX_VALOR_BR = re.compile(r"(?:R\$\s*)?\d{1,3}(?:\.\d{3})*,\d{2}|\b\d+,\d{2}\b")


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


def limpar_celula_pdf(valor_celula: str | None) -> str:
    if not valor_celula:
        return ""
    return re.sub(r"\s+", " ", str(valor_celula)).strip()


def converter_para_float_br(texto_valor) -> float | None:
    """
    Converte 'R$ 1.250.430,50' ou '1250430,50' para float (1250430.50)
    para que o Excel substitua células com 'R$ 0,00' sem erro de formatação.
    """
    if texto_valor is None or pd.isna(texto_valor):
        return None
    if isinstance(texto_valor, (int, float)):
        return float(texto_valor)

    texto = str(texto_valor).strip()
    match = REGEX_VALOR_BR.search(texto)
    alvo = match.group(0) if match else texto
    limpo = re.sub(r"[R$\s]", "", alvo)

    if "," in limpo:
        limpo = limpo.replace(".", "").replace(",", ".")
    try:
        return float(limpo)
    except ValueError:
        return None


def eh_zero_reais(texto_celula) -> bool:
    """Retorna True se o valor for R$ 0,00, 0 ou vazio."""
    if texto_celula is None or pd.isna(texto_celula):
        return True
    val_float = converter_para_float_br(texto_celula)
    if val_float is not None:
        return abs(val_float) < 0.005
    limpo = re.sub(r"[R$\s]", "", str(texto_celula)).upper()
    return limpo in ("", "0,00", "0.00", "0", "-", "NAN", "NONE")


def campo_nt_esta_vazio(valor) -> bool:
    if pd.isna(valor) or valor is None:
        return True
    texto = str(valor).strip()
    return not bool(texto) or texto.lower() in ("nan", "none", "-", "n/a", "não", "nao", "0")


def campo_valor_precisa_atualizar(valor_planilha) -> bool:
    """Retorna True se o valor na planilha estiver em branco OU for R$ 0,00 / 0."""
    return eh_zero_reais(valor_planilha)


def colar_quebras_de_codigo(texto: str) -> str:
    """
    Usado APENAS como fallback: remove quebras de linha e espaços largos
    ao redor de traços e barras (ex: 'PE-\\n   00404-0072/2016' -> 'PE-00404-0072/2016').
    """
    sem_quebra_traco = re.sub(r"([A-Z0-9])\s*[-–—_]\s*(\d)", r"\1-\2", texto, flags=re.IGNORECASE)
    sem_quebra_barra = re.sub(r"(\d)\s*[/_–—-]\s*(\d{4})\b", r"\1/\2", sem_quebra_traco)
    return sem_quebra_barra


def extrair_codigo_projeto(texto_completo: str, nome_arquivo: str = "") -> tuple[str | None, str | None]:
    """
    1º Executa exatamente a busca original que já funcionava.
    2º (Fallback) Se não achar por causa de quebra de linha no 'Assunto' ou 'Objetivo',
       cola as quebras de linha ao redor dos traços e tenta novamente.
    3º (Fallback) Verifica também o nome do arquivo.
    """
    # --------------------------------------------------------------------------
    # 1. MÉTODO PRINCIPAL (Intocado)
    # --------------------------------------------------------------------------
    match_bloco = re.search(
        r"(?:I\s*[-–—.]?\s*DO\s+OBJETIVO)(.*?)(?:II\s*[-–—.]|\n\s*2\s*[-–—.]|$)",
        texto_completo,
        re.IGNORECASE | re.DOTALL,
    )
    if match_bloco:
        trecho_objetivo = match_bloco.group(1)
        codigos_pe = REGEX_PROJETO_PE.findall(trecho_objetivo)
        if codigos_pe:
            return chave_projeto(re.sub(r"\s*[/_–—-]\s*(\d{4})$", r"/\1", codigos_pe[-1])), None
        codigos_outros = REGEX_PROJETO_OUTROS.findall(trecho_objetivo)
        if codigos_outros:
            return None, codigos_outros[-1].upper()

    codigos_pe_geral = REGEX_PROJETO_PE.findall(texto_completo)
    if codigos_pe_geral:
        return chave_projeto(re.sub(r"\s*[/_–—-]\s*(\d{4})$", r"/\1", codigos_pe_geral[0])), None

    codigos_outros_geral = REGEX_PROJETO_OUTROS.findall(texto_completo)
    if codigos_outros_geral:
        return None, codigos_outros_geral[0].upper()

    # --------------------------------------------------------------------------
    # 2. FALLBACK 1: Texto com quebra de linha no meio do código (ex: após Assunto)
    # --------------------------------------------------------------------------
    texto_colado = colar_quebras_de_codigo(texto_completo)
    codigos_pe_colado = REGEX_PROJETO_PE.findall(texto_colado)
    if codigos_pe_colado:
        return chave_projeto(re.sub(r"\s*[/_–—-]\s*(\d{4})$", r"/\1", codigos_pe_colado[0])), None

    codigos_outros_colado = REGEX_PROJETO_OUTROS.findall(texto_colado)
    if codigos_outros_colado:
        return None, codigos_outros_colado[0].upper()

    # --------------------------------------------------------------------------
    # 3. FALLBACK 2: Nome do próprio arquivo
    # --------------------------------------------------------------------------
    if nome_arquivo:
        nome_colado = colar_quebras_de_codigo(nome_arquivo)
        codigos_pe_arq = REGEX_PROJETO_PE.findall(nome_colado)
        if codigos_pe_arq:
            return chave_projeto(re.sub(r"\s*[/_–—-]\s*(\d{4})$", r"/\1", codigos_pe_arq[0])), None

    return None, None


def extrair_nota_tecnica(texto_primeira_pagina: str, texto_completo: str, nome_arquivo: str) -> tuple[str | None, bool]:
    """
    1º Tenta o método original.
    2º (Fallback) Se não achou número válido (por causa do traço '_' embaixo do Nº na 1ª linha
       ou acento decomposto no PDF), inspeciona o topo da 1ª página normalizado antes de cair para o nome do arquivo.
    """
    # Normaliza acentos decompostos do PDF (ex: 'E' + '´' -> 'É') sem alterar o texto
    pag1_norm = unicodedata.normalize("NFC", texto_primeira_pagina or "")
    completo_norm = unicodedata.normalize("NFC", texto_completo or "")

    # 1. MÉTODO PRINCIPAL
    match_pdf = REGEX_NT_CABECALHO.search(pag1_norm) or REGEX_NT_CABECALHO.search(completo_norm)

    esqueceu_xx = False
    if match_pdf:
        num_bruto = match_pdf.group(1).strip().upper()
        ano_pdf = match_pdf.group(2).strip()
        if "X" not in num_bruto and num_bruto.isdigit():
            return f"{int(num_bruto)}/{ano_pdf}", False
        else:
            esqueceu_xx = True

    # 2. FALLBACK NO TOPO DA 1ª PÁGINA (Trata o traço '_' embaixo do Nº na 1ª linha)
    # Se a regex principal falhou na 1ª linha por causa do traço embaixo do Nº,
    # testamos as 10 primeiras linhas com a regex que aceita '_' e símbolos sob o Nº:
    topo_10_linhas = "\n".join(pag1_norm.splitlines()[:10])
    match_topo_fallback = REGEX_NT_TOPO_FALLBACK.search(topo_10_linhas)
    if match_topo_fallback:
        num_fb = match_topo_fallback.group(1).strip().upper()
        ano_fb = match_topo_fallback.group(2).strip()
        if "X" not in num_fb and num_fb.isdigit():
            # Encontrou o número real na 1ª linha! Não precisa usar o nome do arquivo nem dar warning.
            return f"{int(num_fb)}/{ano_fb}", False
        else:
            esqueceu_xx = True

    # 3. FALLBACK FINAL: Nome do arquivo (_XXX_)
    stem = Path(nome_arquivo).stem
    match_arq = REGEX_NUMERO_ARQUIVO.search(stem)
    if match_arq:
        num_arq = int(match_arq.group(1))
        return f"{num_arq}/{ANO_PADRAO}", esqueceu_xx

    return None, esqueceu_xx


def extrair_valor_de_matriz_tabela(tab: list[list[str | None]]) -> tuple[str | None, bool, str]:
    """
    Varre uma tabela de baixo para cima procurando a linha 'Total'
    e extrai o valor da coluna 'Reconhecido (R$)' (imediatamente após 'Total').
    """
    if not tab:
        return None, False, ""

    idx_col_reconhecido = None
    for linha in tab[:3]:
        for idx_c, cel in enumerate(linha):
            txt_c = limpar_celula_pdf(cel)
            if re.search(r"Reconhecid[oa]", txt_c, re.IGNORECASE):
                idx_col_reconhecido = idx_c
                break
        if idx_col_reconhecido is not None:
            break

    for linha in reversed(tab):
        if not linha:
            continue

        texto_linha_inteira = " | ".join(limpar_celula_pdf(c) for c in linha if limpar_celula_pdf(c))
        if re.search(r"\bTotal\b", texto_linha_inteira, re.IGNORECASE):
            celulas_nao_vazias = [limpar_celula_pdf(c) for c in linha if limpar_celula_pdf(c)]

            for idx, texto_cel in enumerate(celulas_nao_vazias):
                if re.search(r"\bTotal\b", texto_cel, re.IGNORECASE):
                    # Caso 'Total' e o valor tenham ficado juntos na mesma célula
                    valores_mesma_celula = REGEX_VALOR_BR.findall(texto_cel)
                    if valores_mesma_celula:
                        val_aprov = valores_mesma_celula[0]
                        glosa = valores_mesma_celula[1] if len(valores_mesma_celula) > 1 else ""
                        if idx + 1 < len(celulas_nao_vazias) and not glosa:
                            glosa = celulas_nao_vazias[idx + 1]
                        teve_glosa = bool(glosa) and not eh_zero_reais(glosa)
                        return val_aprov, teve_glosa, glosa

                    # Caso padrão: coluna imediatamente após 'Total'
                    if idx + 1 < len(celulas_nao_vazias):
                        val_aprov = celulas_nao_vazias[idx + 1]
                        if REGEX_VALOR_BR.search(val_aprov):
                            glosa = celulas_nao_vazias[idx + 2] if idx + 2 < len(celulas_nao_vazias) else ""
                            teve_glosa = bool(glosa) and not eh_zero_reais(glosa)
                            return val_aprov, teve_glosa, glosa

            if idx_col_reconhecido is not None and idx_col_reconhecido < len(linha):
                val_col = limpar_celula_pdf(linha[idx_col_reconhecido])
                if REGEX_VALOR_BR.search(val_col):
                    glosa = ""
                    if idx_col_reconhecido + 1 < len(linha):
                        glosa = limpar_celula_pdf(linha[idx_col_reconhecido + 1])
                    teve_glosa = bool(glosa) and not eh_zero_reais(glosa)
                    return val_col, teve_glosa, glosa

            # Fallback se a linha tem 'Total' mas as células vieram divididas de forma atípica
            match_apos_total = re.search(r"\bTotal\b(.*)$", texto_linha_inteira, re.IGNORECASE)
            if match_apos_total:
                valores_restantes = REGEX_VALOR_BR.findall(match_apos_total.group(1))
                if valores_restantes:
                    val_aprov = valores_restantes[0]
                    glosa = valores_restantes[1] if len(valores_restantes) > 1 else ""
                    teve_glosa = bool(glosa) and not eh_zero_reais(glosa)
                    return val_aprov, teve_glosa, glosa

    return None, False, ""


def busca_relaxada_de_baixo_para_cima(pdf) -> tuple[str | None, bool, str]:
    """
    Acionada automaticamente apenas quando o extrator padrão de tabelas não encontra o valor.
    Mantém os 3 passos originais e inclui suporte a tabelas de apenas 2 linhas (min_words_vertical=1).
    """
    configuracoes_relaxadas = [
        # Configurações originais
        {"vertical_strategy": "text", "horizontal_strategy": "lines", "snap_tolerance": 5},
        {"vertical_strategy": "text", "horizontal_strategy": "text", "snap_tolerance": 6, "join_tolerance": 6},
        # Fallbacks específicos para tabelas curtas de apenas 2 linhas (Cabeçalho + Total)
        {"vertical_strategy": "text", "horizontal_strategy": "lines", "min_words_vertical": 1, "snap_tolerance": 5},
        {"vertical_strategy": "lines", "horizontal_strategy": "text", "min_words_vertical": 1, "snap_tolerance": 5},
    ]

    for idx_pag in range(len(pdf.pages) - 1, -1, -1):
        pagina = pdf.pages[idx_pag]

        # 1. Tenta reconstruir a tabela com estratégias flexíveis
        for cfg in configuracoes_relaxadas:
            try:
                tabelas = pagina.extract_tables(table_settings=cfg) or []
            except Exception:
                tabelas = []
            for tab in reversed(tabelas):
                val, glosa_flag, glosa_txt = extrair_valor_de_matriz_tabela(tab)
                if val and not eh_zero_reais(val):
                    return val, glosa_flag, glosa_txt

        # 2. Alinhamento espacial pelas coordenadas (X, Y) da palavra 'Total' de baixo para cima
        try:
            palavras = pagina.extract_words(keep_blank_chars=False) or []
        except Exception:
            palavras = []

        palavras_baixo_cima = sorted(palavras, key=lambda w: w["top"], reverse=True)
        for w_total in palavras_baixo_cima:
            if re.match(r"^Total:?$", w_total["text"].strip(), re.IGNORECASE):
                y_ref = w_total["top"]
                x_min = w_total["x1"] - 2
                mesma_faixa = [
                    w for w in palavras
                    if abs(w["top"] - y_ref) <= 14 and w["x0"] >= x_min
                ]
                mesma_faixa.sort(key=lambda w: w["x0"])
                texto_faixa = " ".join(w["text"] for w in mesma_faixa)
                valores_achados = REGEX_VALOR_BR.findall(texto_faixa)
                if valores_achados:
                    val_aprov = valores_achados[0]
                    glosa_txt = valores_achados[1] if len(valores_achados) > 1 else ""
                    teve_glosa = bool(glosa_txt) and not eh_zero_reais(glosa_txt)
                    return val_aprov, teve_glosa, glosa_txt

        # 3. Varredura das linhas de texto de baixo para cima (com fallback para linha vizinha)
        try:
            txt_pag = pagina.extract_text(layout=True) or pagina.extract_text() or ""
        except Exception:
            txt_pag = ""

        linhas = [l.strip() for l in txt_pag.splitlines() if l.strip()]
        for idx_l in range(len(linhas) - 1, -1, -1):
            if re.search(r"\bTotal\b", linhas[idx_l], re.IGNORECASE):
                achados = REGEX_VALOR_BR.findall(linhas[idx_l])
                if achados:
                    val_aprov = achados[0]
                    glosa_txt = achados[1] if len(achados) > 1 else ""
                    teve_glosa = bool(glosa_txt) and not eh_zero_reais(glosa_txt)
                    return val_aprov, teve_glosa, glosa_txt

                # Fallback: se na tabela de 2 linhas o número ficou na linha imediatamente abaixo/acima de 'Total'
                for vizinha in (idx_l + 1, idx_l - 1):
                    if 0 <= vizinha < len(linhas):
                        achados_viz = REGEX_VALOR_BR.findall(linhas[vizinha])
                        if achados_viz and not eh_zero_reais(achados_viz[0]):
                            val_aprov = achados_viz[0]
                            glosa_txt = achados_viz[1] if len(achados_viz) > 1 else ""
                            teve_glosa = bool(glosa_txt) and not eh_zero_reais(glosa_txt)
                            return val_aprov, teve_glosa, glosa_txt

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

        codigo_pe, codigo_ignorado = extrair_codigo_projeto(texto_completo, caminho_pdf.name)
        if codigo_ignorado and not codigo_pe:
            return {
                "arquivo": caminho_pdf.name,
                "ignorado_nao_pe": True,
                "codigo_ignorado": codigo_ignorado,
            }

        nota_tecnica, esqueceu_xx = extrair_nota_tecnica(texto_primeira_pagina, texto_completo, caminho_pdf.name)

        valor_aprovado = None
        teve_glosa = False
        texto_glosa = ""

        # 1ª tentativa: Tabelas padrão (de baixo para cima)
        for tab in reversed(todas_tabelas):
            val, glosa_flag, glosa_txt = extrair_valor_de_matriz_tabela(tab)
            if val and not eh_zero_reais(val):
                valor_aprovado = val
                teve_glosa = glosa_flag
                texto_glosa = glosa_txt
                break

        # 2ª tentativa automática (Fallback): Busca relaxada de baixo para cima se não achou na 1ª
        if not valor_aprovado or eh_zero_reais(valor_aprovado):
            valor_aprovado, teve_glosa, texto_glosa = busca_relaxada_de_baixo_para_cima(pdf)

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

    print("=" * 80)
    print(" LENDO ARQUIVOS .PDF (APENAS PE) E ATUALIZANDO PLANILHA AUTOMATICAMENTE")
    print("=" * 80)

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

        if not val or eh_zero_reais(val):
            warnings.append(
                f"Sem valor encontrado | Projeto: {cod} | NT: {nt or 'N/A'} | Arquivo: {arq.name}"
            )
            print(f"   [{idx}/{len(arquivos_pdf)}] [AVISO - SEM VALOR] {cod} -> NT: {nt}")
        else:
            print(f"   [{idx}/{len(arquivos_pdf)}] [OK] {cod} -> NT: {nt} | Valor: {val}")

        dados_extraidos[cod] = info

    # 2. Salvar arquivo de warnings (auditabilidade)
    caminho_warnings.parent.mkdir(parents=True, exist_ok=True)
    with open(caminho_warnings, "w", encoding="utf-8") as f:
        if warnings:
            f.write("\n".join(warnings) + "\n")
    print(f"\n[2/4] Arquivo de auditabilidade salvo em '{caminho_warnings.name}' ({len(warnings)} ocorrência(s)).")

    # 3. Ler planilha e preparar atualizações (reescrevendo sempre que Valor for vazio ou R$ 0,00)
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

        cod_norm = re.sub(r"\s*[/_–—-]\s*(\d{4})$", r"/\1", cod_bruto)
        try:
            chave = chave_projeto(cod_norm)
        except ValueError:
            continue

        if not chave.upper().startswith("PE-"):
            continue

        if chave not in dados_extraidos:
            continue

        info_pdf = dados_extraidos[chave]
        nt_pdf = info_pdf["nota_tecnica"]
        val_pdf_str = info_pdf["valor_aprovado"]

        val_pdf_num = converter_para_float_br(val_pdf_str)
        valor_para_gravar = val_pdf_num if val_pdf_num is not None else val_pdf_str

        aneel_atual = str(row[col_aneel]).strip().upper() if pd.notna(row[col_aneel]) else ""
        nt_vazia = campo_nt_esta_vazio(row[col_nt])
        valor_zerado_ou_vazio = campo_valor_precisa_atualizar(row[col_valor])

        era_nao = aneel_atual not in ("SIM", "S")
        precisa_atualizar = era_nao or nt_vazia or valor_zerado_ou_vazio

        if precisa_atualizar:
            linha_excel = idx_df + 2
            novo_aneel = texto_sim if era_nao else None
            nova_nt = nt_pdf if (era_nao or nt_vazia) and nt_pdf else None
            novo_val = (
                valor_para_gravar
                if (era_nao or valor_zerado_ou_vazio) and valor_para_gravar is not None and not eh_zero_reais(valor_para_gravar)
                else None
            )
            if novo_aneel is not None or nova_nt is not None or novo_val is not None:
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

    print("\n" + "=" * 80)
    print(" RESUMO FINAL (LEITURA DE PDFS)")
    print("=" * 80)
    print(f" - Projetos PE válidos extraídos                  : {len(dados_extraidos)}")
    print(f" - Projetos ignorados (PG, PD, etc.)              : {ignorados_nao_pe}")
    print(f" - Linhas efetivamente atualizadas no Excel       : {len(atualizacoes)}")
    print(f" - Ocorrências salvas no arquivo de warnings      : {len(warnings)}")
    print(f" - Caminho do arquivo de warnings                 : {caminho_warnings}")
    print("=" * 80)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Extrai Projeto (PE), Nota Técnica e Valor (Reconhecido R$ / Total) dos PDFs e atualiza a planilha."
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
