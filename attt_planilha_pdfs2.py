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

# Regex multi-linha para PE: suporta quebra de linha e espaços grandes no meio do código
# Ex: "PE-\n   00404-0072/2016", "PE-00404-\n   0072-2016", "PE-00404-00\n   72/2016"
REGEX_PROJETO_PE = re.compile(
    r"\bPE[\s\r\n]*[-–—_][\s\r\n]*"
    r"(\d(?:[\s\r\n]*\d){2,5})[\s\r\n]*[-–—_][\s\r\n]*"
    r"(\d(?:[\s\r\n]*\d){2,4})[\s\r\n]*[/_–—-][\s\r\n]*"
    r"((?:19|20)\d(?:[\s\r\n]*\d){2})\b",
    re.IGNORECASE,
)

# Identifica projetos não-PE (ex: PG, PD) mesmo com quebra de linha para ignorá-los
REGEX_PROJETO_OUTROS = re.compile(
    r"\b(?!PE[\s\r\n]*[-–—_])([A-Z]{2,3})[\s\r\n]*[-–—_][\s\r\n]*"
    r"(\d(?:[\s\r\n]*\d){2,5})[\s\r\n]*[-–—_][\s\r\n]*"
    r"(\d(?:[\s\r\n]*\d){2,4})[\s\r\n]*[/_–—-][\s\r\n]*"
    r"((?:19|20)\d(?:[\s\r\n]*\d){2})\b",
    re.IGNORECASE,
)

# Regex flexível para "NOTA TÉCNICA Nº XXX/2026":
# Ignora traço embaixo do Nº (_), espaços, quebras de linha ou símbolos entre NOTA TÉCNICA e o número
REGEX_NT_CABECALHO = re.compile(
    r"N\s*O\s*T\s*A\s+T\s*[EÉ]\s*C\s*N\s*I\s*C\s*A"
    r"[^0-9X]{0,30}?"
    r"(\d{1,5}|X{1,5})\s*[/_–—-]\s*(20\d{2})",
    re.IGNORECASE,
)

# Primeiro número entre "_" e "_" no nome do arquivo (_XXX_)
REGEX_NUMERO_ARQUIVO = re.compile(r"_\s*(\d+)\s*_")

# Valores monetários brasileiros (ex: R$ 1.234.567,89 ou 0,00)
REGEX_VALOR_BR = re.compile(r"(?:R\$\s*)?\d{1,3}(?:\.\d{3})*,\d{2}|\b\d+,\d{2}\b")


def normalizar_unicode_pdf(texto: str) -> str:
    """
    Converte caracteres decompostos do PDF (ex: 'E' + acento separado) para NFC ('É')
    e une números que o PDF tenha separado com espaço após ponto/vírgula.
    """
    if not texto:
        return ""
    txt = unicodedata.normalize("NFC", str(texto))
    # Une espaçamentos de kerning em números (ex: '150. 000, 00' -> '150.000,00')
    txt = re.sub(r"(\d)\s*([.,])\s*(\d)", r"\1\2\3", txt)
    return txt


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
    txt = normalizar_unicode_pdf(str(valor_celula))
    return re.sub(r"\s+", " ", txt).strip()


def converter_para_float_br(texto_valor) -> float | None:
    if texto_valor is None or pd.isna(texto_valor):
        return None
    if isinstance(texto_valor, (int, float)):
        return float(texto_valor)

    texto = normalizar_unicode_pdf(str(texto_valor)).strip()
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


def montar_codigo_pe_limpo(grupos_match: tuple[str, str, str]) -> str:
    """Reconstrói 'PE-XXXXX-XXXX/XXXX' removendo quebras de linha e espaços internos."""
    p1 = re.sub(r"\s+", "", grupos_match[0])
    p2 = re.sub(r"\s+", "", grupos_match[1])
    ano = re.sub(r"\s+", "", grupos_match[2])
    return chave_projeto(f"PE-{p1}-{p2}/{ano}")


def extrair_texto_bloco_assunto(primeira_pagina) -> str:
    """
    Isola as palavras da região do campo 'Assunto' (canto direito do documento)
    para evitar que quebras de linha com recuo grande misturem colunas.
    """
    try:
        palavras = primeira_pagina.extract_words(keep_blank_chars=False) or []
    except Exception:
        return ""

    w_assunto = None
    for w in palavras:
        if re.search(r"^Assunto:?$", normalizar_unicode_pdf(w["text"]), re.IGNORECASE):
            w_assunto = w
            break

    if not w_assunto:
        return ""

    # Pega todas as palavras que estão na coluna do Assunto (mesma margem X ou à direita)
    # e abaixo da linha do Assunto (até no máximo 250 pontos abaixo)
    x_min = max(0, w_assunto["x0"] - 25)
    y_min = w_assunto["top"] - 5
    y_max = w_assunto["top"] + 250

    palavras_assunto = [
        w for w in palavras
        if w["x0"] >= x_min and y_min <= w["top"] <= y_max
    ]
    # Ordena linha por linha (agrupando a cada 5pt verticais) e da esquerda para a direita
    palavras_assunto.sort(key=lambda w: (round(w["top"] / 5.0), w["x0"]))
    return normalizar_unicode_pdf(" ".join(w["text"] for w in palavras_assunto))


def extrair_codigo_projeto(primeira_pagina, texto_completo: str) -> tuple[str | None, str | None]:
    """
    Procura o código PE-XXXXX-XXXX/XXXX (mesmo quebrado em 2 linhas):
    1º No bloco 'Assunto' (canto superior do documento)
    2º No bloco 'I - DO OBJETIVO' (última ocorrência)
    3º No corpo completo do documento
    """
    # 1. Bloco Assunto (isolado geometricamente no canto da 1ª página)
    texto_assunto = extrair_texto_bloco_assunto(primeira_pagina)
    if texto_assunto:
        matches_pe_assunto = REGEX_PROJETO_PE.findall(texto_assunto)
        if matches_pe_assunto:
            return montar_codigo_pe_limpo(matches_pe_assunto[-1]), None
        matches_outros_assunto = REGEX_PROJETO_OUTROS.findall(texto_assunto)
        if matches_outros_assunto:
            g = matches_outros_assunto[-1]
            cod_outro = re.sub(r"\s+", "", f"{g[0]}-{g[1]}-{g[2]}/{g[3]}").upper()
            return None, cod_outro

    # 2. Bloco 'I - DO OBJETIVO'
    match_bloco = re.search(
        r"(?:I\s*[-–—.]?\s*DO\s+OBJETIVO)(.*?)(?:II\s*[-–—.]|\n\s*2\s*[-–—.]|$)",
        texto_completo,
        re.IGNORECASE | re.DOTALL,
    )
    if match_bloco:
        trecho_objetivo = match_bloco.group(1)
        matches_pe = REGEX_PROJETO_PE.findall(trecho_objetivo)
        if matches_pe:
            return montar_codigo_pe_limpo(matches_pe[-1]), None
        matches_outros = REGEX_PROJETO_OUTROS.findall(trecho_objetivo)
        if matches_outros:
            g = matches_outros[-1]
            cod_outro = re.sub(r"\s+", "", f"{g[0]}-{g[1]}-{g[2]}/{g[3]}").upper()
            return None, cod_outro

    # 3. Fallback no texto completo (inclui trecho após a palavra Assunto no texto corrido)
    matches_pe_geral = REGEX_PROJETO_PE.findall(texto_completo)
    if matches_pe_geral:
        return montar_codigo_pe_limpo(matches_pe_geral[0]), None

    matches_outros_geral = REGEX_PROJETO_OUTROS.findall(texto_completo)
    if matches_outros_geral:
        g = matches_outros_geral[0]
        cod_outro = re.sub(r"\s+", "", f"{g[0]}-{g[1]}-{g[2]}/{g[3]}").upper()
        return None, cod_outro

    return None, None


def extrair_nota_tecnica(primeira_pagina, texto_primeira_pagina: str, nome_arquivo: str) -> tuple[str | None, bool]:
    """
    Lê a primeira linha escrita no topo do PDF ('NOTA TÉCNICA Nº XXX/2026'),
    tolerando traço sublinhado embaixo do Nº, sobrescrito deslocado e acentos decompostos.
    Só usa o nome do arquivo (_XXX_) e marca warning se no documento estiver 'XX' ou sem número.
    """
    # Reconstrói as primeiras linhas do topo da página 1 agrupando palavras com tolerância vertical de 8pt
    # (isso impede que o 'º' com traço embaixo fique em linha separada de 'NOTA TÉCNICA N')
    texto_topo_ordenado = ""
    try:
        palavras_pag1 = primeira_pagina.extract_words(keep_blank_chars=False) or []
        if palavras_pag1:
            top_min = min(w["top"] for w in palavras_pag1)
            palavras_topo = [w for w in palavras_pag1 if w["top"] <= top_min + 180]
            palavras_topo.sort(key=lambda w: (round(w["top"] / 8.0), w["x0"]))
            texto_topo_ordenado = normalizar_unicode_pdf(" ".join(w["text"] for w in palavras_topo))
    except Exception:
        pass

     Primeiras_linhas_brutas = "\n".join(texto_primeira_pagina.splitlines()[:15])
    fontes_busca_topo = [texto_topo_ordenado, Primeiras_linhas_brutas, texto_primeira_pagina]

    encontrou_xx_literal = False

    for fonte in fontes_busca_topo:
        if not fonte:
            continue
        for match in REGEX_NT_CABECALHO.finditer(fonte):
            num_bruto = match.group(1).strip().upper()
            ano_pdf = match.group(2).strip()
            if "X" not in num_bruto and num_bruto.isdigit():
                return f"{int(num_bruto)}/{ano_pdf}", False
            else:
                encontrou_xx_literal = True

    # Fallback: se na 1ª linha do arquivo houver diretamente um padrão 'XXX/2026' junto à palavra NOTA ou TÉCNICA
    primeiras_3_linhas = "\n".join(texto_primeira_pagina.splitlines()[:4])
    if re.search(r"NOTA|T[EÉ]CNICA", primeiras_3_linhas, re.IGNORECASE):
        match_curto = re.search(r"\b(\d{1,5})\s*/\s*(20\d{2})\b", primeiras_3_linhas)
        if match_curto:
            return f"{int(match_curto.group(1))}/{match_curto.group(2)}", False

    # Se chegou aqui, ou estava com 'XX/2026' ou não tinha número no cabeçalho -> puxa do nome do arquivo (_XXX_)
    stem = Path(nome_arquivo).stem
    match_arq = REGEX_NUMERO_ARQUIVO.search(stem)
    if match_arq:
        num_arq = int(match_arq.group(1))
        # Dispara o warning '-Esqueceu de mudar na própria nota técnica-'
        return f"{num_arq}/{ANO_PADRAO}", True if encontrou_xx_literal or True else False

    return None, encontrou_xx_literal


def extrair_valor_de_matriz_tabela(tab: list[list[str | None]]) -> tuple[str | None, bool, str]:
    """
    Suporta tabelas de 2 linhas, 3 linhas ou mais.
    Varre de baixo para cima procurando a linha 'Total' e pega o valor na coluna 'Reconhecido (R$)'
    (imediatamente após 'Total').
    """
    if not tab:
        return None, False, ""

    idx_col_reconhecido = None
    for linha in tab[:3]:
        if not linha:
            continue
        for idx_c, cel in enumerate(linha):
            txt_c = limpar_celula_pdf(cel)
            if re.search(r"Reconhecid[oa]", txt_c, re.IGNORECASE):
                idx_col_reconhecido = idx_c
                break
        if idx_col_reconhecido is not None:
            break

    # Varre DE BAIXO PARA CIMA (funciona mesmo se a tabela tiver apenas 2 linhas: Cabeçalho + Total)
    for linha in reversed(tab):
        if not linha:
            continue

        texto_linha_inteira = " | ".join(limpar_celula_pdf(c) for c in linha if limpar_celula_pdf(c))
        if not re.search(r"\bTotal\b", texto_linha_inteira, re.IGNORECASE):
            continue

        # 1. Tenta extrair todos os valores monetários que aparecem na linha após a palavra 'Total'
        match_apos_total = re.search(r"\bTotal\b(.*)$", texto_linha_inteira, re.IGNORECASE)
        if match_apos_total:
            trecho_direito = match_apos_total.group(1)
            valores_linha = REGEX_VALOR_BR.findall(trecho_direito)
            if valores_linha:
                val_aprov = valores_linha[0]
                glosa = valores_linha[1] if len(valores_linha) > 1 else ""
                teve_glosa = bool(glosa) and not eh_zero_reais(glosa)
                return val_aprov, teve_glosa, glosa

        # 2. Fallback pelo índice da coluna 'Reconhecido (R$)'
        if idx_col_reconhecido is not None and idx_col_reconhecido < len(linha):
            val_col = limpar_celula_pdf(linha[idx_col_reconhecido])
            if REGEX_VALOR_BR.search(val_col):
                glosa = ""
                if idx_col_reconhecido + 1 < len(linha):
                    glosa = limpar_celula_pdf(linha[idx_col_reconhecido + 1])
                teve_glosa = bool(glosa) and not eh_zero_reais(glosa)
                return val_col, teve_glosa, glosa

    return None, False, ""


def extrair_valor_pdf_completo(pdf) -> tuple[str | None, bool, str]:
    """
    Busca o valor na linha 'Total' / coluna 'Reconhecido (R$)' de baixo para cima usando 3 camadas:
    Camada 1: Tabelas com bordas e tabelas de apenas 2 linhas (min_words_vertical=1).
    Camada 2: Cruzamento geométrico (X, Y) da linha 'Total' com a coluna 'Reconhecido (R$)'.
    Camada 3: Varredura de linhas de texto de baixo para cima (incluindo linha vizinha).
    """
    configuracoes_tabela = [
        {},  # Padrão (linhas fechadas)
        {
            "vertical_strategy": "lines",
            "horizontal_strategy": "text",
            "min_words_vertical": 1,
            "min_words_horizontal": 1,
            "snap_tolerance": 5,
        },
        {
            "vertical_strategy": "text",
            "horizontal_strategy": "lines",
            "min_words_vertical": 1,  # Essencial para tabelas de apenas 2 linhas!
            "min_words_horizontal": 1,
            "snap_tolerance": 5,
        },
        {
            "vertical_strategy": "text",
            "horizontal_strategy": "text",
            "min_words_vertical": 2,
            "min_words_horizontal": 1,
            "snap_tolerance": 6,
            "join_tolerance": 6,
        },
    ]

    # Percorre as páginas DE BAIXO PARA CIMA
    for idx_pag in range(len(pdf.pages) - 1, -1, -1):
        pagina = pdf.pages[idx_pag]

        # ----------------------------------------------------------------------
        # CAMADA 1: Matrizes de Tabela (incluindo tabelas de 2 linhas)
        # ----------------------------------------------------------------------
        for cfg in configuracoes_tabela:
            try:
                tabelas = pagina.extract_tables(table_settings=cfg) if cfg else pagina.extract_tables()
            except Exception:
                tabelas = []

            for tab in reversed(tabelas or []):
                val, glosa_flag, glosa_txt = extrair_valor_de_matriz_tabela(tab)
                if val and not eh_zero_reais(val):
                    return val, glosa_flag, glosa_txt

        # ----------------------------------------------------------------------
        # CAMADA 2: Alinhamento Geométrico (Linha 'Total' x Coluna 'Reconhecido')
        # ----------------------------------------------------------------------
        try:
            palavras = pagina.extract_words(keep_blank_chars=False) or []
        except Exception:
            palavras = []

        palavras_baixo_cima = sorted(palavras, key=lambda w: w["top"], reverse=True)
        for w_total in palavras_baixo_cima:
            txt_w = normalizar_unicode_pdf(w_total["text"]).strip()
            if re.match(r"^Total:?$", txt_w, re.IGNORECASE):
                y_centro = (w_total["top"] + w_total["bottom"]) / 2.0
                x_min = w_total["x1"] - 3

                # Captura todas as palavras na mesma faixa horizontal (+/- 14pt) à direita de 'Total'
                faixa_total = [
                    w for w in palavras
                    if abs(((w["top"] + w["bottom"]) / 2.0) - y_centro) <= 14 and w["x0"] >= x_min
                ]
                faixa_total.sort(key=lambda w: w["x0"])
                texto_faixa = normalizar_unicode_pdf(" ".join(w["text"] for w in faixa_total))
                valores_achados = REGEX_VALOR_BR.findall(texto_faixa)

                if valores_achados:
                    val_aprov = valores_achados[0]
                    glosa_txt = valores_achados[1] if len(valores_achados) > 1 else ""
                    teve_glosa = bool(glosa_txt) and not eh_zero_reais(glosa_txt)
                    if not eh_zero_reais(val_aprov):
                        return val_aprov, teve_glosa, glosa_txt

        # ----------------------------------------------------------------------
        # CAMADA 3: Varredura de Linhas de Texto (de baixo para cima + janela vizinha)
        # ----------------------------------------------------------------------
        for usar_layout in (True, False):
            try:
                txt_pag = normalizar_unicode_pdf(pagina.extract_text(layout=usar_layout) or "")
            except Exception:
                txt_pag = ""

            linhas = [l.strip() for l in txt_pag.splitlines() if l.strip()]
            for idx_l in range(len(linhas) - 1, -1, -1):
                if re.search(r"\bTotal\b", linhas[idx_l], re.IGNORECASE):
                    # Olha a própria linha do Total e, se o PDF tiver desalinhado, olha a linha vizinha
                    candidatas = [linhas[idx_l]]
                    if idx_l + 1 < len(linhas):
                        candidatas.append(linhas[idx_l + 1])
                    if idx_l - 1 >= 0:
                        candidatas.append(linhas[idx_l - 1])

                    for linha_cand in candidatas:
                        # Se a linha contém 'Total', pega apenas o que vem depois de 'Total'
                        m_tot = re.search(r"\bTotal\b(.*)$", linha_cand, re.IGNORECASE)
                        trecho_busca = m_tot.group(1) if m_tot else linha_cand
                        achados = REGEX_VALOR_BR.findall(trecho_busca)
                        if achados and not eh_zero_reais(achados[0]):
                            val_aprov = achados[0]
                            glosa_txt = achados[1] if len(achados) > 1 else ""
                            teve_glosa = bool(glosa_txt) and not eh_zero_reais(glosa_txt)
                            return val_aprov, teve_glosa, glosa_txt

    return None, False, ""


def analisar_pdf(caminho_pdf: Path) -> dict:
    textos_paginas = []

    with pdfplumber.open(str(caminho_pdf)) as pdf:
        for pagina in pdf.pages:
            txt = normalizar_unicode_pdf(pagina.extract_text() or "")
            if txt.strip():
                textos_paginas.append(txt)

        primeira_pagina = pdf.pages[0] if pdf.pages else None
        texto_primeira_pagina = textos_paginas[0] if textos_paginas else ""
        texto_completo = "\n".join(textos_paginas)

        codigo_pe, codigo_ignorado = (
            extrair_codigo_projeto(primeira_pagina, texto_completo)
            if primeira_pagina
            else (None, None)
        )
        if codigo_ignorado and not codigo_pe:
            return {
                "arquivo": caminho_pdf.name,
                "ignorado_nao_pe": True,
                "codigo_ignorado": codigo_ignorado,
            }

        nota_tecnica, esqueceu_xx = (
            extrair_nota_tecnica(primeira_pagina, texto_primeira_pagina, caminho_pdf.name)
            if primeira_pagina
            else (None, False)
        )

        valor_aprovado, teve_glosa, texto_glosa = extrair_valor_pdf_completo(pdf)

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
                f"Valor glosado: {info['texto_glosa']} | Valor aprovado (Reconhecido R$): {val} | Arquivo: {arq.name}"
            )

        if not cod:
            warnings.append(f"[PROJETO PE NÃO IDENTIFICADO] Regex PE-XXXXX-XXXX/XXXX não encontrada em: {arq.name}")
            print(f"   [{idx}/{len(arquivos_pdf)}] [AVISO - SEM CÓDIGO PE] ({arq.name})")
            continue

        if not val or eh_zero_reais(val):
            warnings.append(
                f"Sem valor encontrado | Projeto: {cod} | NT: {nt or 'N/A'} | Arquivo: {arq.name}"
            )
            print(f"   [{idx}/{len(arquivos_pdf)}] [AVISO - SEM VALOR] {cod} -> NT: {nt}")
        else:
            print(f"   [{idx}/{len(arquivos_pdf)}] [OK] {cod} -> NT: {nt} | Valor: {val}")

        dados_extraidos[cod] = info

    # 2. Salvar arquivo de warnings de auditoria
    caminho_warnings.parent.mkdir(parents=True, exist_ok=True)
    with open(caminho_warnings, "w", encoding="utf-8") as f:
        if warnings:
            f.write("\n".join(warnings) + "\n")
    print(f"\n[2/4] Arquivo de auditoria salvo em '{caminho_warnings.name}' ({len(warnings)} ocorrência(s)).")

    # 3. Ler planilha e preparar atualizações (sobrescrevendo R$ 0,00 ou vazios)
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
