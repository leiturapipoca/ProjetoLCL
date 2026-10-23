import argparse
import re
import time
import unicodedata
from pathlib import Path
import pandas as pd
import openpyxl
from playwright.sync_api import sync_playwright, Page

from config import ONEDRIVE_BASE_DIR, inicializar_pastas
from controle_downloads import HistoricoDownloads, chave_projeto

# ==============================================================================
# 1. CONFIGURAÇÕES PADRÃO (Altere o caminho da planilha ou passe via --planilha)
# ==============================================================================
PLANILHA_ORIGINAL_PADRAO = Path(r"C:\automacao_opee\planilha_empresas.xlsx")

# Novo set persistente exclusivo para não repetir projetos já associados
ARQUIVO_SET_ASSOCIADOS = ONEDRIVE_BASE_DIR / "projetos_associados.txt"

# Nova tabela gerada apenas para visualização humana (salva em .xlsx e .txt)
ARQUIVO_NOVA_TABELA_XLSX = ONEDRIVE_BASE_DIR / "tabela_projetos_associados.xlsx"
ARQUIVO_NOVA_TABELA_TXT = ONEDRIVE_BASE_DIR / "tabela_projetos_associados.txt"

AVALIADOR_PADRAO = "GUILHERME SILVA CAVALCANTE"
STATUS_FILTRO = "Aguardando definição de avaliadores"
NOVO_STATUS = "Aguardando avaliação interna"

# As 18 colunas exatas que serão transferidas para a nova tabela
COLUNAS_NOVA_TABELA = [
    "Distribuidora",
    "Codigo do Projeto",
    "Ano do Projeto",
    "Título o Projeto",
    "Tipologia",
    "Invesstimento OPEE - Custos por Categoria Contábil",
    "Investimento Real",
    "RCB Global",
    "Energia Economizada MWh/Ano",
    "Demanda Retirada na Ponta KW",
    "Status",
    "Executado",
    "Valor Correto no OPEE",
    "Avaliado Pela ANEEL",
    "Nota Técnica",
    "Despacho",
    "Valor Aprovado",
    "Ano de Avaliação",
]


def normalizar_texto(texto: str) -> str:
    """Remove acentos, espaços extras e letras duplicadas para encontrar colunas mesmo com pequenas variações."""
    if not texto:
        return ""
    sem_acento = "".join(
        c for c in unicodedata.normalize("NFD", str(texto))
        if unicodedata.category(c) != "Mn"
    )
    limpo = re.sub(r"[^a-zA-Z0-9]+", " ", sem_acento).strip().lower()
    # Reduz letras duplicadas seguidas (ex: 'invesstimento' == 'investimento')
    limpo = re.sub(r"(.)\1+", r"\1", limpo)
    return limpo


def mapear_colunas_planilha(colunas_reais: list[str]) -> dict[str, str]:
    """
    Associa cada nome da lista COLUNAS_NOVA_TABELA ao nome real da coluna no Excel,
    tolerando variações como 'Título o Projeto' vs 'Título do Projeto' ou 'Invesstimento' vs 'Investimento'.
    """
    mapa_norm = {normalizar_texto(col): col for col in colunas_reais}
    resultado = {}

    apelidos = {
        "Título o Projeto": ["Titulo do Projeto", "Título do Projeto", "Titulo o Projeto"],
        "Invesstimento OPEE - Custos por Categoria Contábil": [
            "Investimento OPEE - Custos por Categoria Contábil",
            "Investimento OPEE - Custos por Categoria Contabil",
        ],
        "Codigo do Projeto": ["Código do Projeto", "Codigo do Projeto"],
    }

    for col_alvo in COLUNAS_NOVA_TABELA:
        if col_alvo in colunas_reais:
            resultado[col_alvo] = col_alvo
            continue

        norm_alvo = normalizar_texto(col_alvo)
        if norm_alvo in mapa_norm:
            resultado[col_alvo] = mapa_norm[norm_alvo]
            continue

        achou = False
        for alt in apelidos.get(col_alvo, []):
            if normalizar_texto(alt) in mapa_norm:
                resultado[col_alvo] = mapa_norm[normalizar_texto(alt)]
                achou = True
                break

        if not achou:
            resultado[col_alvo] = col_alvo  # Mantém o nome original caso não exista

    return resultado


def valor_preenchido(val) -> bool:
    """Verifica se o campo Nota Técnica está realmente preenchido."""
    if pd.isna(val) or val is None:
        return False
    texto = str(val).strip()
    return bool(texto) and texto.lower() not in ("nan", "none", "-", "0", "não", "nao")


def atualizar_status_planilha_original(caminho_planilha: Path, nome_aba, linha_excel: int, idx_col_status: int, novo_status: str):
    """
    Atualiza o Status na planilha original usando openpyxl para preservar 100% da formatação.
    Se o arquivo estiver aberto no Excel, pede para fechar antes de salvar.
    """
    while True:
        try:
            wb = openpyxl.load_workbook(caminho_planilha)
            ws = wb[nome_aba] if isinstance(nome_aba, str) else wb.worksheets[int(nome_aba)]
            ws.cell(row=linha_excel, column=idx_col_status, value=novo_status)
            wb.save(caminho_planilha)
            wb.close()
            print(f" [EXCEL ORIGINAL] Linha {linha_excel} atualizada para Status = '{novo_status}'.")
            return
        except PermissionError:
            print("\n [!] ATENÇÃO: A planilha original está aberta no Excel e bloqueada para gravação!")
            input(" >>> Feche a planilha no Excel e pressione ENTER aqui para salvar... ")


def salvar_na_nova_tabela(dados_linha: dict):
    """Adiciona o projeto recém-atualizado na nova tabela (.xlsx e .txt) para visualização humana."""
    df_novo = pd.DataFrame([dados_linha], columns=COLUNAS_NOVA_TABELA)

    # 1. Atualiza o arquivo Excel (.xlsx) da nova tabela
    if ARQUIVO_NOVA_TABELA_XLSX.exists():
        df_existente = pd.read_excel(ARQUIVO_NOVA_TABELA_XLSX)
        df_combinado = pd.concat([df_existente, df_novo], ignore_index=True)
    else:
        df_combinado = df_novo

    while True:
        try:
            df_combinado.to_excel(ARQUIVO_NOVA_TABELA_XLSX, index=False)
            break
        except PermissionError:
            input(f" [!] Feche o arquivo '{ARQUIVO_NOVA_TABELA_XLSX.name}' no Excel e pressione ENTER... ")

    # 2. Atualiza também o arquivo TXT formatado com '|'
    escrever_cabecalho = not ARQUIVO_NOVA_TABELA_TXT.exists()
    with open(ARQUIVO_NOVA_TABELA_TXT, "a", encoding="utf-8") as f:
        if escrever_cabecalho:
            f.write(" | ".join(COLUNAS_NOVA_TABELA) + "\n")
            f.write("-" * 160 + "\n")
        valores_str = [
            "" if pd.isna(dados_linha.get(c)) or dados_linha.get(c) is None else str(dados_linha.get(c)).strip()
            for c in COLUNAS_NOVA_TABELA
        ]
        f.write(" | ".join(valores_str) + "\n")

    print(f" [NOVA TABELA] Dados transferidos para '{ARQUIVO_NOVA_TABELA_XLSX.name}' e '{ARQUIVO_NOVA_TABELA_TXT.name}'.")


def abrir_tela_nova_associacao(page: Page):
    """Navega até Observatório -> Associação Avaliador - Projeto e clica no botão de nova associação."""
    print(" -> Navegando para 'Associação Avaliador - Projeto'...")
    if page.locator("#menu_toggle").is_visible():
        page.locator("#menu_toggle").click()
        page.wait_for_timeout(500)

    obs = page.get_by_text("Observatório", exact=True).first
    if obs.is_visible():
        obs.click()
        page.wait_for_timeout(500)

    page.get_by_role("link", name="Associação Avaliador - Projeto").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)

    print(" -> Abrindo formulário de associação (+)...")
    page.get_by_role("link").filter(has_text=re.compile(r"^$")).nth(2).click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)


def preencher_associacao_no_opee(page: Page, distribuidora: str, codigo_projeto: str, avaliador: str):
    """
    Preenche Distribuidora, Código do Projeto e Avaliador no OPEE,
    parando ANTES de clicar no botão Salvar.
    """
    abrir_tela_nova_associacao(page)

    # --------------------------------------------------------------------------
    # 1. DISTRIBUIDORA
    # --------------------------------------------------------------------------
    print(f" [1/3] Selecionando Distribuidora: '{distribuidora}'...")
    dropdown_dist = page.locator("span").filter(has_text="Nenhum item selecionado").first
    dropdown_dist.get_by_role("button").first.click()
    page.wait_for_timeout(800)

    # Busca o botão da distribuidora (exata ou contendo o nome) e rola até ele automaticamente
    btn_dist = page.get_by_role("button", name=distribuidora, exact=True)
    if btn_dist.count() == 0:
        btn_dist = page.get_by_role("button").filter(
            has_text=re.compile(rf"^\s*{re.escape(distribuidora)}\s*$", re.IGNORECASE)
        )
    if btn_dist.count() == 0:
        btn_dist = page.get_by_role("button").filter(
            has_text=re.compile(re.escape(distribuidora), re.IGNORECASE)
        )

    btn_dist.first.scroll_into_view_if_needed()
    btn_dist.first.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)  # Aguarda carregar os projetos da distribuidora

    # --------------------------------------------------------------------------
    # 2. PROJETO (Digita na busca se houver e clica na opção correspondente)
    # --------------------------------------------------------------------------
    print(f" [2/3] Selecionando Projeto: '{codigo_projeto}'...")
    dropdown_proj = page.locator("span").filter(has_text="Nenhum item selecionado").first
    dropdown_proj.get_by_role("button").first.click()
    page.wait_for_timeout(800)

    # Se houver caixa de busca dentro do dropdown aberto, filtra pelo código do projeto
    caixa_busca_proj = dropdown_proj.locator("input[type='text'], input.inputFilter").first
    if caixa_busca_proj.count() > 0 and caixa_busca_proj.is_visible():
        caixa_busca_proj.fill(codigo_projeto)
        page.wait_for_timeout(600)

    # Localiza o botão que contém o código do projeto (ex: "PE-00044-0016/2019 Inclusao e...")
    btn_proj = dropdown_proj.get_by_role("button").filter(
        has_text=re.compile(re.escape(codigo_projeto), re.IGNORECASE)
    )
    if btn_proj.count() == 0:
        # Tenta sem o prefixo "PE-" caso esteja só o número
        codigo_num = re.sub(r"^[A-Z]+-", "", codigo_projeto.strip(), flags=re.I)
        btn_proj = dropdown_proj.get_by_role("button").filter(
            has_text=re.compile(re.escape(codigo_num), re.IGNORECASE)
        )

    if btn_proj.count() == 0:
        raise RuntimeError(f"Projeto '{codigo_projeto}' não encontrado na lista da distribuidora '{distribuidora}'.")

    btn_proj.first.scroll_into_view_if_needed()
    btn_proj.first.click()
    page.wait_for_timeout(500)

    # Fecha a lista suspensa de projetos clicando no main (como na sua gravação)
    page.get_by_role("main").click()
    page.wait_for_timeout(600)

    # --------------------------------------------------------------------------
    # 3. AVALIADOR INTERNO (GUILHERME SILVA CAVALCANTE)
    # --------------------------------------------------------------------------
    print(f" [3/3] Selecionando Avaliador: '{avaliador}'...")
    dropdown_aval = page.locator("span").filter(has_text="Nenhum item selecionado").first
    dropdown_aval.get_by_role("button").first.click()
    page.wait_for_timeout(800)

    btn_aval = page.get_by_role("button", name=avaliador, exact=True)
    if btn_aval.count() == 0:
        btn_aval = page.get_by_role("button").filter(
            has_text=re.compile(re.escape(avaliador), re.IGNORECASE)
        )

    btn_aval.first.scroll_into_view_if_needed()
    btn_aval.first.click()
    page.wait_for_timeout(500)

    print(" [PRONTO] Campos preenchidos na tela! O robô parou antes do botão 'Salvar'.")


def executar_associacao(caminho_planilha: Path, nome_aba: str | int = 0):
    inicializar_pastas()

    if not caminho_planilha.exists():
        raise FileNotFoundError(f"Planilha não encontrada: {caminho_planilha}")

    historico = HistoricoDownloads(ARQUIVO_SET_ASSOCIADOS)
    print(f"[HISTÓRICO] {len(historico.projetos)} projeto(s) já associados anteriormente.")

    print(f"[LENDO EXCEL] Carregando '{caminho_planilha.name}'...")
    df = pd.read_excel(caminho_planilha, sheet_name=nome_aba)
    colunas_originais = [str(c).strip() for c in df.columns]
    df.columns = colunas_originais

    mapa_colunas = mapear_colunas_planilha(colunas_originais)
    col_dist = mapa_colunas["Distribuidora"]
    col_cod = mapa_colunas["Codigo do Projeto"]
    col_status = mapa_colunas["Status"]
    col_aneel = mapa_colunas["Avaliado Pela ANEEL"]
    col_nt = mapa_colunas["Nota Técnica"]

    for obrigatoria in (col_dist, col_cod, col_status, col_aneel, col_nt):
        if obrigatoria not in df.columns:
            raise ValueError(
                f"Coluna '{obrigatoria}' não encontrada na planilha.\nColunas disponíveis: {colunas_originais}"
            )

    # Posição da coluna Status no Excel (1-indexed para o openpyxl)
    idx_col_status_excel = colunas_originais.index(col_status) + 1

    # Filtra todos os candidatos elegíveis
    candidatos = []
    for idx_df, row in df.iterrows():
        status_atual = str(row[col_status]).strip() if pd.notna(row[col_status]) else ""
        avaliado_aneel = str(row[col_aneel]).strip() if pd.notna(row[col_aneel]) else ""
        nt_val = row[col_nt]
        cod_bruto = str(row[col_cod]).strip() if pd.notna(row[col_cod]) else ""

        if (
            status_atual.lower() == STATUS_FILTRO.lower()
            and avaliado_aneel.lower() in ("sim", "s")
            and valor_preenchido(nt_val)
            and cod_bruto
        ):
            try:
                chave = chave_projeto(cod_bruto)
            except ValueError:
                chave = cod_bruto.upper()

            if not historico.contem(chave):
                # Linha no Excel é idx_df + 2 (1 do cabeçalho + 1 porque pandas começa em 0)
                candidatos.append((idx_df + 2, chave, row))

    print(f"[FILTRO] Encontrados {len(candidatos)} projeto(s) elegíveis ('{STATUS_FILTRO}' + ANEEL=Sim + NT preenchida).")

    if not candidatos:
        print("[FIM] Nenhum projeto pendente atende aos critérios.")
        return

    with sync_playwright() as p:
        print("Conectando ao Microsoft Edge aberto na porta 9222...")
        browser = p.chromium.connect_over_cdp("http://localhost:9222")
        context = browser.contexts[0]

        page = context.pages[0]
        for aba in context.pages:
            if "opee" in aba.url.lower() or "aneel" in aba.url.lower():
                page = aba
                break

        print(f"[CONECTADO] Aba ativa: {page.title()}")

        for pos, (linha_excel, chave, row) in enumerate(candidatos, start=1):
            distribuidora = str(row[col_dist]).strip()
            codigo_projeto = str(row[col_cod]).strip()
            nota_tecnica = str(row[col_nt]).strip()

            print("\n" + "=" * 75)
            print(f" PROJETO [{pos}/{len(candidatos)}] — {codigo_projeto} ({distribuidora}) | NT: {nota_tecnica}")
            print("=" * 75)

            try:
                preencher_associacao_no_opee(page, distribuidora, codigo_projeto, AVALIADOR_PADRAO)
            except Exception as erro:
                print(f"\n [ERRO] Não foi possível preencher o projeto {codigo_projeto} na tela: {erro}")
                resp_erro = input(" >>> Deseja tentar o próximo projeto da lista? [S/N]: ").strip().lower()
                if resp_erro in ("s", "sim", "y", ""):
                    continue
                else:
                    break

            print("\n" + "-" * 75)
            print(" AÇÃO HUMANA NECESSÁRIA:")
            print(" 1. Confira na tela do Edge se Distribuidora, Projeto e Avaliador estão corretos.")
            print(" 2. Clique manualmente no botão 'Salvar' lá no navegador.")
            print("-" * 75)
            confirmacao = input(
                " >>> Você clicou em 'Salvar' no navegador? [ENTER ou S = Sim, atualizar planilhas | P = Pular sem salvar]: "
            ).strip().lower()

            if confirmacao in ("p", "pular", "n", "nao", "não"):
                print(f" [PULADO] Projeto {codigo_projeto} não foi alterado na planilha.")
            else:
                # 1. Atualiza o Status na planilha original para "Aguardando avaliação interna"
                atualizar_status_planilha_original(
                    caminho_planilha=caminho_planilha,
                    nome_aba=nome_aba,
                    linha_excel=linha_excel,
                    idx_col_status=idx_col_status_excel,
                    novo_status=NOVO_STATUS,
                )

                # 2. Monta o dicionário com as 18 colunas (já com o novo Status) e insere na nova tabela
                dados_transferencia = {}
                for col_destino in COLUNAS_NOVA_TABELA:
                    col_origem = mapa_colunas.get(col_destino, col_destino)
                    dados_transferencia[col_destino] = row[col_origem] if col_origem in row else ""
                dados_transferencia["Status"] = NOVO_STATUS

                salvar_na_nova_tabela(dados_transferencia)

                # 3. Registra no set persistente para nunca mais repetir
                try:
                    historico.registrar(chave)
                except ValueError:
                    # Caso algum código fuja do regex PE/PD/PG, grava direto no arquivo
                    with open(ARQUIVO_SET_ASSOCIADOS, "a", encoding="utf-8") as f:
                        f.write(f"{chave}\n")
                print(f" [SET ATUALIZADO] Projeto {chave} registrado em '{ARQUIVO_SET_ASSOCIADOS.name}'.")

            # Pergunta se quer fazer mais um projeto
            if pos < len(candidatos):
                continuar = input("\n >>> Deseja executar mais um projeto? [S/N]: ").strip().lower()
                if continuar not in ("s", "sim", "y", ""):
                    print("\n[ENCERRADO] Execução finalizada pelo usuário.")
                    break
            else:
                print("\n[FIM] Todos os projetos elegíveis da planilha foram processados!")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Associa avaliador no OPEE para projetos com Nota Técnica e atualiza as planilhas."
    )
    parser.add_argument(
        "--planilha", "-x",
        type=Path,
        default=PLANILHA_ORIGINAL_PADRAO,
        help="Caminho completo da planilha Excel original (.xlsx).",
    )
    parser.add_argument(
        "--aba", "-a",
        default=0,
        help="Nome ou índice da aba da planilha (padrão: 0).",
    )
    args = parser.parse_args()
    executar_associacao(caminho_planilha=args.planilha, nome_aba=args.aba)
