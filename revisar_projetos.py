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
# 1. CONFIGURAÇÕES DE CAMINHOS E STATUS
# ==============================================================================
PLANILHA_ORIGINAL_PADRAO = Path(r"C:\automacao_opee\planilha_empresas.xlsx")

# Set dos projetos que já foram associados a você na etapa anterior
ARQUIVO_SET_ASSOCIADOS = ONEDRIVE_BASE_DIR / "projetos_associados.txt"

# Novo set para registrar projetos cuja revisão já foi concluída (evita repetir)
ARQUIVO_SET_REVISADOS = ONEDRIVE_BASE_DIR / "projetos_revisados.txt"

# Opcional: Arquivo TXT gerado pelo matching_notas.py (usado como fallback de consulta)
ARQUIVO_MATCHING_TXT = ONEDRIVE_BASE_DIR / "05_Matching_Notas" / "dados_pareados_opee.txt"

STATUS_EM_AVALIACAO = "Em avaliação interna"
STATUS_FINAL_VAZIO = ""  # Deixa o Status em branco/vazio na planilha ao concluir


def normalizar_texto(texto: str) -> str:
    """Remove acentos e espaços extras para localizar nomes de colunas com segurança."""
    if not texto:
        return ""
    sem_acento = "".join(
        c for c in unicodedata.normalize("NFD", str(texto))
        if unicodedata.category(c) != "Mn"
    )
    limpo = re.sub(r"[^a-zA-Z0-9]+", " ", sem_acento).strip().lower()
    return re.sub(r"(.)\1+", r"\1", limpo)


def encontrar_coluna(colunas_planilha: list[str], nomes_possiveis: list[str]) -> str:
    """Retorna o nome exato da coluna na planilha correspondente a um dos nomes possíveis."""
    mapa = {normalizar_texto(c): c for c in colunas_planilha}
    for candidato in nomes_possiveis:
        if candidato in colunas_planilha:
            return candidato
        norm = normalizar_texto(candidato)
        if norm in mapa:
            return mapa[norm]
    raise ValueError(f"Nenhuma das colunas {nomes_possiveis} foi encontrada em: {colunas_planilha}")


def carregar_fallback_matching_txt(caminho_txt: Path) -> dict[str, tuple[str, str]]:
    """
    Caso o arquivo dados_pareados_opee.txt exista, carrega um dicionário
    {chave_projeto: (nota_tecnica, valor_aprovado)} como fonte auxiliar.
    """
    dados = {}
    if not caminho_txt.exists():
        return dados

    for linha in caminho_txt.read_text(encoding="utf-8-sig").splitlines():
        if "|" not in linha or linha.startswith("Distribuidora") or linha.startswith("---"):
            continue
        partes = [p.strip() for p in linha.split("|")]
        if len(partes) >= 4:
            _, cod_proj, nt, valor = partes[:4]
            try:
                chave = chave_projeto(cod_proj)
            except ValueError:
                chave = cod_proj.upper()
            dados[chave] = (nt, valor)
    return dados


def formatar_texto_nota_tecnica(nt_bruta: str) -> str:
    """
    Garante que o texto seja 'Avaliado pela Nota Técnica XXX/2026'.
    Se na planilha estiver apenas '589', transforma em '589/2026'.
    """
    nt = str(nt_bruta).strip()
    if nt.lower().startswith("avaliado pela"):
        return nt
    # Se for apenas número inteiro (ex: '589' ou '589.0'), adiciona '/2026'
    if re.match(r"^\d+(\.0)?$", nt):
        nt = f"{int(float(nt))}/2026"
    return f"Avaliado pela Nota Técnica {nt}"


def formatar_valor_para_campo(valor_bruto) -> str:
    """
    Formata o Valor Aprovado para preenchimento no campo #txtValorAGlosar.
    Converte floats do Excel (ex: 1250430.5) para formato numérico BR ('1250430,50').
    """
    if pd.isna(valor_bruto) or valor_bruto is None:
        return ""
    if isinstance(valor_bruto, (int, float)):
        return f"{ float(valor_bruto):.2f}".replace(".", ",")

    texto = str(valor_bruto).strip()
    texto = re.sub(r"[R$\s]", "", texto)
    # Se vier como '123456.78' (ponto decimal sem vírgula), troca ponto por vírgula
    if "." in texto and "," not in texto:
        partes = texto.split(".")
        if len(partes) == 2 and len(partes[1]) <= 2:
            texto = f"{partes[0]},{partes[1]:0<2}"
    return texto


def atualizar_status_planilha(caminho_planilha: Path, nome_aba, linha_excel: int, idx_col_status: int, novo_status: str):
    """
    Altera a célula de Status na planilha original via openpyxl sem perder formatação.
    """
    while True:
        try:
            wb = openpyxl.load_workbook(caminho_planilha)
            ws = wb[nome_aba] if isinstance(nome_aba, str) else wb.worksheets[int(nome_aba)]
            # Se novo_status for string vazia "", grava None para deixar a célula realmente vazia no Excel
            valor_celula = novo_status if novo_status != "" else None
            ws.cell(row=linha_excel, column=idx_col_status, value=valor_celula)
            wb.save(caminho_planilha)
            wb.close()
            label_log = f"'{novo_status}'" if novo_status != "" else "VAZIO (em branco)"
            print(f" [PLANILHA ATUALIZADA] Linha {linha_excel} -> Status alterado para {label_log}.")
            return
        except PermissionError:
            print("\n [!] ATENÇÃO: A planilha principal está aberta no Excel!")
            input(" >>> Feche a planilha no Excel e pressione ENTER para salvar a mudança de Status... ")


def navegar_e_buscar_projeto(page: Page, codigo_projeto: str):
    """
    Abre Observatório -> Repositório de Projetos e pesquisa pelo Código do Projeto.
    """
    print(f" -> Navegando até 'Repositório de Projetos' para buscar: {codigo_projeto}...")
    if page.locator("#menu_toggle").is_visible():
        page.locator("#menu_toggle").click()
        page.wait_for_timeout(400)

    obs = page.get_by_text("Observatório", exact=True).first
    if obs.is_visible():
        obs.click()
        page.wait_for_timeout(400)

    page.get_by_role("link", name="Repositório de Projetos").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1200)

    # Abre a gaveta de busca
    botao_lupa = page.get_by_title("Buscar Projeto").first
    caixa_codigo = page.get_by_role("textbox").first

    if not caixa_codigo.is_visible():
        botao_lupa.click()
        caixa_codigo.wait_for(state="visible", timeout=10000)
        page.wait_for_timeout(500)

    # Garante que o filtro de Status (combobox nth(1)) não restrinja a busca caso estivesse filtrado antes
    try:
        combobox_status = page.get_by_role("combobox").nth(1)
        if combobox_status.is_visible():
            combobox_status.select_option("")
    except Exception:
        pass

    # Tenta buscar primeiro com o código exato da planilha (ou sem o prefixo PE- se necessário)
    caixa_codigo.click()
    caixa_codigo.fill(codigo_projeto)
    page.get_by_role("button", name="Buscar").click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)

    # Se não apareceu nenhum botão de avaliação, tenta buscar apenas a parte numérica (ex: '0044-0016/2019')
    btn_iniciar = page.get_by_title("Iniciar avaliação do projeto")
    btn_avaliar = page.get_by_title("Avaliar projeto internamente")

    if btn_iniciar.count() == 0 and btn_avaliar.count() == 0:
        cod_curto = re.sub(r"^[A-Z]+-0?", "", codigo_projeto.strip(), flags=re.I)
        if cod_curto != codigo_projeto:
            print(f" -> Refazendo busca usando formato curto: '{cod_curto}'...")
            if not caixa_codigo.is_visible():
                botao_lupa.click()
                caixa_codigo.wait_for(state="visible", timeout=5000)
            caixa_codigo.click()
            caixa_codigo.fill(cod_curto)
            page.get_by_role("button", name="Buscar").click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(1500)


def iniciar_e_abrir_avaliacao_interna(
    page: Page,
    caminho_planilha: Path,
    nome_aba,
    linha_excel: int,
    idx_col_status: int,
):
    """
    Verifica se é necessário clicar em 'Iniciar avaliação do projeto' -> 'Sim'.
    Atualiza a planilha para 'Em avaliação interna' e clica em 'Avaliar projeto internamente'.
    """
    btn_iniciar = page.get_by_title("Iniciar avaliação do projeto").first
    btn_avaliar = page.get_by_title("Avaliar projeto internamente").first

    if btn_iniciar.is_visible():
        print(" -> [1ª VEZ] Clicando em 'Iniciar avaliação do projeto'...")
        btn_iniciar.click()
        page.wait_for_timeout(800)

        btn_sim = page.get_by_role("button", name="Sim").first
        btn_sim.wait_for(state="visible", timeout=10000)
        btn_sim.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1500)

        # Atualiza imediatamente na planilha para "Em avaliação interna"
        atualizar_status_planilha(
            caminho_planilha=caminho_planilha,
            nome_aba=nome_aba,
            linha_excel=linha_excel,
            idx_col_status=idx_col_status,
            novo_status=STATUS_EM_AVALIACAO,
        )
    else:
        print(" -> Projeto já teve avaliação iniciada anteriormente (sem popup 'Iniciar').")
        atualizar_status_planilha(
            caminho_planilha=caminho_planilha,
            nome_aba=nome_aba,
            linha_excel=linha_excel,
            idx_col_status=idx_col_status,
            novo_status=STATUS_EM_AVALIACAO,
        )

    print(" -> Clicando em 'Avaliar projeto internamente'...")
    btn_avaliar.wait_for(state="visible", timeout=15000)
    btn_avaliar.click()
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(1500)


def preencher_abas_avaliacao(page: Page, texto_nota_tecnica: str, valor_projeto: str):
    """
    Preenche 'Avaliado pela Nota Técnica XXX/2026' em todas as abas textuais
    e preenche o valor em '#txtValorAGlosar' na aba 'Custos do Projeto'.
    """
    print(f" -> Preenchendo abas com: '{texto_nota_tecnica}' | Valor: '{valor_projeto}'")

    # 1. Aba Inicial (#regraVigente)
    print("    [1/7] Preenchendo Regra Vigente (#regraVigente)...")
    campo_regra = page.locator("#regraVigente").get_by_role("textbox").first
    campo_regra.wait_for(state="visible", timeout=15000)
    campo_regra.click()
    campo_regra.fill(texto_nota_tecnica)
    page.wait_for_timeout(400)

    # 2. Aba Benefícios (#benef)
    print("    [2/7] Preenchendo aba 'Benefícios' (#benef)...")
    page.get_by_role("tab", name="Benefícios").click()
    page.wait_for_timeout(500)
    campo_benef = page.locator("#benef").get_by_role("textbox").first
    campo_benef.click()
    campo_benef.fill(texto_nota_tecnica)
    page.wait_for_timeout(400)

    # 3. Aba M&V (#mev)
    print("    [3/7] Preenchendo aba 'M&V' (#mev)...")
    page.get_by_role("tab", name="M&V").click()
    page.wait_for_timeout(500)
    campo_mev = page.locator("#mev").get_by_role("textbox").first
    campo_mev.click()
    campo_mev.fill(texto_nota_tecnica)
    page.wait_for_timeout(400)

    # 4. Aba Viabilidade Econômica (#viabEconomica)
    print("    [4/7] Preenchendo aba 'Viabilidade Econômica' (#viabEconomica)...")
    page.get_by_role("tab", name="Viabilidade Econômica").click()
    page.wait_for_timeout(500)
    campo_viab = page.locator("#viabEconomica").get_by_role("textbox").first
    campo_viab.click()
    campo_viab.fill(texto_nota_tecnica)
    page.wait_for_timeout(400)

    # 5. Aba Razoabilidade de Custos (#razoabCustos)
    print("    [5/7] Preenchendo aba 'Razoabilidade de Custos' (#razoabCustos)...")
    page.get_by_role("tab", name="Razoabilidade de Custos").click()
    page.wait_for_timeout(500)
    campo_razoab = page.locator("#razoabCustos").get_by_role("textbox").first
    campo_razoab.click()
    campo_razoab.fill(texto_nota_tecnica)
    page.wait_for_timeout(400)

    # 6. Aba Custos do Projeto (#txtValorAGlosar) -> Recebe o VALOR!
    print(f"    [6/7] Preenchendo aba 'Custos do Projeto' (#txtValorAGlosar) com valor '{valor_projeto}'...")
    page.get_by_role("tab", name="Custos do Projeto").click()
    page.wait_for_timeout(500)
    campo_valor = page.locator("#txtValorAGlosar").first
    campo_valor.wait_for(state="visible", timeout=10000)
    campo_valor.click()
    campo_valor.fill(valor_projeto)
    page.wait_for_timeout(400)

    # 7. Aba Parecer Conclusivo (#parecerConcl)
    print("    [7/7] Preenchendo aba 'Parecer Conclusivo' (#parecerConcl)...")
    page.get_by_role("tab", name="Parecer Conclusivo").click()
    page.wait_for_timeout(500)
    campo_parecer = page.locator("#parecerConcl").get_by_role("textbox").first
    campo_parecer.click()
    campo_parecer.fill(texto_nota_tecnica)
    page.wait_for_timeout(400)

    print(" [PRONTO] Todas as 7 abas foram preenchidas! Aguardando sua revisão final.")


def executar_revisao(caminho_planilha: Path, nome_aba: str | int = 0):
    inicializar_pastas()

    if not caminho_planilha.exists():
        raise FileNotFoundError(f"Planilha principal não encontrada: {caminho_planilha}")

    # Carrega o set de projetos associados (criado pelo associar_avaliador.py)
    set_associados = HistoricoDownloads(ARQUIVO_SET_ASSOCIADOS)
    # Carrega o set de projetos já revisados/concluídos por este script
    set_revisados = HistoricoDownloads(ARQUIVO_SET_REVISADOS)
    # Fallback do TXT de matching (caso algum dado complementar venha de lá)
    mapa_matching_txt = carregar_fallback_matching_txt(ARQUIVO_MATCHING_TXT)

    print(f"[SETS] Projetos no set de associados: {len(set_associados.projetos)} | Já revisados: {len(set_revisados.projetos)}")

    df = pd.read_excel(caminho_planilha, sheet_name=nome_aba)
    colunas_originais = [str(c).strip() for c in df.columns]
    df.columns = colunas_originais

    col_cod = encontrar_coluna(colunas_originais, ["Codigo do Projeto", "Código do Projeto"])
    col_status = encontrar_coluna(colunas_originais, ["Status"])
    col_nt = encontrar_coluna(colunas_originais, ["Nota Técnica", "Nota Tecnica"])
    col_valor = encontrar_coluna(
        colunas_originais,
        ["Valor Aprovado", "Valor Correto no OPEE", "Investimento Real"],
    )

    idx_col_status_excel = colunas_originais.index(col_status) + 1

    # Seleciona os projetos para revisar:
    # 1. Que estejam no set 'projetos_associados.txt' (ou cujo Status na planilha seja 'Aguardando avaliação interna' / 'Em avaliação interna')
    # 2. E que ainda NÃO estejam no set 'projetos_revisados.txt'
    fila_revisao = []
    for idx_df, row in df.iterrows():
        cod_bruto = str(row[col_cod]).strip() if pd.notna(row[col_cod]) else ""
        if not cod_bruto:
            continue

        try:
            chave = chave_projeto(cod_bruto)
        except ValueError:
            chave = cod_bruto.upper()

        if set_revisados.contem(chave):
            continue

        status_atual = str(row[col_status]).strip().lower() if pd.notna(row[col_status]) else ""
        esta_no_set_associados = set_associados.contem(chave)
        status_indica_avaliacao = status_atual in (
            "aguardando avaliação interna",
            "aguardando avaliacao interna",
            "em avaliação interna",
            "em avaliacao interna",
        )

        if esta_no_set_associados or status_indica_avaliacao:
            nt_planilha = str(row[col_nt]).strip() if pd.notna(row[col_nt]) else ""
            val_planilha = row[col_valor] if pd.notna(row[col_valor]) else ""

            # Se faltar NT ou Valor na planilha, consulta o TXT de matching como fallback
            if (not nt_planilha or val_planilha == "") and chave in mapa_matching_txt:
                nt_txt, val_txt = mapa_matching_txt[chave]
                if not nt_planilha:
                    nt_planilha = nt_txt
                if val_planilha == "":
                    val_planilha = val_txt

            fila_revisao.append((idx_df + 2, chave, cod_bruto, nt_planilha, val_planilha))

    print(f"[FILA] {len(fila_revisao)} projeto(s) prontos para revisão no OPEE.")
    if not fila_revisao:
        print("[FIM] Nenhum projeto pendente de revisão encontrado.")
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

        for pos, (linha_excel, chave, codigo_projeto, nt_bruta, val_bruto) in enumerate(fila_revisao, start=1):
            texto_nt = formatar_texto_nota_tecnica(nt_bruta)
            valor_formatado = formatar_valor_para_campo(val_bruto)

            print("\n" + "=" * 75)
            print(f" REVISANDO PROJETO [{pos}/{len(fila_revisao)}]: {codigo_projeto}")
            print(f" - Texto Nota Técnica : {texto_nt}")
            print(f" - Valor (Custos)     : {valor_formatado}")
            print("=" * 75)

            try:
                # 1. Busca o projeto no Repositório de Projetos
                navegar_e_buscar_projeto(page, codigo_projeto)

                # 2. Clica em 'Iniciar avaliação' (se 1ª vez), muda Status na planilha para 'Em avaliação interna'
                #    e entra em 'Avaliar projeto internamente'
                iniciar_e_abrir_avaliacao_interna(
                    page=page,
                    caminho_planilha=caminho_planilha,
                    nome_aba=nome_aba,
                    linha_excel=linha_excel,
                    idx_col_status=idx_col_status_excel,
                )

                # 3. Preenche todas as 7 abas e para antes do botão final
                preencher_abas_avaliacao(page, texto_nt, valor_formatado)

            except Exception as erro:
                print(f"\n [ERRO] Falha ao preencher avaliação do projeto {codigo_projeto}: {erro}")
                tentar_prox = input(" >>> Deseja ir para o próximo projeto? [S/N]: ").strip().lower()
                if tentar_prox in ("s", "sim", "y", ""):
                    continue
                else:
                    break

            print("\n" + "-" * 75)
            print(" AÇÃO HUMANA NECESSÁRIA:")
            print(" 1. Confira os campos preenchidos na tela do Edge.")
            print(" 2. Clique nos botões finais ('Salvar' e 'Concluir') lá no navegador.")
            print("-" * 75)
            confirmacao = input(
                " >>> Você já clicou em Salvar e Concluir? [ENTER ou S = Sim, limpar Status na planilha | P = Pular]: "
            ).strip().lower()

            if confirmacao in ("p", "pular", "n", "nao", "não"):
                print(f" [MANTIDO] O projeto {codigo_projeto} permaneceu como '{STATUS_EM_AVALIACAO}' na planilha.")
            else:
                # Limpa o Status na planilha principal (deixa vazio)
                atualizar_status_planilha(
                    caminho_planilha=caminho_planilha,
                    nome_aba=nome_aba,
                    linha_excel=linha_excel,
                    idx_col_status=idx_col_status_excel,
                    novo_status=STATUS_FINAL_VAZIO,
                )
                # Registra no set de projetos revisados para nunca repetir
                set_revisados.registrar(chave)
                print(f" [CONCLUÍDO] Projeto {chave} salvo em '{ARQUIVO_SET_REVISADOS.name}'.")

            if pos < len(fila_revisao):
                continuar = input("\n >>> Deseja revisar o próximo projeto? [S/N]: ").strip().lower()
                if continuar not in ("s", "sim", "y", ""):
                    print("\n[ENCERRADO] Revisão pausada pelo usuário.")
                    break
            else:
                print("\n[FIM] Todos os projetos da fila foram revisados com sucesso!")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Realiza a revisão interna no OPEE dos projetos associados e atualiza o Status na planilha."
    )
    parser.add_argument(
        "--planilha", "-x",
        type=Path,
        default=PLANILHA_ORIGINAL_PADRAO,
        help="Caminho completo da planilha Excel principal (.xlsx).",
    )
    parser.add_argument(
        "--aba", "-a",
        default=0,
        help="Nome ou índice da aba da planilha (padrão: 0).",
    )
    args = parser.parse_args()
    executar_revisao(caminho_planilha=args.planilha, nome_aba=args.aba)
