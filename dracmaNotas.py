from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from playwright.sync_api import (
    Playwright,
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
)


# ============================================================
# CONFIGURAÇÕES
# ============================================================

URL_DRACMA = "https://dracma.sdasystems.org/accounts-payable/create?entityId=322"

BASE_DIR = Path(__file__).resolve().parent

PASTA_CACHE = BASE_DIR / "cache_dracma"

ARQUIVO_EXCEL = BASE_DIR / "Pasta1.xlsx"
ABA_EXCEL = "Planilha1"

PASTA_NOTAS = BASE_DIR / "notas"


# ============================================================
# XPATHS DO DRACMA
# ============================================================

XPATH_TIPO_DOCUMENTO = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/"
    "compose/div[2]/div[1]/basic-section/div/form/div[2]/"
    "dropdown-wrapper/div/div/input"
)

XPATH_EMITENTE = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/"
    "compose/div[2]/div[1]/basic-section/div/form/div[3]/div/"
    "person-search/div/div/dropdown-wrapper/div/div/input"
)

XPATH_DATA_EMISSAO = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/"
    "compose/div[2]/div[1]/basic-section/div/form/div[4]/div[1]/"
    "div[1]/date-picker/div/input"
)

XPATH_NUMERO = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/"
    "compose/div[2]/div[1]/basic-section/div/form/div[4]/div[1]/"
    "div[2]/input"
)

XPATH_VALOR = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/"
    "compose/div[2]/div[1]/basic-section/div/form/div[4]/div[2]/"
    "div[1]/currency-input2/div/input"
)

XPATH_FINALIDADE = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/"
    "compose/div[2]/div[1]/basic-section/div/form/div[4]/div[4]/textarea"
)

XPATH_ADICIONAR_ANEXO = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/"
    "compose/div[2]/div[2]/div[2]/div[2]/attachments/compose/"
    "div/div[1]/div"
)

XPATH_SALVAR = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/"
    "compose/div[2]/div[1]/basic-section/div/div/div[2]/button"
)

XPATH_CRIAR_CATEGORIA_E_CENTRO_CUSTO = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/"
    "compose/div[3]/div[1]/ap-cost-div/div[1]/div/button"
)

XPATH_CATEGORIA = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/"
    "compose/div[3]/div[1]/ap-cost-div/div[2]/compose/div/div/div[1]/"
    "div[1]/dropdown-wrapper/div/div/input"
)

XPATH_CENTRO_CUSTO = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/"
    "compose/div[3]/div[1]/ap-cost-div/div[2]/compose/div/div/div[1]/"
    "div[2]/dropdown-wrapper/div/div/input"
)

XPATH_SALVAR_CATEGORIA_E_CENTRO_CUSTO = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/"
    "compose/div[3]/div[1]/ap-cost-div/div[1]/div/div/button[4]"
)

XPATH_FORMA_PAGAMENTO = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/"
    "compose/div[3]/div[2]/payment/div/div/compose/div/div/div[1]/"
    "payment-fields/div/compose/div[1]/dropdown-wrapper/div/div/input"
)

XPATH_CODIGO_PAGAMENTO = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/"
    "compose/div[3]/div[2]/payment/div/div/compose/div/div/div[1]/"
    "payment-fields/div/compose/div[2]/div[2]/div[1]/div/input"
)

XPATH_SALVAR_PAGAMENTO = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/compose/div[3]/div[2]/payment/div/div/compose/payment-header/div/div[1]/button[4]"
) # CLICAR

XPATH_ENVIAR = (
    "xpath=/html/body/navigation/div/main/div/router-view/router-view/compose/div[4]/button[2]"
) # CLICAR E RECOMEÇAR LOOP


# ============================================================
# EXCEL
# ============================================================

def obter_mapa_colunas(ws) -> dict:
    colunas = {}

    for celula in ws[1]:
        if celula.value is not None:
            colunas[str(celula.value).strip()] = celula.column

    return colunas


def ler_documento_excel(numero_linha: int = 2) -> dict:
    if not ARQUIVO_EXCEL.exists():
        raise FileNotFoundError(
            f"Arquivo Excel não encontrado: {ARQUIVO_EXCEL}"
        )

    wb = load_workbook(
        ARQUIVO_EXCEL,
        data_only=True,
    )

    if ABA_EXCEL not in wb.sheetnames:
        wb.close()
        raise ValueError(
            f"A aba '{ABA_EXCEL}' não existe no arquivo."
        )

    ws = wb[ABA_EXCEL]
    colunas = obter_mapa_colunas(ws)

    colunas_necessarias = [
        "Nome do Documento",
        "Tipo de Documento",
        "Emitente",
        "Data de Emissão",
        "Número",
        "Valor",
        "Finalidade",
        "Categorias de Gasto",
        "Centros de Custo",
        "Forma de Pagamento",
        "Campo de Pagamento",
    ]

    faltando = [
        nome
        for nome in colunas_necessarias
        if nome not in colunas
    ]

    if faltando:
        wb.close()
        raise ValueError(
            "Faltam colunas no Excel: "
            + ", ".join(faltando)
        )

    dados = {}

    for nome_coluna in colunas_necessarias:
        dados[nome_coluna] = ws.cell(
            row=numero_linha,
            column=colunas[nome_coluna],
        ).value

    wb.close()

    if not dados["Nome do Documento"]:
        raise ValueError(
            f"A linha {numero_linha} não possui Nome do Documento."
        )

    return dados


# ============================================================
# FORMATAÇÃO
# ============================================================

def texto_seguro(valor) -> str:
    if valor is None:
        return ""

    return str(valor).strip()


def formatar_data_para_dracma(valor) -> str:
    if valor is None:
        return ""

    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")

    texto = str(valor).strip()

    for formato in (
        "%d/%m/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(
                texto,
                formato,
            ).strftime("%d/%m/%Y")
        except ValueError:
            pass

    return texto


def formatar_valor_para_dracma(valor) -> str:
    if valor is None:
        return ""

    if isinstance(valor, (int, float)):
        return f"{valor:.2f}".replace(".", ",")

    texto = str(valor).strip()

    if texto.startswith("R$"):
        texto = texto.replace("R$", "").strip()

    return texto


# ============================================================
# LOGS / UTILITÁRIOS PLAYWRIGHT
# ============================================================

def etapa(nome: str):
    print()
    print("-" * 60)
    print(f"[ETAPA] {nome}")
    print("-" * 60)


def clicar(
    page,
    locator,
    nome: str,
    timeout: int = 30000,
):
    print(f"[AGUARDANDO] {nome}")

    locator.wait_for(
        state="visible",
        timeout=timeout,
    )

    locator.scroll_into_view_if_needed()
    locator.click()

    print(f"[OK] Clique: {nome}")


def preencher_campo(
    page,
    locator,
    valor,
    nome_campo: str,
    pressionar_enter: bool = False,
):
    valor = texto_seguro(valor)

    print(f"[AGUARDANDO] {nome_campo}")

    locator.wait_for(
        state="visible",
        timeout=30000,
    )

    locator.scroll_into_view_if_needed()
    locator.click()
    locator.fill(valor)

    if pressionar_enter:
        locator.press("Enter")

    print(f"[OK] {nome_campo}: {valor}")


def preencher_dropdown(
    page,
    locator,
    valor,
    nome_campo: str,
    espera_ms: int = 1200,
):
    valor = texto_seguro(valor)

    if not valor:
        print(
            f"[AVISO] {nome_campo} está vazio no Excel. "
            "Campo ignorado."
        )
        return False

    print(f"[AGUARDANDO] {nome_campo}")

    locator.wait_for(
        state="visible",
        timeout=30000,
    )

    locator.scroll_into_view_if_needed()
    locator.click()
    locator.fill(valor)

    page.wait_for_timeout(espera_ms)

    locator.press("ArrowDown")
    locator.press("Enter")

    print(f"[OK] {nome_campo}: {valor}")

    return True


def anexar_pdf(page, nome_documento: str):
    caminho_pdf = PASTA_NOTAS / str(nome_documento)

    if not caminho_pdf.exists():
        raise FileNotFoundError(
            f"PDF não encontrado: {caminho_pdf}"
        )

    botao_anexo = page.locator(
        XPATH_ADICIONAR_ANEXO
    )

    print("[AGUARDANDO] Botão Adicionar Anexo")

    botao_anexo.wait_for(
        state="visible",
        timeout=30000,
    )

    botao_anexo.scroll_into_view_if_needed()

    with page.expect_file_chooser(
        timeout=10000,
    ) as file_chooser_info:
        botao_anexo.click()

    file_chooser_info.value.set_files(
        str(caminho_pdf)
    )

    print(
        f"[OK] PDF anexado: {caminho_pdf.name}"
    )


def manter_navegador_aberto(context):
    print()
    print("=" * 60)
    print("NAVEGADOR CONTINUA ABERTO")
    print("=" * 60)
    print('Digite "fechar" para encerrar.')
    print("=" * 60)

    while True:
        comando = input(
            "\nComando: "
        ).strip().lower()

        if comando == "fechar":
            print()
            print("Fechando navegador...")
            break

        print(
            'Comando não reconhecido. '
            'Digite "fechar" para encerrar.'
        )

    context.close()


# ============================================================
# AUTOMAÇÃO
# ============================================================

def run(playwright: Playwright) -> None:
    context = None

    try:
        # ====================================================
        # 1. LER DADOS DO EXCEL
        # ====================================================

        dados = ler_documento_excel(
            numero_linha=2
        )

        nome_documento = texto_seguro(
            dados["Nome do Documento"]
        )

        tipo_documento = texto_seguro(
            dados["Tipo de Documento"]
        )

        emitente = texto_seguro(
            dados["Emitente"]
        )

        data_emissao = formatar_data_para_dracma(
            dados["Data de Emissão"]
        )

        numero = texto_seguro(
            dados["Número"]
        )

        valor = formatar_valor_para_dracma(
            dados["Valor"]
        )

        finalidade = texto_seguro(
            dados["Finalidade"]
        )

        categoria = texto_seguro(
            dados["Categorias de Gasto"]
        )

        centro_custo = texto_seguro(
            dados["Centros de Custo"]
        )

        forma_pagamento = texto_seguro(
            dados["Forma de Pagamento"]
        )

        codigo_pagamento = texto_seguro(
            dados["Campo de Pagamento"]
        )

        print()
        print("=" * 60)
        print("DADOS CARREGADOS DO EXCEL")
        print("=" * 60)
        print(f"Documento: {nome_documento}")
        print(f"Tipo de Documento: {tipo_documento}")
        print(f"Emitente: {emitente}")
        print(f"Data de Emissão: {data_emissao}")
        print(f"Número: {numero}")
        print(f"Valor: {valor}")
        print(f"Finalidade: {finalidade}")
        print(f"Categoria de Gasto: {categoria}")
        print(f"Centro de Custo: {centro_custo}")
        print(f"Forma de Pagamento: {forma_pagamento}")
        print(f"Campo de Pagamento: {codigo_pagamento}")
        print("=" * 60)

        # ====================================================
        # 2. CACHE
        # ====================================================

        PASTA_CACHE.mkdir(
            parents=True,
            exist_ok=True,
        )

        # ====================================================
        # 3. NAVEGADOR COM CACHE
        # ====================================================

        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(PASTA_CACHE),
            headless=False,
            no_viewport=True,
            args=[
                "--window-size=1920,1080",
                "--window-position=0,0",
            ],
        )

        # ====================================================
        # 4. GUIA
        # ====================================================

        if context.pages:
            page = context.pages[0]
        else:
            page = context.new_page()

        # ====================================================
        # 5. ABRIR DRACMA
        # ====================================================

        etapa("ABRIR DRACMA")

        page.goto(
            URL_DRACMA,
            wait_until="domcontentloaded",
            timeout=60000,
        )

        page.locator(
            XPATH_TIPO_DOCUMENTO
        ).wait_for(
            state="visible",
            timeout=300000,
        )

        print("[OK] Formulário carregado.")

        # ====================================================
        # 6. DADOS PRINCIPAIS
        # ====================================================

        etapa("DADOS PRINCIPAIS")

        preencher_dropdown(
            page,
            page.locator(XPATH_TIPO_DOCUMENTO),
            tipo_documento,
            "Tipo de Documento",
        )

        preencher_dropdown(
            page,
            page.locator(XPATH_EMITENTE),
            emitente,
            "Emitente",
        )

        preencher_campo(
            page,
            page.locator(XPATH_DATA_EMISSAO),
            data_emissao,
            "Data de Emissão",
        )

        preencher_campo(
            page,
            page.locator(XPATH_NUMERO),
            numero,
            "Número",
        )

        preencher_campo(
            page,
            page.locator(XPATH_VALOR),
            valor,
            "Valor",
        )

        preencher_campo(
            page,
            page.locator(XPATH_FINALIDADE),
            finalidade,
            "Finalidade",
        )

        # ====================================================
        # 7. ANEXO
        # ====================================================

        etapa("ANEXO")

        anexar_pdf(
            page,
            nome_documento,
        )

        # ====================================================
        # 8. SALVAR DADOS PRINCIPAIS
        # ====================================================

        etapa("SALVAR DADOS PRINCIPAIS")

        clicar(
            page,
            page.locator(XPATH_SALVAR),
            "Salvar dados principais",
        )

        print(
            "[AGUARDANDO] 2 segundos para o Dracma "
            "carregar Categoria/Centro de Custo..."
        )

        page.wait_for_timeout(2000)

        print("[OK] Espera concluída.")

        # ====================================================
        # 9. CRIAR CATEGORIA + CENTRO DE CUSTO
        # ====================================================

        etapa("CRIAR CATEGORIA E CENTRO DE CUSTO")




        page.wait_for_timeout(20000)
        try:
            clicar(
                page,
                page.locator(
                    XPATH_CRIAR_CATEGORIA_E_CENTRO_CUSTO
                ),
                "Criar Categoria/Centro de Custo",
                timeout=5000,
            )

            print(
                "[AGUARDANDO] 1 segundo para os campos "
                "de Categoria/Centro de Custo aparecerem..."
            )

            page.wait_for_timeout(1000)

        except PlaywrightTimeoutError:
            print(
                "[AVISO] Não foi possível clicar em "
                "XPATH_CRIAR_CATEGORIA_E_CENTRO_CUSTO."
            )
            print(
                "[CONTINUANDO] Pulando diretamente para XPATH_CATEGORIA."
            )

        # ====================================================
        # 10. CATEGORIA
        # ====================================================

        etapa("CATEGORIA DE GASTO")

        preencher_dropdown(
            page=page,
            locator=page.locator(
                XPATH_CATEGORIA
            ),
            valor=categoria,
            nome_campo="Categoria de Gasto",
        )

        # ====================================================
        # 11. CENTRO DE CUSTO
        # ====================================================

        etapa("CENTRO DE CUSTO")

        preencher_dropdown(
            page=page,
            locator=page.locator(
                XPATH_CENTRO_CUSTO
            ),
            valor=centro_custo,
            nome_campo="Centro de Custo",
        )

        # ====================================================
        # 12. SALVAR CATEGORIA + CENTRO DE CUSTO
        # ====================================================

        etapa("SALVAR CATEGORIA E CENTRO DE CUSTO")

        clicar(
            page,
            page.locator(
                XPATH_SALVAR_CATEGORIA_E_CENTRO_CUSTO
            ),
            "Salvar Categoria/Centro de Custo",
        )

        print(
            "[OK] Categoria e Centro de Custo salvos."
        )

        page.wait_for_timeout(1000)

        # ====================================================
        # 13. PAGAMENTO
        # ====================================================

        etapa("PAGAMENTO")

        preencher_dropdown(
            page=page,
            locator=page.locator(
                XPATH_FORMA_PAGAMENTO
            ),
            valor=forma_pagamento,
            nome_campo="Forma de Pagamento",
        )

        preencher_campo(
            page=page,
            locator=page.locator(
                XPATH_CODIGO_PAGAMENTO
            ),
            valor=codigo_pagamento,
            nome_campo="Campo de Pagamento",
            pressionar_enter=True,
        )

        # ====================================================
        # 14. SALVAR PAGAMENTO
        # ====================================================

        etapa("SALVAR PAGAMENTO")

        clicar(
            page,
            page.locator(
                XPATH_SALVAR_PAGAMENTO
            ),
            "Salvar Pagamento",
        )

        print()
        print("=" * 60)
        print("AUTOMAÇÃO CONCLUÍDA COM SUCESSO")
        print("=" * 60)

    except PlaywrightTimeoutError as erro:
        print()
        print("=" * 60)
        print("[ERRO] TIMEOUT DO PLAYWRIGHT")
        print("=" * 60)
        print(
            "A automação ficou esperando um elemento "
            "que não apareceu no tempo definido."
        )
        print()
        print("Detalhes:")
        print(erro)
        print("=" * 60)

    except Exception as erro:
        print()
        print("=" * 60)
        print("[ERRO] A AUTOMAÇÃO FOI INTERROMPIDA")
        print("=" * 60)
        print(f"Tipo do erro: {type(erro).__name__}")
        print(f"Mensagem: {erro}")
        print("=" * 60)

    finally:
        if context is not None:
            manter_navegador_aberto(
                context
            )


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)