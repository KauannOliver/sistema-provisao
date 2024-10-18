import flet as ft
from openpyxl import load_workbook

# Função para salvar os dados no Excel
def salvar_cliente(caminho_arquivo, dados_cliente):
    try:
        # Carregar a planilha existente e a sheet "Impostos"
        planilha = load_workbook(caminho_arquivo)
        sheet = planilha["Impostos"]

        # Adicionar os dados à próxima linha vazia
        sheet.append(dados_cliente)

        # Salvar a planilha atualizada
        planilha.save(caminho_arquivo)
        return True
    except Exception as e:
        print(f"Erro ao salvar cliente: {e}")
        return False

# Função para criar o formulário de cadastro de cliente
def criar_formulario_cliente(page):
    caminho_arquivo = "banco/ProvisaoBD.xlsx"  # Ajuste o caminho conforme necessário

    # Função para salvar os dados do formulário
    def salvar_dados(e):
        try:
            # Obter valores dos campos e formatar os valores de impostos removendo o símbolo de porcentagem
            dados_cliente = [
                cliente_field.value,
                ic_field.value,
                und_negocio_field.value,
                float(icms_field.value.replace(",", ".").replace("%", "").strip()) / 100,  # Remover o "%" e converter
                float(iss_field.value.replace(",", ".").replace("%", "").strip()) / 100,   # Remover o "%" e converter
                float(pis_field.value.replace(",", ".").replace("%", "").strip()) / 100,   # Remover o "%" e converter
                float(cofins_field.value.replace(",", ".").replace("%", "").strip()) / 100,# Remover o "%" e converter
                float(cprb_field.value.replace(",", ".").replace("%", "").strip()) / 100,  # Remover o "%" e converter
            ]

            # Salvar os dados no Excel
            if salvar_cliente(caminho_arquivo, dados_cliente):
                page.show_snack_bar(ft.SnackBar(content=ft.Text("Cliente salvo com sucesso!"), bgcolor=ft.colors.GREEN))
            else:
                page.show_snack_bar(ft.SnackBar(content=ft.Text("Erro ao salvar o cliente."), bgcolor=ft.colors.RED))

            # Fechar o modal
            fechar_modal(page)
        except Exception as ex:
            page.show_snack_bar(ft.SnackBar(content=ft.Text(f"Erro ao salvar o cliente: {ex}"), bgcolor=ft.colors.RED))


    # Função para aplicar a máscara de porcentagem
    def aplicar_mascara_porcentagem(e):
        valor = e.control.value.replace("%", "").replace(",", ".").strip()
        if valor.isdigit() or valor.replace(".", "").isdigit():
            e.control.value = f"{float(valor):.2f}".replace(".", ",") + "%"
        page.update()

    ### Criação dos componentes do formulário ###
    cliente_field = ft.TextField(label="Cliente", width="100%")
    ic_field = ft.TextField(label="I.C", width="100%")
    und_negocio_field = ft.TextField(label="UND Negócio", width="100%")

    # Campos com máscara de porcentagem
    icms_field = ft.TextField(label="ICMS", width="100%", suffix_text="%", on_blur=aplicar_mascara_porcentagem)
    iss_field = ft.TextField(label="ISS", width="100%", suffix_text="%", on_blur=aplicar_mascara_porcentagem)
    pis_field = ft.TextField(label="PIS", width="100%", suffix_text="%", on_blur=aplicar_mascara_porcentagem)
    cofins_field = ft.TextField(label="COFINS", width="100%", suffix_text="%", on_blur=aplicar_mascara_porcentagem)
    cprb_field = ft.TextField(label="CPRB", width="100%", suffix_text="%", on_blur=aplicar_mascara_porcentagem)

    ### Layout do formulário usando ResponsiveRow ###
    form_layout = ft.Column(
        controls=[
            ft.ResponsiveRow([
                ft.Column(col={"sm": 6}, controls=[cliente_field]),
                ft.Column(col={"sm": 3}, controls=[ic_field]),
                ft.Column(col={"sm": 3}, controls=[und_negocio_field]),
            ]),
            ft.ResponsiveRow([
                ft.Column(col={"sm": 4}, controls=[icms_field]),
                ft.Column(col={"sm": 4}, controls=[iss_field]),
                ft.Column(col={"sm": 4}, controls=[pis_field]),
            ]),
            ft.ResponsiveRow([
                ft.Column(col={"sm": 4}, controls=[cofins_field]),
                ft.Column(col={"sm": 4}, controls=[cprb_field]),
            ]),
            ft.ResponsiveRow([
                ft.Column(col={"sm": 3}, controls=[ft.ElevatedButton(
                    text="Salvar",
                    on_click=salvar_dados,
                    style=ft.ButtonStyle(
                        bgcolor=ft.colors.with_opacity(1, "#f36229"),
                        color=ft.colors.WHITE,
                        shape=ft.RoundedRectangleBorder(radius=8),
                    ),
                    width=145,
                    height=50,
                )]),
                ft.Column(col={"sm": 3}, controls=[ft.ElevatedButton(
                    text="Fechar",
                    on_click=lambda e: fechar_modal(page),
                    style=ft.ButtonStyle(
                        bgcolor=ft.colors.with_opacity(1, "#f36229"),
                        color=ft.colors.WHITE,
                        shape=ft.RoundedRectangleBorder(radius=8),
                    ),
                    width=145,
                    height=50,
                )]),
            ],
            alignment=ft.MainAxisAlignment.START,
            spacing=10,
            ),
        ],
        alignment=ft.MainAxisAlignment.CENTER,
        spacing=20,
        width=600,
        height=250
    )

    return form_layout

# Função para fechar o modal
def fechar_modal(page):
    if page.dialog:
        page.dialog.open = False
        page.dialog = None
        page.update()
