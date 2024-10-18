import flet as ft

### Função para fechar o modal ###
def fechar_modal(page):
    if page.dialog:
        page.dialog.open = False
        page.dialog = None
        page.update()

### Função para criar botões no menu lateral ###
def botao_menu_lateral(text, on_click):
    return ft.ElevatedButton(
        text,
        style=ft.ButtonStyle(
            color=ft.colors.BLACK,
            bgcolor=ft.colors.WHITE,
            shape=ft.RoundedRectangleBorder(radius=8),
        ),
        width=230,
        height=60,
        on_click=on_click,
    )

### Função para formatar valores monetários ###
def format_currency(value):
    return f"R$ {value:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

### Função para aplicar a máscara de data ###
def aplicar_mascara_data(e, page):
    texto = e.control.value
    texto = texto.replace("/", "")
    if len(texto) > 2:
        texto = texto[:2] + "/" + texto[2:]
    if len(texto) > 5:
        texto = texto[:5] + "/" + texto[5:]
    e.control.value = texto
    page.update()

### Função chamada quando um arquivo é selecionado ###
def arquivo_selecionado(e, page, file_picker):
    if file_picker.result is not None:
        caminho_arquivo = file_picker.result.files[0].path
        page.show_snack_bar(ft.SnackBar(content=ft.Text(f"Arquivo selecionado: {file_picker.result.files[0].name}")))
        return caminho_arquivo
    return None
