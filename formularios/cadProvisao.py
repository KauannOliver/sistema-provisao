import flet as ft
from banco.banco import obter_clientes, carregar_impostos_de_json, salvar_dados_excel
from funcoes.funcoes import aplicar_mascara_data, format_currency, arquivo_selecionado
from datetime import datetime
import os
import re
import locale
import random
import string
import pandas as pd

# Configurando a localização para português do Brasil
locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

def criar_formulario_provisao(page):
    ### Função para calcular impostos e receita líquida ###
    def calcular_impostos(e=None):
        if cliente_dropdown.value and tipo_doc_dropdown.value and receita_bruta_field.value:
                receita_bruta = float(receita_bruta_field.value.replace(".", "").replace(",", ".").replace("R$", "").strip())

                ### Carrega os impostos do JSON ###
                impostos_dict = carregar_impostos_de_json()

                ### Busca o I.C. correspondente ao cliente selecionado ###
                cliente_nome = cliente_dropdown.value
                ic_cliente = None

                for ic, dados in impostos_dict.items():
                    if dados['CLIENTE'] == cliente_nome:
                        ic_cliente = ic
                        break

                if ic_cliente and ic_cliente in impostos_dict:
                    impostos = impostos_dict[ic_cliente]
                    icms = 0
                    iss = 0

                    ### Calcula o ICMS e ISS com base no tipo de documento ###
                    if tipo_doc_dropdown.value == "CTE":
                        icms = receita_bruta * impostos['ICMS']
                    elif tipo_doc_dropdown.value == "NOTA FISCAL":
                        iss = receita_bruta * impostos['ISS']

                    ### Calcula PIS, COFINS, CPRB e Receita Líquida ###
                    base_calculo = receita_bruta - icms
                    pis = base_calculo * impostos['PIS']
                    cofins = base_calculo * impostos['COFINS']
                    cprb = receita_bruta * impostos['CPRB']
                    receita_liquida = receita_bruta - icms - iss - pis - cofins - cprb

                    ### Define os valores nos campos correspondentes ###
                    icms_field.value = f"{icms:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
                    iss_field.value = f"{iss:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
                    pis_field.value = f"{pis:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
                    cofins_field.value = f"{cofins:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
                    cprb_field.value = f"{cprb:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
                    receita_liquida_field.value = f"{receita_liquida:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

                    page.update()
                else:
                    page.show_snack_bar(ft.SnackBar(content=ft.Text(f"I.C. não encontrado para o cliente: {cliente_nome}")))

    ### Função para gerar chave única ###
    def gerar_chave_unica(cliente_nome, data_provisao, todas_chaves_existentes):
        """
        Gera uma chave única para a provisão com base nas regras:
        - Primeira letra do nome do cliente
        - Número do mês
        - Últimos dois dígitos do ano
        - 1 letra aleatória
        - 1 número aleatório
        - 1 letra aleatória
        - 1 número aleatório
        """
        primeira_letra_cliente = cliente_nome[0].upper() if cliente_nome else 'X'
        mes = datetime.strptime(data_provisao, "%d/%m/%Y").strftime('%m')
        ano = datetime.strptime(data_provisao, "%d/%m/%Y").strftime('%y')

        # Gerar os componentes aleatórios
        letra1 = random.choice(string.ascii_uppercase)
        numero1 = random.randint(0, 9)
        letra2 = random.choice(string.ascii_uppercase)
        numero2 = random.randint(0, 9)

        chave = f"{primeira_letra_cliente}{mes}{ano}{letra1}{numero1}{letra2}{numero2}"

        # Verificar se a chave já existe e repetir até encontrar uma chave única
        while chave in todas_chaves_existentes:
            letra1 = random.choice(string.ascii_uppercase)
            numero1 = random.randint(0, 9)
            letra2 = random.choice(string.ascii_uppercase)
            numero2 = random.randint(0, 9)
            chave = f"{primeira_letra_cliente}{mes}{ano}{letra1}{numero1}{letra2}{numero2}"

        return chave

    ### Função para salvar os dados no Excel ###
    def salvar_dados(e):
        try:
            # Carregar a planilha existente
            caminho_planilha = "banco/ProvisaoBD.xlsx"

            # Obter o I.C. e UND Negócio do cliente selecionado
            impostos_dict = carregar_impostos_de_json()
            cliente_nome = cliente_dropdown.value
            ic_cliente = None
            und_negocio = None

            for ic, dados in impostos_dict.items():
                if dados['CLIENTE'] == cliente_nome:
                    ic_cliente = ic
                    und_negocio = dados['UND NEGOCIO']
                    break

            if ic_cliente is None:
                page.show_snack_bar(ft.SnackBar(content=ft.Text("I.C. não encontrado para o cliente."), bgcolor=ft.colors.RED))
                return

            # Formatar a data para o formato mmmm/aa em português
            data_provisao_formatada = datetime.strptime(data_field.value, "%d/%m/%Y").strftime("%B/%y").lower()

            # Carregar todas as chaves existentes da planilha
            df = pd.read_excel(caminho_planilha, sheet_name='Provisões')
            todas_chaves_existentes = df['CHAVE'].tolist()

            # Gerar a nova chave
            chave = gerar_chave_unica(cliente_nome, data_field.value, todas_chaves_existentes)

            # Gerar um link para o arquivo no SharePoint
            link_arquivo = f"Link para o arquivo: [Clique aqui]({caminho_arquivo})" if caminho_arquivo else "Nenhum arquivo selecionado"

            # Adicionar os dados à próxima linha vazia
            nova_linha = [
                data_provisao_formatada,
                cliente_dropdown.value,
                und_negocio,
                ic_cliente,
                tipo_doc_dropdown.value,
                chave,
                num_doc_field.value,
                classificacao_dropdown.value,
                receita_bruta_field.value.replace("R$", "").strip(),
                icms_field.value,
                iss_field.value,
                pis_field.value,
                cofins_field.value,
                cprb_field.value,
                receita_liquida_field.value,
                observacao_field.value,
                link_arquivo
            ]

            mensagem = salvar_dados_excel(caminho_planilha, nova_linha)
            cor_mensagem = ft.colors.GREEN if "sucesso" in mensagem else ft.colors.RED
            page.show_snack_bar(ft.SnackBar(content=ft.Text(mensagem), bgcolor=cor_mensagem))
        except Exception as ex:
            page.show_snack_bar(ft.SnackBar(content=ft.Text(f"Erro ao salvar o cadastro: {ex}"), bgcolor=ft.colors.RED))

    ### Criação dos componentes do formulário ###
    data_field = ft.TextField(
        label="Data Provisão",
        hint_text="DD/MM/AAAA",
        width="100%",
    )

    cliente_dropdown = ft.Dropdown(
        label="Cliente",
        hint_text="Selecione o cliente",
        options=[ft.dropdown.Option(cliente) for cliente in obter_clientes()],
        width="100%",
        on_change=calcular_impostos,
        bgcolor=ft.colors.WHITE
    )

    tipo_doc_dropdown = ft.Dropdown(
        label="Tipo Documento",
        hint_text="Selecione o tipo de documento",
        options=[
            ft.dropdown.Option("CTE"),
            ft.dropdown.Option("NOTA FISCAL"),
            ft.dropdown.Option("N/D")
        ],
        width="100%",
        on_change=calcular_impostos,
        bgcolor=ft.colors.WHITE
    )

    num_doc_field = ft.TextField(
        label="Número Documento",
        hint_text="Digite o número do documento",
        width="100%"
    )

    classificacao_dropdown = ft.Dropdown(
        label="Classificação",
        hint_text="Selecione a classificação",
        options=[
            ft.dropdown.Option("CONTÁBIL"),
            ft.dropdown.Option("GERENCIAL")
        ],
        width="100%",
        bgcolor=ft.colors.WHITE
    )

    receita_bruta_field = ft.TextField(
        label="Receita Bruta",
        hint_text="Digite o valor da receita bruta",
        width="100%",
        on_blur=calcular_impostos
    )

    icms_field = ft.TextField(
        label="ICMS",
        read_only=True,
        value="Automático",
        width="100%"
    )

    iss_field = ft.TextField(
        label="ISS",
        read_only=True,
        value="Automático",
        width="100%"
    )

    pis_field = ft.TextField(
        label="PIS",
        read_only=True,
        value="Automático",
        width="100%"
    )

    cofins_field = ft.TextField(
        label="COFINS",
        read_only=True,
        value="Automático",
        width="100%"
    )

    cprb_field = ft.TextField(
        label="CPRB",
        read_only=True,
        value="Automático",
        width="100%"
    )

    receita_liquida_field = ft.TextField(
        label="Receita Líquida",
        read_only=True,
        value="Automático",
        width="100%"
    )

    observacao_field = ft.TextField(
        label="Observação",
        hint_text="Digite observações adicionais",
        multiline=True,
        max_lines=5,
        width="100%"
    )

    # Aplicar a função de máscara de data
    data_field.on_change = lambda e: aplicar_mascara_data(e, page)

    caminho_arquivo = None
    file_picker = ft.FilePicker(on_result=lambda e: arquivo_selecionado(e, page, file_picker))
    page.overlay.append(file_picker)

    ### Layout do formulário usando ResponsiveRow ###
    form_layout = ft.Column(
        controls=[
            ft.ResponsiveRow([
                ft.Column(col={"sm": 3}, controls=[data_field]),
                ft.Column(col={"sm": 6}, controls=[cliente_dropdown]),
                ft.Column(col={"sm": 3}, controls=[tipo_doc_dropdown]),
            ]),
            ft.ResponsiveRow([
                ft.Column(col={"sm": 4}, controls=[num_doc_field]),
                ft.Column(col={"sm": 4}, controls=[classificacao_dropdown]),
                ft.Column(col={"sm": 4}, controls=[receita_bruta_field]),
            ]),
            ft.ResponsiveRow([
                ft.Column(col={"sm": 2}, controls=[icms_field]),
                ft.Column(col={"sm": 2}, controls=[iss_field]),
                ft.Column(col={"sm": 2}, controls=[pis_field]),
                ft.Column(col={"sm": 2}, controls=[cofins_field]),
                ft.Column(col={"sm": 2}, controls=[cprb_field]),
                ft.Column(col={"sm": 2}, controls=[receita_liquida_field]),
            ]),
            ft.ResponsiveRow([
                ft.Column(col={"sm": 12}, controls=[observacao_field]),
            ]),
            ft.ResponsiveRow([
                ft.Column(col={"sm": 2}, controls=[ft.ElevatedButton(
                    text="Salvar",
                    on_click=salvar_dados,
                    style=ft.ButtonStyle(
                        bgcolor=ft.colors.with_opacity(1, "#f36229"),
                        color=ft.colors.WHITE,
                        shape=ft.RoundedRectangleBorder(radius=8),
                    ),
                    width=145,
                    height=50,
                ),]),
                ft.Column(col={"sm": 2}, controls=[ft.ElevatedButton(
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
        width=750,
        height=320,
    )

    return form_layout

### Função para fechar o modal ###
def fechar_modal(page):
    if page.dialog:
        page.dialog.open = False
        page.dialog = None
        page.update()
