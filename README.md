# G_insumos

from inspect import FrameInfo
import PySimpleGUI as sg
import openpyxl

# Definindo o layout da interface gráfica
sg.theme('lightgrey 1')

# Definindo o layout da janela inicial
layout = [[sg.Text('Gerenciamento de Insumos', font=('Arial', 20), justification='center', size=(30, 2))],
          [sg.Column([[sg.Button('Financeiro', key='financeiro', size=(20, 2)), sg.Button('Almoxarifado', key='almoxarifado', size=(20, 2)), sg.Button('Compras', key='compras', size=(20, 2)), sg.Button('Sair', key='saida_inicio', size=(20, 2))]])]]
# Criando a janela inicial

window = sg.Window('Menu', layout).Finalize()
window.Maximize()


# Loop principal da janela inicial
while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED:
        break
    elif event == 'financeiro':
                # Carregar dados da planilha
        workbook = openpyxl.load_workbook('compras.xlsx')
        sheet = workbook['compras']
        data = [[cell.value for cell in row] for row in sheet.iter_rows(values_only=True)]

            # Opções de forma de pagamento
        opcoes_pagamento = ['pix', 'debito', 'credito a vista', 'credito parcelado', 'faturamento']

            # Função para atualizar a planilha
        def atualizar_planilha(linha, forma_pagamento):
                sheet.cell(row=linha, column=4, value=forma_pagamento)
                workbook.save('compras.xlsx')

        
        layout = [
                [sg.Text('Selecione a forma de pagamento:')],
                [sg.Combo(opcoes_pagamento, size=(20, 1), key='forma_pagamento')],
                [sg.Button('Enviar', key='enviar')],
                [sg.Table(values=data, headings=sheet.columns, max_col_width=25,
                        display_row_numbers=True, auto_size_columns=False, justification='center',
                        num_rows=min(sheet.max_row, 20), key='tabela')]
            ]

        window = sg.Window('Almoxarifado', layout)

        while True:
                event, values = window.read()
                if event == sg.WIN_CLOSED:
                    break
                elif event == 'enviar':
                    forma_pagamento = values['forma_pagamento']
                    linha_selecionada = values['tabela'][0]
                    if linha_selecionada != None:
                        atualizar_planilha(linha_selecionada + 1, forma_pagamento)
                        sg.popup('Forma de pagamento salva com sucesso!')

        window.close()
 
    # Saida inicio##################################
    elif event == 'saida_inicio':
        confirmation_layout = [[sg.Text('Deseja realmente sair?')],
                               [sg.Button('Sim'), sg.Button('Não')]]
                               
        confirmation_window = sg.Window(
            'Confirmação de saída', confirmation_layout)
        while True:
            event_confirm, _ = confirmation_window.read()
            if event_confirm == sg.WINDOW_CLOSED or event_confirm == 'Não':
                break
            elif event_confirm == 'Sim':
                window.close()
            break
        # saida inicio fim ################################################

        # inicio compras ################################
    elif event == 'compras':
            user = sg.popup_get_text('Digite o user:')
            password = sg.popup_get_text('Digite a senha:', password_char='*')

            if password == '2505' and user == 'compras':
                # Criação da janela
                    layout_compras = [
                    [sg.Text('Fornecedor'), sg.InputText(key='fornecedor')],
                    [sg.Text('NF'), sg.InputText(key='nf')],
                    [sg.Text('Nome do cliente'), sg.InputText(key='cliente')],
                    [sg.Text('Valor da compra'), sg.InputText(key='valor')],
                    [sg.Text('Forma de pagamento solicitada'), sg.InputText(key='pagamento')],
                    [sg.Button('Adicionar produtos', key='add')],
                    [sg.Text('Produtos')],
                    [sg.Multiline(key='produtos')],
                    [sg.Text('Anexar PDF')],
                    [sg.FileBrowse()],
                    [sg.Button('Salvar', key='save')],
                    [sg.Button('Buscar', key='buscar')]
                ]

                # Criação da janela
                    window = sg.Window('Pedido de compra', layout_compras)

                # Lista de produtos
                    produtos = []

                # Loop de eventos
                    while True:
                        event, values = window.read()
                        if event == sg.WIN_CLOSED:
                            break
                        elif event == 'add':
                            while True:
                                # Abre uma nova janela para adicionar produtos
                                    layout_produtos = [
                                    [sg.Text('Produto'), sg.InputText(key='produto')],
                                    [sg.Text('Quantidade'), sg.InputText(key='quantidade')],
                                    [sg.Button('Adicionar', key='add_produto')],
                                    [sg.Button('Fechar', key='close_produto')]
                                ]
                                    window_produtos = sg.Window('Adicionar', layout_produtos)

                                    event_produtos, values_produtos = window_produtos.read()
                                    if event_produtos in (sg.WINDOW_CLOSED, 'close_produto'):
                                        window_produtos.close()
                                        break

                                # Adiciona o produto à lista de produtos
                                    produtos.append((values_produtos['produto'], values_produtos['quantidade']))
                                    window_produtos.close()
                                    
                                # Atualiza a lista de produtos na interface gráfica
                                    window['produtos'].update('\n'.join([f'{p[0]} - {p[1]}' for p in produtos]))

                                    window_produtos.close()



                        elif event == 'save':
                            # Salva os dados do pedido em uma planilha do Excel
                            wb = openpyxl.Workbook()
                            ws = wb.active
                            ws.title = 'Compras'
                            ws.append(['Fornecedor', 'NF', 'Cliente', 'Valor', 'Pagamento', 'Produtos'])
                            ws.append([values['fornecedor'], values['nf'], values['cliente'], values['valor'], values['pagamento'], ', '.join([f'{p[0]} - {p[1]}' for p in produtos])])
                            wb.save('compras.xlsx')
                            sg.popup('Pedido salvo com sucesso!')

                        elif event == 'buscar':
                            # Abre uma nova janela para buscar pedidos
                                layout_busca = [
                                [sg.Text('Digite o nome do cliente'), sg.InputText(key='cliente')],
                                [sg.Button('Buscar', key='buscar_pedidos')],
                                [sg.Text('Pedidos')],
                                [sg.Multiline(key='pedidos')]]
                                
                                window_busca = sg.Window('Buscar pedidos', layout_busca)

                        while True:
                            event_busca, values_busca = window_busca.read()

                            if event_busca == sg.WIN_CLOSED:
                                break

                            elif event_busca == 'buscar_pedidos':
                                        # Abre a planilha de pedidos e busca pelo fornecedor
                                        wb = openpyxl.load_workbook('compras.xlsx')
                                        ws = wb.active
                                        pedidos = []

                                        for row in ws.iter_rows(min_row=2):
                                            if row[0].value == values_busca['cliente']:
                                                pedidos.append(f'NF: {row[1].value} - Cliente: {row[2].value} - Valor: {row[3].value} - Pagamento: {row[4].value} - Produtos: {row[5].value}')

                                        # Atualiza a lista de pedidos na interface gráfica
                                        window_busca['pedidos'].update('\n'.join(pedidos)) 
                                        window.close()

    # fim compras #################################

### almoxarifado inicio ########################################################

    elif event == 'almoxarifado' :
        # Pede a senha para o usuário
        user = sg.popup_get_text('Digite o user :')
        password = sg.popup_get_text('Digite a senha:', password_char='*')

        # Verifica se a senha está correta
        if password == '2505' and user == 'almoxarifado':
            window.close()
            layout = [
                [sg.Text('Produto:'), sg.Input(key='produto')],
                [sg.Text('Quantidade:'), sg.Input(key='quantidade')],
                [sg.Text('Localização:'), sg.Input(key='localizacao')],
                [sg.Button('Adicionar', key='add'), sg.Button('Remover', key='rem'),
                 sg.Button('Atualizar', key='upd'), sg.Button(
                 'Listar', key='lst'),
                sg.Button('Buscar', key='search'), sg.Button('Sair', key='sair')],
                [sg.Table(values=[], headings=['Produto', 'Quantidade', 'Localizacao'],
                key='table', justification='left', auto_size_columns=False, col_widths=[50, 30, 100])]]
               

# Criando a janela da interface gráfica
            window = sg.Window('Controle de Estoque', layout).Finalize()
            window.Maximize()

            # Carregando a planilha do Excel
        workbook = openpyxl.load_workbook('Estoque.xlsx')
        sheet = workbook.active

        # Função para buscar os dados da planilha e atualizar a tabela na interface gráfica
    def update_table():
        data = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            data.append(row)
        window['table'].update(values=data)

        # Função para buscar um produto na planilha
    def search_product(produto):
        data = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0] == produto:
                data.append(row)
        if data:
            window['table'].update(values=data)
        else:
            sg.popup('Produto não encontrado.')

        # Função para atualizar os dados do estoque

    def update_stock(produto, quantidade, localizacao):
        sheet.append([produto, quantidade, localizacao])
        workbook.save('Estoque.xlsx')
        update_table()

        # Função para remover os dados do estoque

    def remove_stock(indices):
        for index in reversed(indices):
            sheet.delete_rows(index + 2)
        workbook.save('Estoque.xlsx')
        update_table()

        # Função para atualizar os dados de um produto
    def update_product(indices):
        for index in indices:
            row = sheet[index + 2]
        layout = [[sg.Text('Produto:'), sg.Input(key='produto', default_text=row[0].value)],
                  [sg.Text('Quantidade:'), sg.Input(
                      key='quantidade', default_text=row[1].value)],
                  [sg.Text('Localização:'), sg.Input(
                      key='localizacao', default_text=row[2].value)],
                  [sg.Button('Salvar', key='save')]]
        window_edit = sg.Window('Editar produto', layout)
        while True:
            event_edit, values_edit = window_edit.read()
            if event_edit == sg.WINDOW_CLOSED:
                break
            elif event_edit == 'save':
                row[0].value = values_edit['produto']
                row[1].value = values_edit['quantidade']
                row[2].value = values_edit['localizacao']
                workbook.save('Estoque.xlsx')
                update_table()
                window_edit.close()

        # Loop principal da interface gráfica
    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        elif event == 'sair':
            confirmation_layout = [[sg.Text('Deseja realmente sair?')],
                                   [sg.Button('Sim'), sg.Button('Não')]]
            confirmation_window = sg.Window(
                'Confirmação de saída', confirmation_layout)
            while True:
                event_confirm, _ = confirmation_window.read()
                if event_confirm == sg.WINDOW_CLOSED or event_confirm == 'Não':
                    break
                elif event_confirm == 'Sim':
                    window.close()
                    break
        elif event == 'add':
            produto = values['produto']
            quantidade = values['quantidade']
            localizacao = values['localizacao']
            # Cria uma nova janela para aprovação financeira
            approval_layout = [[sg.Text(f'Produto: {produto}\nQuantidade: {quantidade}\n Localização :{localizacao}\n\nAprovar adição ao estoque?')],
                               [sg.Button('Sim'), sg.Button('Não')]]
            approval_window = sg.Window(
                'Aprovação Financeira', approval_layout)
            # Aguarda a escolha do usuário na janela de aprovação financeira
            while True:
                approval_event, _ = approval_window.read()
                if approval_event == sg.WINDOW_CLOSED:
                    break
                elif approval_event == 'Sim':
                    update_stock(produto, quantidade, localizacao)
                    break
                elif approval_event == 'Não':
                    break
            approval_window.close()
        elif event == 'rem':
            indices = window['table'].SelectedRows
            for index in reversed(indices):
                sheet.delete_rows(index + 2)
            workbook.save('Estoque.xlsx')
            update_table()
        elif event == 'upd':
            indices = window['table'].SelectedRows
            if len(indices) == 1:
                update_product(indices)
            else:
                sg.popup('Selecione apenas uma linha para atualizar.')
        elif event == 'lst':
            update_table()
        elif event == 'search':
            search_product(values['produto'])
        print('Almoxarifado')

        # FIM ALMOXARIFADO####################################################################333
