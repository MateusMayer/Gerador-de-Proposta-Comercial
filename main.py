import tkinter as tk
from tkinter import messagebox, filedialog
import os
import shutil
from datetime import datetime, date
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT  # Para alinhamento do parágrafo
from babel.dates import format_date
import logging
import sys

logging.basicConfig(filename='app.log', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Define o caminho base uma única vez
if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

def referenciascampos():
    data_atual = date.today()
    data_formatada = format_date(data_atual, format='long', locale='pt_BR')

    n001 = n001_entry.get()
    s001 = s001_entry.get()
    s003 = s003_entry.get()
    s004 = s004_entry.get()
    s005 = s005_entry.get()
    c001 = c001_entry.get()
    h001 = h001_entry.get()
    d003 = d003_entry.get()
    v001 = v001_entry.get()

    data_formatada_d002 = datetime.now().strftime("%d/%m/%Y")
    data_formatada_d001 = data_formatada

    referencias = {
        "N001": n001, "S001": s001, "S003": s003, "D001": data_formatada_d001,
        "D002": data_formatada_d002,
        "S004": s004, "S005": s005,
        "C001": c001, "H001": h001, "D003": d003, "V001": v001,
    }
    return referencias

def replace_text_in_runs(runs, code, value):
    for run in runs:
        run.text = run.text.replace(code, value)


def novosmodulos():
    # Inicializa o dicionário de referências com valores de campos estáticos
    referencias = referenciascampos()

    # Itera sobre cada módulo dinâmico para processar seus valores
    for modulo_idx, module_widgets in enumerate(dynamic_modules_widgets, start=2):
        modulo_valores = [widget.get() for widget in module_widgets if isinstance(widget, tk.Entry)]

        # Verifica se o módulo atual tem todos os 4 valores necessários
        if len(modulo_valores) >= 4:
            c, h, d, v = modulo_valores  # Desempacota os valores

            # Atualiza o dicionário de referências com valores do módulo atual
            referencias[f'C{modulo_idx:03}'] = c
            referencias[f'H{modulo_idx:03}'] = h
            referencias[f'D{modulo_idx+2:03}'] = d
            referencias[f'V{modulo_idx:03}'] = v
        else:
            # Se um módulo não tem todos os 4 valores, para o processamento
            logging.error("Modulo não tem todos os 4 valores")
            break
    print("Antes novos modulos:",referencias)
    return referencias

def replace_marker_with_image(documento, marcador, caminho_imagem, width):
    for paragrafo in documento.paragraphs:
        if marcador in paragrafo.text:
            # Limpa o texto do parágrafo removendo o marcador
            texto_limpo = paragrafo.text.replace(marcador, "")
            paragrafo.text = texto_limpo
            # Agora adiciona a imagem ao final do parágrafo
            # Note que isso adicionará a imagem após qualquer texto existente no parágrafo
            paragrafo.add_run().add_picture(caminho_imagem, width=width)

def preencher_linha_tabela(tabela, linha_idx, referencias, config_colunas):
    # Verifica se a tabela tem linhas suficientes
    if len(tabela.rows) <= linha_idx:
        return  # Sai da função se não houver linhas suficientes

    dados = []
    for coluna in config_colunas:
        valor = coluna['prefixo']  # Inicializa com o prefixo, se houver
        if coluna['codigo'] in referencias:
            if coluna['posicao'] == 'antes':
                valor = referencias[coluna['codigo']] + valor
            else:
                valor += referencias[coluna['codigo']]
        dados.append(valor)

    # Preenche a linha especificada com os dados configurados
    for col_num, texto in enumerate(dados):
        celula = tabela.cell(linha_idx, col_num)
        paragrafo = celula.paragraphs[0]
        paragrafo.text = texto
        paragrafo.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY


import os

def escolher_local_salvamento():
    # Obtém o caminho para a pasta de downloads do usuário atual
    downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')

    # Define as opções para a caixa de diálogo
    opcoes = {
        'defaultextension': '.docx',
        'filetypes': [('documentos do Word', '.docx')],
        'initialdir': downloads_path,  # Usa o caminho para a pasta de downloads
        'title': 'Salvar documento como...'
    }

    # Abre a caixa de diálogo e guarda o caminho escolhido pelo usuário
    caminho_salvamento = filedialog.asksaveasfilename(**opcoes)

    # Verifica se um caminho foi selecionado (ou seja, o usuário não cancelou a operação)
    if caminho_salvamento:
        messagebox.showinfo("Alerta", f"Caminho escolhido: {caminho_salvamento}")
        return caminho_salvamento
    else:
        messagebox.showinfo("Erro", f"Caminho Inválido ou erro de permissão.")
        # Retorna None ou trata de maneira adequada se o usuário cancelar a operação
        return None



def atualizar_tabela_com_campos_novos(tabela, referencias):
    # Definição dos códigos de campos esperados para cada módulo novo
    codigos_por_modulo = [
        ['C002', 'H002', 'D004', 'V002'],
        ['C003', 'H003', 'D005', 'V003'],
        ['C004', 'H004', 'D006', 'V004'],
        ['C005', 'H005', 'D007', 'V005'],
    ]

    # Inicializa o índice da linha a ser preenchida na tabela
    linha_idx = 2  # Começando da terceira linha

    for codigos in codigos_por_modulo:
        config_colunas = []
        todos_codigos_presentes = all(codigo in referencias for codigo in codigos)

        # Se todos os códigos deste módulo estão presentes em referencias, prepara para preenchimento
        if todos_codigos_presentes:
            for codigo in codigos:
                if codigo.startswith('C'):  # Para 'Consultor'
                    config_colunas.append({'codigo': codigo, 'prefixo': 'Consultor ', 'posicao': 'depois'})
                elif codigo.startswith('H'):  # Para 'Horas'
                    config_colunas.append({'codigo': codigo, 'prefixo': 'hs', 'posicao': 'antes'})
                elif codigo.startswith('D'):  # Para 'Data'
                    config_colunas.append({'codigo': codigo, 'prefixo': '', 'posicao': 'depois'})
                elif codigo.startswith('V'):  # Para 'Valor'
                    config_colunas.append({'codigo': '', 'prefixo': 'A combinar', 'posicao': 'depois'})

            # Preenche a linha atual da tabela com os dados configurados
            preencher_linha_tabela(tabela, linha_idx, referencias, config_colunas)

            # Incrementa o índice da linha para a próxima iteração
            linha_idx += 1
        else:
            print(codigos)  # Para ver quais códigos estão sendo verificados
            print(referencias.keys())  # Para ver quais chaves estão presentes em referencias

            logging.error("Nem todos os códigos estão presentes.")
            break  # Sai do loop se um conjunto de campos não estiver completo
def add_dynamic_module():
    global dynamic_module_count
    if dynamic_module_count >= 4:
        messagebox.showwarning("Limite Atingido", "Número máximo de 5 módulos atingido.")
        return

    row_base = 10 + dynamic_module_count * 4
    dynamic_module_count += 1

    # Armazenar widgets temporariamente para poder removê-los mais tarde
    module_widgets = []

    labels_texts = [f"Módulo {dynamic_module_count + 1}", f"Horas Módulo {dynamic_module_count + 1}", f"Data Início Módulo {dynamic_module_count + 1}", f"Valor Módulo {dynamic_module_count + 1}"]
    for i, text in enumerate(labels_texts):
        row_offset = row_base + i//2
        column_offset = 0 if i % 2 == 0 else 2
        label = create_label(main_frame, text, row_offset, column_offset)
        entry = create_entry(main_frame, row_offset, column_offset + 1)
        module_widgets.extend([label, entry])

    dynamic_modules_widgets.append(module_widgets)


    add_module_button.grid(row=row_base + 1, column=4, pady=(15, 15))
    upload_logo_button.grid(row=row_base + 5, column=1, pady=(15, 15))
    save_button.grid(row=row_base + 5, column=2, pady=(10, 10))
    return module_widgets

########################################################################################################################
####################################-----Salva o Documento-----#########################################################
def save_document():
    # Capturando os valores dos campos
    global run

    # Define o caminho para a pasta de Downloads do usuário
    # Isso funciona para Windows, macOS e Linux
    #downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')

    # Verifica se algum campo está vazio
    if not all([n001_entry.get().strip(), s001_entry.get().strip(),
                s003_entry.get().strip(), s004_entry.get().strip(), s005_entry.get().strip()]):
        messagebox.showerror("Erro", "Todos os campos devem ser preenchidos.")
        return

    # Combina o caminho da pasta de Downloads com o nome do arquivo
    #documento_final_path = os.path.join(downloads_path, nome_arquivo)
    logging.info("Abrindo diálogo para escolher o local de salvamento")
    try:
        documento_final_path = escolher_local_salvamento()
        if documento_final_path is None:
            raise ValueError("Nenhum caminho de salvamento selecionado pelo usuário.")
        logging.info(f"Caminho de salvamento escolhido: {documento_final_path}")
    except Exception as e:
        logging.error(f"Erro ao escolher o local de salvamento: {e}")
        return

    documento_path = os.path.join(base_path, "Documentos", "Comercial.docx")
    if not documento_path:
        messagebox.showinfo("Erro", f"Problema no caminho de origem do documento")
        logging.error(f"Problema no caminho de origem do documento")
        return None

    # Cria uma cópia do documento original
    documento_copia_path = os.path.join(base_path, "Documentos", "Comercial - Copia.docx")
    try:
        shutil.copy(documento_path, documento_copia_path)
        logging.info(f"Cópia realizada")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao criar cópia do documento: {e}")
        logging.error(f"Erro ao criar cópia do documento")
        return

    #Abre a cópia do documento para edição
    try:
        documento = Document(documento_copia_path)
        logging.info(f"Documento de cópia aberto para edição com sucesso")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao abrir o documento para edição: {e}")
        logging.error(f"Erro ao abrir o documento para edição")
        return
    try:
        referencias1 = novosmodulos()
        print("Após novosmodulos:", referencias1)
    except Exception as e:
        logging.error(f"Erro ao atribuir as referências dos novos módulos: {str(e)}")
        messagebox.showerror("Erro", f"Erro ao atribuir as referências dos novos módulos: {str(e)}")
        return  # Para garantir que a execução não prossiga em caso de erro

    try:
        referencias2 = referenciascampos()
        if(referencias1 == 0):
            referencias = referencias2
            logging.info(f"Referencias iniciais")
        else:
            referencias = referencias1
            logging.info(f"Referencias campos adicionais")
    except Exception as e:
        logging.error(f"Problema na validação das referências")
        return

    for paragrafo in documento.paragraphs:
        for codigo, valor in referencias.items():
            replace_text_in_runs(paragrafo.runs, codigo, valor)

    imagens_para_adicionar = {
        "{Imagem}": (os.path.join(base_path,"Imagens", "image123.png"), Cm(18)),
        "{logo}": (os.path.join(base_path,"Imagens", "logo.png"), Cm(8)),
        "{Certificações}": (os.path.join(base_path,"Imagens", "certificados.png"), Cm(20)),
        "{projetos}": (os.path.join(base_path,"Imagens","projetos.png"), Cm(14)),
        "{clientes}": (os.path.join(base_path,"Imagens","clientes.png"), Cm(15))
        # Adicionar mais imagens conforme necessário
    }

    for marcador, (caminho_imagem, width) in imagens_para_adicionar.items():
        replace_marker_with_image(documento, marcador, caminho_imagem, width)

    for tabela in documento.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for paragrafo in celula.paragraphs:
                    for codigo, valor in referencias.items():
                        replace_text_in_runs(paragrafo.runs, codigo, valor)

    tabela1 = documento.tables[1]  # Acessa a segunda tabela do documento
    try:
        atualizar_tabela_com_campos_novos(tabela1, referencias)
        logging.info(f"Tabela atualizada com campos novos")
    except Exception as e:
        logging.error(f"Erro ao atualizar tabela com os campos novos")
    # Salva o documento editado na pasta que o usuario escolher
    try:
        documento.save(documento_final_path)
        messagebox.showinfo("Sucesso", f"Documento salvo com sucesso em: {documento_final_path}")
        logging.info(f"Documento salvo com sucesso em: {documento_final_path}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar o documento: {e}")
        logging.error(f"Erro ao salvar o documento")
        return

    # Pergunta ao usuário se deseja encerrar o programa
    if messagebox.askyesno("Encerrar", "Deseja encerrar o programa?"):
        app.destroy()  #Encerra o programa

background_color = "#FFFFFF"
text_color = "#000000"
button_color = "#FFA500"

def center_window(width, height):
    screen_width = main.winfo_screenwidth()
    screen_height = main.winfo_screenheight()
    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2)
    main.geometry('%dx%d+%d+%d' % (width, height, x, y))

def create_label(master, text, row, column, pady=5, padx=10, sticky="W"):
    label = tk.Label(master, text=text, bg=background_color, fg=text_color)
    label.grid(row=row, column=column, pady=pady, padx=padx, sticky=sticky)

def create_entry(master, row, column, pady=5, padx=10):
    entry = tk.Entry(master)
    entry.grid(row=row, column=column, pady=pady, padx=padx)
    return entry


def upload_logo():
    # Abre o diálogo de seleção de arquivos para o usuário escolher um arquivo de imagem
    filepath = filedialog.askopenfilename(title="Selecione o arquivo do logo", filetypes=(
    ("PNG files", "*.png"), ("JPEG files", "*.jpg"), ("All files", "*.*")))

    if filepath:  # Verifica se um arquivo foi selecionado
        try:
            # Cria o diretório "Imagens" se não existir
            os.makedirs("Imagens", exist_ok=True)
            # Define o caminho de destino dentro da pasta "Imagens"
            destino = (os.path.join(base_path,"Imagens", "logo.png"))
            # Copia o arquivo selecionado para o destino
            shutil.copy(filepath, destino)
            messagebox.showinfo("Sucesso", "Logo carregado com sucesso na pasta 'Imagens'.")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível carregar o logo: {e}")

dynamic_modules_widgets = []
dynamic_module_count = 0

app = tk.Tk()
# Configura a janela para abrir maximizada
app.state('zoomed')
app.title("Gerador de Proposta T&M")

# Configurando o ícone da janela
app.iconbitmap((os.path.join(base_path,"GUI","icone.ico")))  # Certifique-se de ter o arquivo icone.ico no diretório do projeto
app.configure(bg=background_color)

# Criação do main_frame
main_frame = tk.Frame(app, bg="#FFFFFF")
main_frame.pack(expand=True, fill='both')

# Carregando e exibindo o logotipo
logo = tk.PhotoImage(file=(os.path.join(base_path,"GUI","logotipo fusion.png"))) # Assegure-se de que o arquivo esteja no diretório correto
logo_label = tk.Label(app, image=logo, bg=background_color)
logo_label.pack(pady=(2, 2))


#screen.center_window(650, 800)

main_frame = tk.Frame(app, bg=background_color)
main_frame.pack(expand=True)

create_label(main_frame, "Número da Proposta", 0, 0)
create_label(main_frame, "Nome do Cliente", 1, 0)
create_label(main_frame, "Gerente de Contas", 2, 0)
create_label(main_frame, "Necessidade", 0, 2)
create_label(main_frame, "Nome do Consultor", 1, 2)
create_label(main_frame, "Nome do Executivo Comercial", 2, 0)
create_label(main_frame, "Módulo 1", 6, 0)
create_label(main_frame, "Horas Módulo 1", 6, 2)
create_label(main_frame, "Data Início Módulo 1", 7, 0)
create_label(main_frame, "Valor Módulo 1", 7, 2)

n001_entry = create_entry(main_frame, 0, 1)
s001_entry = create_entry(main_frame, 1, 1)
s003_entry = create_entry(main_frame, 0, 3)
s004_entry = create_entry(main_frame, 1, 3)
s005_entry = create_entry(main_frame, 2, 1)
c001_entry = create_entry(main_frame, 6, 1)
h001_entry = create_entry(main_frame, 6, 3)
d003_entry = create_entry(main_frame, 7, 1)
v001_entry = create_entry(main_frame, 7, 3)

# Adicione o botão para adicionar novos módulos na interface, ajustando sua posição inicial
add_module_button = tk.Button(main_frame, text="+", command=add_dynamic_module, bg=button_color, fg=text_color)
add_module_button.grid(row=7, column=4, pady=(15, 15))

upload_logo_button = tk.Button(main_frame, text="Upload Logo", command=upload_logo, bg=button_color, fg=text_color)
upload_logo_button.grid(row=12, column=1, pady=(10, 10))

save_button = tk.Button(main_frame, text="Salvar Documento", command=save_document, bg=button_color, fg=text_color)
save_button.grid(row=12, column=2, pady=(10, 10))


app.mainloop()
