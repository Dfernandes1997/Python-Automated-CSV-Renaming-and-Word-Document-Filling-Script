# Importar as bibliotecas necessárias
from pathlib import Path
from docx2pdf import convert
import openpyxl
from docxtpl import DocxTemplate
import os
from tkinter import Tk, messagebox, Label, filedialog

# Criar a função para abrir o explorador de arquivos e selecionar o arquivo
def select_excel_file():
    Tk().withdraw()  # Esconde a janela principal do Tkinter
    excel_file = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Excel Files", "*.xlsx;*.xls")])
    return excel_file

# Função para mostrar a janela "A gerar documentos..."
def show_processing_window():
    processing_window = Tk()
    processing_window.title("A Gerar...")
    
    # Mensagem a exibir enquanto os documentos são gerados
    label = Label(processing_window, text="A gerar documentos, por favor aguarde...", padx=20, pady=20)
    label.pack()
    
    return processing_window

# Definir os paths
base_dir = Path("C:/Scripts")  # Definir o diretório fixo
word_template_path = base_dir / "ZP.docx"  # Caminho fixo para o ZP.docx

# Pedir para selecionar o arquivo Excel
excel_path = select_excel_file()
if not excel_path:
    print("Não selecionou nenhum arquivo Excel.")
    exit()

# O diretório para o output será o mesmo que o do Excel selecionado
output_dir = Path(excel_path).parent / "OUTPUT"  # Criar a pasta OUTPUT no diretório do Excel selecionado
output_dir.mkdir(exist_ok=True)

# Carregar dados do Excel
workbook = openpyxl.load_workbook(excel_path)
sheet = workbook.active
list_values = list(sheet.values)

# Mostrar a janela "A gerar documentos..."
processing_window = show_processing_window()

# Função para processar os dados do Excel
def process_files():
    for value_tuple in list_values[1:]:  # Ignorar a primeira linha (cabeçalho)
        if value_tuple[0] and value_tuple[6]:  # Certificar que os campos essenciais não estão vazios
            # Carregar o template Word
            doc = DocxTemplate(word_template_path)

            # Renderizar o template com os dados específicos da linha do Excel
            doc.render({
                "OT": value_tuple[0],
                "Descrição": value_tuple[6],
            })

            # Guardar o documento renderizado
            output_path = output_dir / f"TR_{value_tuple[0]}"  # Nome do arquivo sem extensão .docx
            doc.save(f"{output_path}.docx")

            # Converter para PDF
            convert(f"{output_path}.docx")

            # Excluir o arquivo .docx após a conversão para PDF
            os.remove(f"{output_path}.docx")

            print(f"Generated and converted document for OT: {value_tuple[0]}")
        else:
            print(f"Skipping row with missing values: {value_tuple}")

    # Fechar a janela
    processing_window.destroy()

    # Mostrar mensagem de conclusão com o caminho do diretório
    messagebox.showinfo("Processo concluído", f"Todos os documentos foram gerados com sucesso!\n\nOs arquivos estão localizados em:\n{output_dir}")

# Executar a função de processamento após a iniciar a janela
processing_window.after(100, process_files)

# Iniciar o loop principal do Tkinter
processing_window.mainloop()
