from customtkinter import *
from tkinter import filedialog, messagebox, simpledialog
from PIL import Image
from proposta_btk import criar_pdf
import os

app = CTk()
app.geometry("600x480")
app.resizable(0,0)
app.configure(background_color='#131313')

side_img_data = Image.open('Third_Solution\Templates\side_image.png')
side_img = CTkImage(dark_image=side_img_data, light_image=side_img_data, size=(300, 480))
CTkLabel(master=app, text="", image=side_img).pack(expand=True, side="left")
frame = CTkFrame(master=app, width=300, height=480, fg_color='#131313')
frame.pack_propagate(0)
frame.pack(expand=True, side="right")

# Função para selecionar o arquivo Excel
def select_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[('Arquivos Excel', '*.xlsx *.xls')])
    if file_path:
        excel_path_label.configure(text=file_path)  # Atualiza o label com o caminho do arquivo

# Função para selecionar o local de salvamento
def select_save_location():
    folder_path = filedialog.askdirectory()
    if folder_path:
        save_path_label.configure(text=folder_path)  # Atualiza o label com o caminho da pasta

# Função para gerar o PDF
def generate_pdf():
    excel_file = excel_path_label.cget("text")
    save_path = save_path_label.cget("text")
    if excel_file and save_path:
        file_name = simpledialog.askstring("Salvar como", "Digite o nome do arquivo:", parent=app)
        if file_name:
            full_path = os.path.join(save_path, f"{file_name}.pdf")
            try:
                criar_pdf(excel_file, full_path)
                messagebox.showinfo('Sucesso', 'Seu PDF foi gerado com sucesso!')
            except Exception as e:
                messagebox.showerror('Erro', 'Falha ao gerar PDF.')
        else:
            messagebox.showinfo('Informação', 'Operação cancelada.')
    else:
        messagebox.showerror('Erro', 'Por favor, selecione o arquivo Excel e o local de salvamento.')

CTkLabel(master=frame, text='BTK SOLUTIONS', text_color="#F27F1B", anchor="w", justify="left", font=("Arial Bold", 24)).pack(anchor="w", pady=(50, 5), padx=(25, 0))
CTkLabel(master=frame, text="Orçamento automático em PDF", text_color="#7E7E7E", anchor="w", justify="left", font=("Arial Bold", 12)).pack(anchor="w", padx=(25, 0))

# Botão para selecionar o arquivo Excel
CTkButton(master=frame, text="Selecione o arquivo Excel", command=select_excel_file, fg_color="#EEEEEE", text_color="#000000", width=225).pack(anchor="w", pady=(38, 0), padx=(25, 0))
excel_path_label = CTkLabel(master=frame, text="", text_color="#CFCFCF", anchor="w", justify="left", font=("Arial", 12))
excel_path_label.pack(anchor="w", padx=(25, 0))

# Botão para selecionar o local de salvamento
CTkButton(master=frame, text="Selecione o local de salvamento", command=select_save_location, fg_color="#EEEEEE", text_color="#000000", width=225).pack(anchor="w", pady=(21, 0), padx=(25, 0))
save_path_label = CTkLabel(master=frame, text="", text_color="#CFCFCF", anchor="w", justify="left", font=("Arial", 12))
save_path_label.pack(anchor="w", padx=(25, 0))

# Botão para gerar o PDF
CTkButton(master=frame, text="Gerar", command=generate_pdf, fg_color="#CD5F00", hover_color="#E44982", font=("Arial Bold", 15), text_color="#ffffff", width=225).pack(anchor="w", pady=(40, 0), padx=(25, 0))

app.mainloop()
