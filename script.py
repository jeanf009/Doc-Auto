import os
import shutil
import platform
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
from docxtpl import DocxTemplate
from jinja2 import Environment, FileSystemLoader

def processar_documentos():
    dados = {
        'nome_discente': entry_nome.get().strip(),
        'data_hora': entry_data.get().strip(),
        'local_defesa': entry_local.get().strip(),
        'titulo_trabalho': entry_titulo.get().strip(),
        'especializacao': entry_especializacao.get().strip(),
        'area_trabalho': entry_area.get().strip(),
        'nome_orientador': entry_orientador.get().strip(),
        'nome_coorientador': entry_coorientador.get().strip() if chk_var_coorientador.get() == 1 else "",
        'examinadores_internos': entry_internos.get().strip(),
        'examinadores_externos': entry_externos.get().strip(),
        'numero_defesa': entry_numero.get().strip(),
        'resumo_trabalho': entry_resumo.get("1.0", tk.END).strip()
    }
    caminho_anexo = lbl_caminho_anexo.cget("text")

    if not dados['nome_discente']:
        messagebox.showwarning("Aviso", "Preencha Nome do Discente.")
        return

    if not dados['examinadores_internos'] or not dados['examinadores_externos']:
        messagebox.showwarning(
            "Erro de Validação",
            "É obrigatório informar pelo menos um Examinador Interno e um Examinador Externo"
        )
        return
    
    DIR_ATUAL = os.path.dirname(os.path.abspath(__file__))
    
    pasta_base = os.path.join(DIR_ATUAL, "Saida")
    identificador = f"{dados['nome_discente'].replace(' ', '_')}"
    pasta_liberacao = os.path.join(pasta_base, "Liberacao", identificador)
    pasta_comprovacao = os.path.join(pasta_base, "Comprovacao", identificador)

    try:
        os.makedirs(pasta_liberacao, exist_ok=True)
        os.makedirs(pasta_comprovacao, exist_ok=True)
 
        templates_docx = [
            {"template": "doc_defesa.docx", "prefixo_saida": "Declaração_de_Defesa", "gerar": True},
            {"template": "doc_orientador.docx", "prefixo_saida": "Declaração_Orientador", "gerar": True},
            {"template": "doc_banca.docx", "prefixo_saida": "Declaração_Banca_Examinadora", "gerar": True},
            {"template": "doc_convocacao.docx", "prefixo_saida": "Resolução_Banca_Examinadora", "gerar": True},
            {"template": "doc_coorientador.docx", "prefixo_saida": "Declaração_Coorientador", "gerar": bool(dados['nome_coorientador'])},
        ]
        
        for t in templates_docx:
            if not t["gerar"]:
                continue

            caminho_docx_template = os.path.join(DIR_ATUAL, "Templates", t["template"])
            
            if os.path.exists(caminho_docx_template):
                doc = DocxTemplate(caminho_docx_template)
                doc.render(dados)
                
                nome_saida_docx = f"{t['prefixo_saida']}_{identificador}.docx"
                caminho_docx_saida = os.path.join(pasta_liberacao, nome_saida_docx)
                doc.save(caminho_docx_saida)

            else:
                print(f"Aviso: Template não encontrado: {t['template']}")

        caminho_pasta_templates = os.path.join(DIR_ATUAL, "Templates")
        if os.path.exists(os.path.join(caminho_pasta_templates, "chamada.html")):
            env = Environment(loader=FileSystemLoader(caminho_pasta_templates))
            template_html = env.get_template("chamada.html")
            
            html_renderizado = template_html.render(dados)
            
            caminho_html_saida = os.path.join(pasta_liberacao, f"Chamada_{identificador}.html")
            with open(caminho_html_saida, "w", encoding="utf-8") as f:
                f.write(html_renderizado)
        else:
            print("Template HTML não encontrado.")

        if caminho_anexo and caminho_anexo != "Nenhum arquivo selecionado":
            nome_arquivo = os.path.basename(caminho_anexo)
            shutil.copy2(caminho_anexo, os.path.join(pasta_comprovacao, nome_arquivo))

        sistema = platform.system()
        if sistema == "Windows":
            os.startfile(pasta_base)
        elif sistema == "Darwin":
            subprocess.Popen(['open', pasta_base])
        else:
            subprocess.Popen(['xdg-open', pasta_base])

        messagebox.showinfo("Sucesso", "Textos dos documentos gerados!")
        limpar_formulario()

    except Exception as e:
        messagebox.showerror("Erro Crítico", f"Falha no sistema de arquivos ou conversão:\n{str(e)}")


def selecionar_arquivo():
    caminho = filedialog.askopenfilename()
    if caminho:
        lbl_caminho_anexo.config(text=caminho)

def limpar_formulario():
    for entry in entradas_simples:
        entry.delete(0, tk.END)
    
    entry_resumo.delete("1.0", tk.END)

    chk_var_coorientador.set(0)
    toggle_coorientador()

    lbl_caminho_anexo.config(text="Nenhum arquivo selecionado")

def toggle_coorientador():
    if chk_var_coorientador.get() == 1:
        entry_coorientador.config(state="normal")

    else:
        entry_coorientador.config(state="normal")
        entry_coorientador.delete(0, tk.END)
        entry_coorientador.config(state="disabled")

root = tk.Tk()
root.title("GUI - Automatização de Documentos")
root.geometry("850x850") 

lbl_titulo = tk.Label(root, text="Preenchimento de Defesa", font=("Arial", 16, "bold"))
lbl_titulo.pack(pady=(10, 5))

main_frame = tk.Frame(root)
main_frame.pack(fill="both", expand=True, padx=20, pady=5)

font_label = ("Arial", 10, "bold")

campos_iniciais = [
    ("Nome do Discente:", "entry_nome"),
    ("Data-Hora da Defesa:", "entry_data"),
    ("Local Defesa:", "entry_local"),
    ("Título do Trabalho:", "entry_titulo"),
    ("Especialização:", "entry_especializacao"),
    ("Área do Trabalho:", "entry_area"),
    ("Orientador:", "entry_orientador")
]

entradas_dict = {}
entradas_simples = []

for label_text, var_name in campos_iniciais:
    tk.Label(main_frame, text=label_text, font=font_label).pack(anchor="w", pady=(4, 0))
    entry = tk.Entry(main_frame, width=100)
    entry.pack(anchor="w", padx=5)
    entradas_dict[var_name] = entry
    entradas_simples.append(entry)

entry_nome = entradas_dict["entry_nome"]
entry_data = entradas_dict["entry_data"]
entry_local = entradas_dict["entry_local"]
entry_titulo = entradas_dict["entry_titulo"]
entry_especializacao = entradas_dict["entry_especializacao"]
entry_area = entradas_dict["entry_area"]
entry_orientador = entradas_dict["entry_orientador"]

frame_coor = tk.Frame(main_frame)
frame_coor.pack(fill="x", pady=(6, 2))

chk_var_coorientador = tk.IntVar(value=0)
chk_coorientador = tk.Checkbutton(
    frame_coor, 
    text="Habilitar Coorientador", 
    variable=chk_var_coorientador, 
    command=toggle_coorientador,
    font=font_label
)
chk_coorientador.pack(anchor="w")

entry_coorientador = tk.Entry(frame_coor, width=100, state="disabled")
entry_coorientador.pack(anchor="w", padx=5)

campos_finais = [
    ("Examinador(es) Interno(s):", "entry_internos"),
    ("Examinador(es) Externo(s):", "entry_externos"),
    ("Número da Defesa:", "entry_numero")
]

for label_text, var_name in campos_finais:
    tk.Label(main_frame, text=label_text, font=font_label).pack(anchor="w", pady=(4, 0))
    entry = tk.Entry(main_frame, width=100)
    entry.pack(anchor="w", padx=5)
    entradas_dict[var_name] = entry
    entradas_simples.append(entry)

entry_internos = entradas_dict["entry_internos"]
entry_externos = entradas_dict["entry_externos"]
entry_numero = entradas_dict["entry_numero"]


tk.Label(main_frame, text="Resumo:", font=font_label).pack(anchor="w", pady=(4, 0))
entry_resumo = tk.Text(main_frame, width=75, height=4, font=("Arial", 10))
entry_resumo.pack(anchor="w", padx=5, pady=(2, 5))

frame_rodape = tk.Frame(root)
frame_rodape.pack(fill="x", padx=20, pady=(5, 15))

tk.Label(frame_rodape, text="Anexos de Comprovação (Opcional):", font=font_label).pack(anchor="w")

btn_anexo = tk.Button(frame_rodape, text="Procurar...", command=selecionar_arquivo, width=15)
btn_anexo.pack(side="left", pady=5)

lbl_caminho_anexo = tk.Label(frame_rodape, text="Nenhum arquivo selecionado", fg="gray")
lbl_caminho_anexo.pack(side="left", padx=15, pady=5)

btn_gerar = tk.Button(
    root, 
    text="Gerar Documentos", 
    font=("Arial", 12, "bold"), 
    bg="#2a82da", 
    fg="white",
    command=processar_documentos
)
btn_gerar.pack(fill="x", padx=20, pady=(0, 10))

root.mainloop()