import os
import shutil
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
        'nome_coorientador': entry_coorientador.get().strip(),
        'examinadores_internos': entry_internos.get().strip(),
        'examinadores_externos': entry_externos.get().strip(),
        'numero_defesa': entry_numero.get().strip(),
        'resumo_trabalho': entry_resumo.get().strip()
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
            {"template": "declaracao11.docx", "prefixo_saida": "Declaracao1"},
            {"template": "declaracao22.docx", "prefixo_saida": "Declaracao2"},
            {"template": "declaracao33.docx", "prefixo_saida": "Declaracao3"},
            {"template": "declaracao44.docx", "prefixo_saida": "Declaracao4"},
            {"template": "declaracao55.docx", "prefixo_saida": "Declaracao5"},
        ]
        
        for t in templates_docx:
            caminho_docx_template = os.path.join(DIR_ATUAL, "Templates", t["template"])
            
            if os.path.exists(caminho_docx_template):
                doc = DocxTemplate(caminho_docx_template)
                doc.render(dados)
                
                nome_saida_docx = f"{t['prefixo_saida']}_{identificador}.docx"
                caminho_docx_saida = os.path.join(pasta_liberacao, nome_saida_docx)
                doc.save(caminho_docx_saida)

                # subprocess.run([
                #     "soffice", "--headless", "--invisible", "--nologo", "--nodefault", "--nofirststartwizard",
                #     "--convert-to", "pdf",
                #     caminho_docx_saida, "--outdir", pasta_liberacao
                # ], check=True)

                try:
                    processo = subprocess.run([
                        "soffice", "--headless", "--convert-to", "pdf",
                        caminho_docx_saida, "--outdir", pasta_liberacao
                    ], capture_output=True, text=True, check=True)
                except subprocess.CalledProcessError as e:
                    
                    erro_detalhado = f"Erro no LibreOffice: {e.stderr}"
                    print(erro_detalhado) 
                    raise Exception(erro_detalhado)
                
               
                if os.path.exists(caminho_docx_saida):
                    os.remove(caminho_docx_saida)
            else:
                raise FileNotFoundError(f"Template não encontrado em: {caminho_docx_template}")

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

        subprocess.Popen(['xdg-open', pasta_base])
        messagebox.showinfo("Sucesso", "Documentos (PDF e HTML) gerados!")
        limpar_formulario()

    except Exception as e:
        messagebox.showerror("Erro Crítico", f"Falha no sistema de arquivos ou conversão:\n{str(e)}")

def selecionar_arquivo():
    caminho = filedialog.askopenfilename()
    if caminho:
        lbl_caminho_anexo.config(text=caminho)

def limpar_formulario():
    for entry in entradas:
        entry.delete(0, tk.END)
    lbl_caminho_anexo.config(text="Nenhum arquivo selecionado")

root = tk.Tk()
root.title("GUI - Automatização de Documentos")
root.geometry("1600x900")
root.configure(padx=30, pady=10)

campos = ["Nome do Discente:", "Data-Hora da Defesa:", "Local Defesa:", "Título do Trabalho:", "Especialização:", "Área do Trabalho:",
          "Orientador:", "Coorientador (opcional):", "Examinador(es) Interno(s): ", "Examinador(es) Externo(s): ",
            "Número da Defesa:", "Resumo:"]
entradas = []

for label_text in campos:
    tk.Label(root, text=label_text).pack(anchor="w")
    entry = tk.Entry(root, width=50)
    entry.pack(pady=2)
    entradas.append(entry)

(entry_nome, entry_data, entry_local, entry_titulo, entry_especializacao, entry_area,
  entry_orientador, entry_coorientador, entry_internos, entry_externos, entry_numero, entry_resumo) = entradas

tk.Label(root, text="Anexos de Comprovação:").pack(anchor="w", pady=(15, 0))
tk.Button(root, text="Procurar...", command=selecionar_arquivo).pack(anchor="w", pady=2)
lbl_caminho_anexo = tk.Label(root, text="Nenhum arquivo selecionado", fg="gray")
lbl_caminho_anexo.pack(anchor="w")

tk.Button(root, text="Gerar PDF e HTML", bg="#2a82da", fg="white", font=("Arial", 10, "bold"), command=processar_documentos).pack(fill="x", pady=20)

root.mainloop()