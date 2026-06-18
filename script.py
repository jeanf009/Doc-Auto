import os
import re
import shutil
import platform
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
from docxtpl import DocxTemplate
from jinja2 import Environment, FileSystemLoader

arquivos_selecionados = []

MESES_EXTENSO = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
    7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
}

# Dicionário estendido para converter o nome do mês em número se necessário
MESES_NOME_PARA_NUM = {v: k for k, v in MESES_EXTENSO.items()}

NUMEROS_EXTENSO = {
    1: "um", 2: "dois", 3: "três", 4: "quatro", 5: "cinco", 6: "seis", 7: "sete", 8: "oito", 9: "nove", 10: "dez",
    11: "onze", 12: "doze", 13: "treze", 14: "catorze", 15: "quinze", 16: "dezesseis", 17: "dezessete", 18: "dezoito",
    19: "dezenove", 20: "vinte", 21: "vinte e um", 22: "vinte e dois", 23: "vinte e três", 24: "vinte e quatro",
    25: "vinte e cinco", 26: "vinte e seis", 27: "vinte e sete", 28: "vinte e oito", 29: "vinte e nove", 30: "trinta", 31: "trinta e um"
}

def converter_entrada_livre_extenso(texto_data):
    texto_limpo = texto_data.lower()
    
    busca_dia = re.search(r'\b(\d{1,2})\b', texto_limpo)
    dia = int(busca_dia.group(1)) if busca_dia else 1
    
    mes = 6
    for nome_mes, num_mes in MESES_NOME_PARA_NUM.items():
        if nome_mes in texto_limpo:
            mes = num_mes
            break
    else:
        busca_mes_num = re.search(r'[/.-](\d{1,2})[/.-]', texto_limpo)
        if busca_mes_num:
            mes = int(busca_mes_num.group(1))

    busca_ano = re.search(r'\b(\d{4})\b', texto_limpo)
    ano = int(busca_ano.group(1)) if busca_ano else 2026

    busca_hora = re.search(r'(?:às\s+|h\s*|:\s*)(\d{1,2})', texto_limpo)
    if not busca_hora:
        busca_hora = re.search(r'\b(\d{1,2})h', texto_limpo)
    
    hora = int(busca_hora.group(1)) if busca_hora else 14
    
    busca_minutos = re.search(r'(?::|h)(\d{2})\b', texto_limpo)
    minuto = int(busca_minutos.group(1)) if busca_minutos else 0

    dia_txt = "primeiro" if dia == 1 else NUMEROS_EXTENSO.get(dia, str(dia))
    mes_txt = MESES_EXTENSO.get(mes, "março")
    
    anos_dict = {2024: "dois mil e vinte e quatro", 2025: "dois mil e vinte e cinco", 2026: "dois mil e vinte e seis", 2027: "dois mil e vinte e sete"}
    ano_txt = anos_dict.get(ano, str(ano))
    
    horas_txt = NUMEROS_EXTENSO.get(hora, str(hora))
    minutos_txt = "" if minuto == 0 else f" e {NUMEROS_EXTENSO.get(minuto, str(minuto))}"
    
    data_extenso = f"{dia_txt} do mês de {mes_txt} do ano {ano_txt}"
    hora_extenso = f"{horas_txt} horas{minutos_txt}"
    data_curta = f"{dia} de {mes_txt} de {ano}"
    
    return data_extenso, hora_extenso, data_curta

def processar_documentos():
    global arquivos_selecionados

    texto_data_hora = entry_data.get().strip()
    data_extenso, hora_extenso, data_curta = converter_entrada_livre_extenso(texto_data_hora)
    data = texto_data_hora.split(',')[0].strip() if ',' in texto_data_hora else texto_data_hora.split()[0].strip()

    link = entry_link.get().strip()
    plataforma = "plataforma online"
    if "meet.google" in link.lower():
        plataforma = "Google Meet"
    elif "teams" in link.lower():
        plataforma = "Microsoft Teams"
    elif "zoom" in link.lower():
        plataforma = "Zoom"

    dados = {
        'nome_discente': entry_nome.get().strip(),
        'data_hora': texto_data_hora,
        'data_extenso': data_extenso,      
        'hora_extenso': hora_extenso,      
        'data_curta': data_curta,          
        'plataforma_remota': plataforma,   
        'data': data,
        'local_defesa': entry_local.get().strip() if chk_var_remoto.get() == 0 else "",
        'titulo_trabalho': entry_titulo.get().strip(),
        'especializacao': entry_especializacao.get().strip(),
        'area_trabalho': entry_area.get().strip(),
        'nome_orientador': entry_orientador.get().strip(),
        'nome_coorientador': entry_coorientador.get().strip() if chk_var_coorientador.get() == 1 else "",
        'examinadores_internos': entry_internos.get().strip(),
        'examinadores_externos': entry_externos.get().strip(),
        'numero_defesa': entry_numero.get().strip(),
        'resumo_trabalho': entry_resumo.get("1.0", tk.END).strip(),
        'formato_remoto': bool(chk_var_remoto.get()),
        'link_defesa': link
    }

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
            {"template": "doc_ata.docx", "prefixo_saida": "Ata_de_Defesa", "gerar": True},
            {"template": "doc_coorientador.docx", "prefixo_saida": "Declaração_Coorientador",
                                                                                 "gerar": bool(dados['nome_coorientador'])},
                                
        ]
        
        for t in templates_docx:
            if not t["gerar"]:
                continue

            caminho_docx_template = os.path.join(DIR_ATUAL, "Templates", t["template"])
            
            if os.path.exists(caminho_docx_template):
                doc = DocxTemplate(caminho_docx_template)
                doc.render(dados)
                
                nome_saida_docx = f"{t['prefixo_saida']}_{identificador}.docx"
                if t["template"] == "doc_ata.docx":
                    caminho_docx_saida = os.path.join(pasta_liberacao, nome_saida_docx)
                else:
                    caminho_docx_saida = os.path.join(pasta_comprovacao, nome_saida_docx)
                    
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

        if arquivos_selecionados:
            for caminho in arquivos_selecionados:
                nome_arquivo = os.path.basename(caminho)
                shutil.copy2(caminho, os.path.join(pasta_liberacao, nome_arquivo))

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
    global arquivos_selecionados

    caminhos = filedialog.askopenfilenames(filetypes=[("Todos os Arquivos", "*.*")])

    if caminhos:
        arquivos_selecionados.extend(caminhos)
        arquivos_selecionados = list(set(arquivos_selecionados))

        qtd = len(arquivos_selecionados)

        lbl_caminho_anexo.config(text=f"{qtd} arquivo(s) selecionado(s)")

def limpar_formulario():
    global arquivos_selecionados

    entry_coorientador.config(state="normal")
    entry_local.config(state="normal")
    entry_link.config(state="normal")

    for entry in entradas_simples:
        entry.delete(0, tk.END)
    
    entry_resumo.delete("1.0", tk.END)

    chk_var_coorientador.set(0)
    chk_var_remoto.set(0)
    toggle_coorientador()

    arquivos_selecionados = []
    lbl_caminho_anexo.config(text="Nenhum arquivo selecionado")

def toggle_coorientador():
    if chk_var_coorientador.get() == 1:
        entry_coorientador.config(state="normal")

    else:
        entry_coorientador.config(state="normal")
        entry_coorientador.delete(0, tk.END)
        entry_coorientador.config(state="disabled")

def toggle_link_remoto():
    if chk_var_remoto.get() == 1:
        entry_link.config(state="normal")
        entry_local.delete(0, tk.END)
        entry_local.config(state="disabled")

    else:
        entry_link.config(state="normal")
        entry_link.delete(0, tk.END)
        entry_link.config(state="disabled")

        entry_local.config(state="normal")


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
    ("Local Defesa (Presencial):", "entry_local"),
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

frame_formato = tk.Frame(main_frame)
frame_formato.pack(fill="x", pady=(2, 4))

chk_var_remoto = tk.IntVar(value=0)
chk_remoto = tk.Checkbutton(
    frame_formato, 
    text="Defesa em Formato Remoto (Online)",
    command=toggle_link_remoto, 
    variable=chk_var_remoto,
    font=font_label
)

chk_remoto.pack(anchor="w", padx=5)

tk.Label(frame_formato, text="Link da Defesa (Teams, Meet, Zoom):", font=("Arial", 9, "italic")).pack(anchor="w", padx=5)
entry_link = tk.Entry(frame_formato, width=100, state="disabled")
entry_link.pack(anchor="w", padx=5)

tk.Label(main_frame, text="Resumo:", font=font_label).pack(anchor="w", pady=(4, 0))
entry_resumo = tk.Text(main_frame, width=75, height=4, font=("Arial", 10))
entry_resumo.pack(anchor="w", padx=5, pady=(2, 5))

frame_rodape = tk.Frame(root)
frame_rodape.pack(fill="x", padx=20, pady=(5, 15))

tk.Label(frame_rodape, text="Anexos de Material (Opcional):", font=font_label).pack(anchor="w")

btn_anexo = tk.Button(frame_rodape, text="Procurar...", command=selecionar_arquivo, width=15)
btn_anexo.pack(side="left", pady=5)

lbl_caminho_anexo = tk.Label(frame_rodape, text="Nenhum arquivo selecionado", fg="gray")
lbl_caminho_anexo.pack(side="left", padx=15, pady=5)

btn_gerar = tk.Button(
    root, 
    text="Gerar Textos dos Documentos", 
    font=("Arial", 12, "bold"), 
    bg="#2a82da", 
    fg="white",
    command=processar_documentos
)
btn_gerar.pack(fill="x", padx=20, pady=(0, 10))

root.mainloop()