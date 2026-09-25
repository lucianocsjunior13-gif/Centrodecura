import sqlite3
from datetime import datetime
import pandas as pd
import os
import barcode
from barcode.writer import ImageWriter
import customtkinter as ctk
from tkinter import messagebox

def inicializar_banco():
    conn = sqlite3.connect('controle_ponto.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS funcionarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            codigo_barra TEXT UNIQUE NOT NULL
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS registros (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            funcionario_id INTEGER,
            data_hora DATETIME NOT NULL,
            FOREIGN KEY (funcionario_id) REFERENCES funcionarios(id)
        )
    ''')
    conn.commit()
    conn.close()

def registrar_leitura(codigo_lido):
    conn = sqlite3.connect('controle_ponto.db')
    cursor = conn.cursor()
    cursor.execute('SELECT id, nome FROM funcionarios WHERE codigo_barra = ?', (codigo_lido,))
    funcionario = cursor.fetchone()
    
    if funcionario:
        agora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        cursor.execute('INSERT INTO registros (funcionario_id, data_hora) VALUES (?, ?)', (funcionario[0], agora))
        conn.commit()
        conn.close()
        return True, f"✅ {funcionario[1]}\nRegistrado às {agora}"
    else:
        conn.close()
        return False, f"❌ Código ({codigo_lido}) não cadastrado!"

def cadastrar_funcionario(nome, codigo_barra):
    conn = sqlite3.connect('controle_ponto.db')
    cursor = conn.cursor()
    cursor.execute('INSERT INTO funcionarios (nome, codigo_barra) VALUES (?, ?)', (nome, codigo_barra))
    conn.commit()
    conn.close()

def exportar_excel():
    conn = sqlite3.connect('controle_ponto.db')
    query = '''
        SELECT f.nome as Nome, f.codigo_barra as Codigo, r.data_hora as Data_Hora
        FROM registros r
        JOIN funcionarios f ON r.funcionario_id = f.id
        ORDER BY r.id DESC
    '''
    df = pd.read_sql_query(query, conn)
    df.to_excel('relatorio_leituras.xlsx', index=False)
    conn.close()

def gerar_imagem_barcode(base_codigo, nome):
    folder = "crachas"
    if not os.path.exists(folder):
        os.makedirs(folder)
    
    tipo = 'ean8' if len(base_codigo) == 7 else 'ean13'
    barcode_class = barcode.get_barcode_class(tipo)
    
    # Calcula o dígito verificador oficial automaticamente
    bar = barcode_class(base_codigo, writer=ImageWriter())
    codigo_completo = bar.get_fullcode()
    
    safe_nome = "".join(c for c in nome if c.isalnum() or c in (' ', '_', '-')).rstrip()
    full_path = os.path.join(folder, f"{safe_nome}_{codigo_completo}")
    saved_path = bar.save(full_path)
    return saved_path, codigo_completo

if __name__ == "__main__":
    inicializar_banco()
    
    app = ctk.CTk()
    app.title("Controle de Ponto - Central")
    
    # Abre o app maximizado (tela inteira ocupando o monitor)
    app.state("zoomed")
    
    # Cores personalizadas
    COR_BG = "#1F3320"       # Fundo principal
    COR_CARD = "#215223"     # Fundo dos blocos / Botões padrão
    COR_HOVER = "#1B8520"    # Cor ao passar o mouse por cima do botão
    COR_TEXTO = "#FDFBD4"    # Texto creme claro
    
    app.configure(fg_color=COR_BG)
    
    # Layout centralizado preenchendo a tela
    container = ctk.CTkFrame(app, fg_color=COR_BG)
    container.pack(fill="both", expand=True, padx=40, pady=40)
    
    # Título Principal do App
    lbl_titulo = ctk.CTkLabel(container, text="SISTEMA DE CONTROLE DE PONTO", font=("Arial", 28, "bold"), text_color=COR_TEXTO)
    lbl_titulo.pack(pady=(10, 20))
    
    # Abas
    tabview = ctk.CTkTabview(container, fg_color=COR_CARD, segmented_button_fg_color=COR_BG, 
                             segmented_button_selected_color=COR_HOVER, segmented_button_selected_hover_color=COR_HOVER,
                             text_color=COR_TEXTO)
    tabview.pack(fill="both", expand=True, padx=20, pady=10)
    
    tab_leitura = tabview.add("Leitura de Crachá")
    tab_cadastro = tabview.add("Cadastrar Funcionário")
    
    # --- ABA DE LEITURA ---
    frame_leitura_centro = ctk.CTkFrame(tab_leitura, fg_color="transparent")
    frame_leitura_centro.place(relx=0.5, rely=0.4, anchor="center")
    
    lbl_instrucao = ctk.CTkLabel(frame_leitura_centro, text="Passe o crachá no leitor USB (ou digite e aperte Enter):", font=("Arial", 18), text_color=COR_TEXTO)
    lbl_instrucao.pack(pady=10)

    entry_leitura = ctk.CTkEntry(frame_leitura_centro, width=380, height=50, justify="center", font=("Arial", 20), placeholder_text="Aguardando crachá...", text_color=COR_TEXTO)
    entry_leitura.pack(pady=10)
    entry_leitura.focus_set()

    lbl_feedback = ctk.CTkLabel(frame_leitura_centro, text="", font=("Arial", 22, "bold"), text_color=COR_TEXTO)
    lbl_feedback.pack(pady=20)

    def on_enter(event):
        codigo = entry_leitura.get().strip()
        if codigo:
            entry_leitura.delete(0, 'end')
            sucesso, mensagem = registrar_leitura(codigo)
            cor_msg = "#2ecc71" if sucesso else "#e74c3c"
            lbl_feedback.configure(text=mensagem, text_color=cor_msg)
        return "break"
    
    entry_leitura.bind('<Return>', on_enter)

    def gerar_relatorio():
        exportar_excel()
        messagebox.showinfo("Sucesso", "Planilha 'relatorio_leituras.xlsx' gerada na pasta do programa!")
        entry_leitura.focus_set()
    
    btn_relatorio = ctk.CTkButton(container, text="📊 Exportar Relatório em Excel", command=gerar_relatorio, 
                                  fg_color=COR_HOVER, hover_color="#146617", text_color=COR_TEXTO, font=("Arial", 16, "bold"), width=300, height=45)
    btn_relatorio.pack(pady=20)

    def ao_trocar_aba():
        if tabview.get() == "Leitura de Crachá":
            entry_leitura.focus_set()
    tabview.configure(command=ao_trocar_aba)
    
    # --- ABA DE CADASTRO ---
    frame_cad_centro = ctk.CTkFrame(tab_cadastro, fg_color="transparent")
    frame_cad_centro.place(relx=0.5, rely=0.4, anchor="center")
    
    label_nome = ctk.CTkLabel(frame_cad_centro, text="Nome do Funcionário:", font=("Arial", 16), text_color=COR_TEXTO)
    label_nome.grid(row=0, column=0, padx=15, pady=15, sticky="w")
    entry_nome = ctk.CTkEntry(frame_cad_centro, width=300, height=35, font=("Arial", 14), placeholder_text="Ex: Ana Souza", text_color=COR_TEXTO)
    entry_nome.grid(row=0, column=1, padx=15, pady=15, sticky="w")
    
    label_base = ctk.CTkLabel(frame_cad_centro, text="Número base (7 ou 12 dígitos):", font=("Arial", 16), text_color=COR_TEXTO)
    label_base.grid(row=1, column=0, padx=15, pady=15, sticky="w")
    entry_base = ctk.CTkEntry(frame_cad_centro, width=300, height=35, font=("Arial", 14), placeholder_text="Ex: 7891011", text_color=COR_TEXTO)
    entry_base.grid(row=1, column=1, padx=15, pady=15, sticky="w")
    
    def cadastrar_e_gerar():
        nome = entry_nome.get().strip()
        base = entry_base.get().strip()
        if not nome or not base.isdigit() or len(base) not in (7, 12):
            messagebox.showwarning("Atenção", "Preencha o nome e digite exatamente 7 dígitos (EAN-8) ou 12 dígitos (EAN-13).")
            return
        try:
            img_path, codigo_completo = gerar_imagem_barcode(base, nome)
            cadastrar_funcionario(nome, codigo_completo)
            messagebox.showinfo("Sucesso", f"Funcionário cadastrado!\n\nCódigo final: {codigo_completo}\nCrachá salvo em: {img_path}")
            entry_nome.delete(0, 'end')
            entry_base.delete(0, 'end')
        except sqlite3.IntegrityError:
            messagebox.showwarning("Erro", "Esse código de barras já está cadastrado para outra pessoa!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
    
    btn_cadastrar = ctk.CTkButton(frame_cad_centro, text="Gerar Crachá e Cadastrar", command=cadastrar_e_gerar, 
                                  fg_color=COR_HOVER, hover_color="#146617", text_color=COR_TEXTO, font=("Arial", 16, "bold"), width=280, height=45)
    btn_cadastrar.grid(row=2, column=0, columnspan=2, pady=30)
    
    app.mainloop()