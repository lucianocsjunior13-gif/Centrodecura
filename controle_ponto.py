import sqlite3
from datetime import datetime
import pandas as pd  # Biblioteca essencial para gerar o Excel
import os
import barcode
from barcode.writer import ImageWriter

def inicializar_banco():
    # Cria (ou conecta) ao arquivo local na máquina
    conn = sqlite3.connect('controle_ponto.db')
    cursor = conn.cursor()
    
    # Tabela de quem tem crachá
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS funcionarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            codigo_barra TEXT UNIQUE NOT NULL
        )
    ''')
    
    # Tabela de quando o crachá foi lido
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

# Função que será chamada quando o leitor apitar
def registrar_leitura(codigo_lido):
    conn = sqlite3.connect('controle_ponto.db')
    cursor = conn.cursor()
    
    # Busca se o código existe
    cursor.execute('SELECT id, nome FROM funcionarios WHERE codigo_barra = ?', (codigo_lido,))
    funcionario = cursor.fetchone()
    
    if funcionario:
        agora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        cursor.execute('INSERT INTO registros (funcionario_id, data_hora) VALUES (?, ?)', (funcionario[0], agora))
        conn.commit()
        print(f"Sucesso: {funcionario[1]} registrado às {agora} desu~ ✨")
    else:
        print("Erro: Código de barras não encontrado! Verifique se o funcionário está cadastrado. 😢")
        
    conn.close()

# NOVA FUNÇÃO: Cadastrar funcionário (use antes de ler crachás)
def cadastrar_funcionario(nome, codigo_barra):
    conn = sqlite3.connect('controle_ponto.db')
    cursor = conn.cursor()
    try:
        cursor.execute('INSERT INTO funcionarios (nome, codigo_barra) VALUES (?, ?)', (nome, codigo_barra))
        conn.commit()
        print(f"Funcionário {nome} cadastrado com sucesso! Código: {codigo_barra} desu~ 💖")
    except sqlite3.IntegrityError:
        print(f"Erro: Código de barras {codigo_barra} já está em uso! Tente outro. 😅")
    finally:
        conn.close()

# Função de exportar para Excel (versão solicitada)
def exportar_excel():
    conn = sqlite3.connect('controle_ponto.db')
    
    # Faz um JOIN para pegar o nome do funcionário e a hora do registro
    query = '''
        SELECT f.nome as Nome, f.codigo_barra as Codigo, r.data_hora as Data_Hora
        FROM registros r
        JOIN funcionarios f ON r.funcionario_id = f.id
    '''
    
    df = pd.read_sql_query(query, conn)
    df.to_excel('relatorio_leituras.xlsx', index=False)
    print("Planilha relatorio_leituras.xlsx gerada com sucesso!")
    
    conn.close()

# NOVA FUNÇÃO: Ver registros no terminal (pra checagem rápida)
def listar_registros():
    conn = sqlite3.connect('controle_ponto.db')
    cursor = conn.cursor()
    cursor.execute('''
        SELECT f.nome, r.data_hora 
        FROM registros r
        JOIN funcionarios f ON r.funcionario_id = f.id
        ORDER BY r.data_hora DESC
        LIMIT 10
    ''')
    registros = cursor.fetchall()
    conn.close()
    
    if not registros:
        print("Nenhum registro ainda. 😴")
        return
        
    print("\nÚltimos 10 registros:")
    for nome, data_hora in registros:
        print(f"  {nome} - {data_hora}")
    print()

# Função para calcular dígito verificador EAN
def calcular_digito_verificador(base):
    """base: string of digits (length 7 for EAN-8 or 12 for EAN-13)"""
    # Ensure it's digits only
    if not base.isdigit():
        raise ValueError("Base must contain only digits")
    # Convert to list of ints
    digits = [int(d) for d in base]
    # EAN checksum: sum of odd positions *3 + sum of even positions, from left, position 1
    total = 0
    for i, d in enumerate(digits):
        # position i+1
        if (i+1) % 2 == 1:  # odd position
            total += d * 3
        else:
            total += d
    checksum = (10 - (total % 10)) % 10
    return str(checksum)

# Função para gerar e salvar imagem de código de barras
def gerar_imagem_barcode(codigo_completo, nome):
    """Salva PNG em pasta crachas/ e retorna caminho completo."""
    # Ensure folder exists
    folder = "crachas"
    if not os.path.exists(folder):
        os.makedirs(folder)
    # Determine barcode type
    if len(codigo_completo) == 8:
        barcode_class = barcode.get_barcode_class('ean8')
    elif len(codigo_completo) == 13:
        barcode_class = barcode.get_barcode_class('ean13')
    else:
        raise ValueError("Código de barras deve ter 8 ou 13 dígitos")
    # Create barcode instance
    bar = barcode_class(codigo_completo, writer=ImageWriter())
    # Sanitize filename: remove invalid chars
    safe_nome = "".join(c for c in nome if c.isalnum() or c in (' ', '_', '-')).rstrip()
    filename = f"{safe_nome}_{codigo_completo}"
    full_path = os.path.join(folder, filename)
    # Save (returns full path with extension)
    saved_path = bar.save(full_path)
    return saved_path

if __name__ == "__main__":
    try:
        import customtkinter as ctk
        from tkinter import messagebox
    except ImportError:
        print("CustomTkinter não está instalado. Instale com: pip install customtkinter")
        exit(1)

    def main_gui():
        inicializar_banco()
        
        app = ctk.CTk()
        app.title("Controle de Ponto")
        app.geometry("400x250")
        app.resizable(False, False)
        
        # Tabview
        tabview = ctk.CTkTabview(app, width=380, height=200)
        tabview.pack(padx=10, pady=10, fill="both", expand=True)
        tab_leitura = tabview.add("Leitura")
        tab_cadastro = tabview.add("Cadastrar Funcionário")
        
        # --- Leitura Tab ---
        # Hidden entry for barcode input
        entry_hidden = ctk.CTkEntry(tab_leitura, width=1, height=1)
        entry_hidden.place(x=-100, y=-100)  # Hide outside window
        entry_hidden.focus_set()
        
        def on_enter(event):
            codigo = entry_hidden.get().strip()
            if codigo:  # only process if not empty
                entry_hidden.delete(0, 'end')
                registrar_leitura(codigo)
            return "break"  # prevent further propagation
        
        entry_hidden.bind('<Return>', on_enter)
        
        # Generate Report button
        def gerar_relatorio():
            exportar_excel()
            messagebox.showinfo("Sucesso", "Relatório gerado com sucesso! desu~ 💖")
        
        btn_relatorio = ctk.CTkButton(tab_leitura, text="Gerar Relatório", command=gerar_relatorio)
        btn_relatorio.place(relx=0.5, rely=0.5, anchor='center')
        
        # --- Cadastro Tab ---
        label_nome = ctk.CTkLabel(tab_cadastro, text="Nome do Funcionário:")
        label_nome.grid(row=0, column=0, padx=10, pady=(10,5), sticky="w")
        entry_nome = ctk.CTkEntry(tab_cadastro, width=200)
        entry_nome.grid(row=0, column=1, padx=10, pady=(10,5), sticky="w")
        
        label_base = ctk.CTkLabel(tab_cadastro, text="Número base (7 ou 12 dígitos):")
        label_base.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        entry_base = ctk.CTkEntry(tab_cadastro, width=200)
        entry_base.grid(row=1, column=1, padx=10, pady=5, sticky="w")
        
        def cadastrar_e_gerar():
            nome = entry_nome.get().strip()
            base = entry_base.get().strip()
            if not nome:
                messagebox.showwarning("Erro", "Por favor, insira o nome do funcionário.")
                return
            if not base.isdigit():
                messagebox.showwarning("Erro", "O número base deve conter apenas dígitos.")
                return
            if len(base) not in (7, 12):
                messagebox.showwarning("Erro", "O número base deve ter 7 dígitos (EAN-8) ou 12 dígitos (EAN-13).")
                return
            try:
                dv = calcular_digito_verificador(base)
                codigo_completo = base + dv
                # Check if already exists
                conn = sqlite3.connect('controle_ponto.db')
                cursor = conn.cursor()
                cursor.execute('SELECT id FROM funcionarios WHERE codigo_barra = ?', (codigo_completo,))
                if cursor.fetchone():
                    conn.close()
                    messagebox.showwarning("Erro", f"O código de barras {codigo_completo} já está cadastrado.")
                    return
                conn.close()
                # Generate barcode image
                img_path = gerar_imagem_barcode(codigo_completo, nome)
                # Register in DB
                cadastrar_funcionario(nome, codigo_completo)
                messagebox.showinfo("Sucesso", 
                                    f"Funcionário {nome} cadastrado!\n"
                                    f"Código completo: {codigo_completo}\n"
                                    f"Imagem salva em: {img_path}\ndesu~ 💖")
                # Clear fields
                entry_nome.delete(0, 'end')
                entry_base.delete(0, 'end')
            except Exception as e:
                messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
        
        btn_cadastrar = ctk.CTkButton(tab_cadastro, text="Gerar Código de Barras e Cadastrar", command=cadastrar_e_gerar)
        btn_cadastrar.grid(row=2, column=0, columnspan=2, pady=20)
        
        app.mainloop()

    main_gui()