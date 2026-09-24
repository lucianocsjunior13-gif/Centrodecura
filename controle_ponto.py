import sqlite3
from datetime import datetime
import pandas as pd  # Biblioteca essencial para gerar o Excel

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

# NOVA FUNÇÃO: Exportar para Excel (chame quando precisar exportar)
def exportar_excel(nome_arquivo='relatorio_ponto.xlsx'):
    conn = sqlite3.connect('controle_ponto.db')
    # Query junta funcionário + registro pra ficar legível
    query = '''
        SELECT f.nome AS Nome, f.codigo_barra AS 'Código de Barras', r.data_hora AS 'Data e Hora'
        FROM registros r
        JOIN funcionarios f ON r.funcionario_id = f.id
        ORDER BY r.data_hora DESC
    '''
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    if df.empty:
        print("Nenhum registro encontrado pra exportar. 😴")
        return
        
    df.to_excel(nome_arquivo, index=False)
    print(f"Relatório gerado com sucesso! Salvo como {nome_arquivo} desu~ 📊✨")

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
        app.geometry("250x100")
        app.resizable(False, False)
        
        # Hidden entry for barcode input
        entry = ctk.CTkEntry(app, width=1, height=1)
        entry.place(x=-100, y=-100)  # Hide outside window
        entry.focus_set()
        
        def on_enter(event):
            codigo = entry.get().strip()
            if codigo:  # only process if not empty
                entry.delete(0, 'end')
                registrar_leitura(codigo)
            return "break"  # prevent further propagation
        
        entry.bind('<Return>', on_enter)
        
        # Generate Report button
        def gerar_relatorio():
            exportar_excel()
            messagebox.showinfo("Sucesso", "Relatório gerado com sucesso! desu~ 💖")
        
        btn = ctk.CTkButton(app, text="Gerar Relatório", command=gerar_relatorio)
        btn.place(relx=0.5, rely=0.5, anchor='center')
        
        app.mainloop()

    main_gui()