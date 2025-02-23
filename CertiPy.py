import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import os
import sys
from pathlib import Path
import ctypes
import threading

icone = os.path.join(os.path.dirname(__file__), 'CertiPy.ico')

class CertificadoInstalador(tk.Tk):
    def __init__(self):
        super().__init__()

        bg_color = '#f0f0f0'
        self.title("CertiPy")
        self.geometry("600x400")
        self.iconbitmap(icone)
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure(bg=bg_color)
        self.style.configure('TButton', padding=5)
        self.style.configure('TFrame', background=bg_color)
        self.style.configure('TLabel', background=bg_color)
        self.style.configure('TLabelframe', background=bg_color)
        self.style.configure('TLabelframe.Label', background=bg_color)
        self.create_widgets()
        
        if not self.is_admin():
            self.reiniciar_como_admin()

    def create_widgets(self):
        main_frame = ttk.Frame(self)
        main_frame.pack(expand=True, fill='both', padx=20, pady=20)

        title_label = ttk.Label(
            main_frame,
            text="Gerenciador de Certificados Digitais",
            font=('Helvetica', 16, 'bold'),
            background='#f0f0f0'
        )
        title_label.pack(pady=10)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(expand=True, fill='both', pady=20)

        self.select_btn = ttk.Button(
            button_frame,
            text="Selecionar Certificado",
            command=self.selecionar_certificado
        )
        self.select_btn.pack(pady=5, fill='x')

        self.install_btn = ttk.Button(
            button_frame,
            text="Instalar Certificado",
            command=self.instalar_certificado_selecionado,
            state='disabled'
        )
        self.install_btn.pack(pady=5, fill='x')

        self.list_btn = ttk.Button(
            button_frame,
            text="Listar Certificados Instalados",
            command=self.listar_certificados
        )
        self.list_btn.pack(pady=5, fill='x')

        self.uninstall_btn = ttk.Button(
            button_frame,
            text="Desinstalar Certificado",
            command=self.desinstalar_certificado
        )
        self.uninstall_btn.pack(pady=5, fill='x')

        log_frame = ttk.LabelFrame(main_frame, text="Log de Operações")
        log_frame.pack(expand=True, fill='both', pady=10)

        self.log_text = tk.Text(
            log_frame,
            height=8,
            wrap=tk.WORD,
            font=('Consolas', 9),
            background='white',
            relief='flat'
        )
        self.log_text.pack(expand=True, fill='both', padx=5, pady=5)

        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

        self.status_var = tk.StringVar()
        self.status_var.set("Pronto")
        status_bar = ttk.Label(
            self,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.certificado_selecionado = None

    def is_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def reiniciar_como_admin(self):
        if not self.is_admin():
            messagebox.showwarning(
                "Privilégios Necessários",
                "Este programa precisa ser executado como administrador."
            )
            try:
                ctypes.windll.shell32.ShellExecuteW(
                    None,
                    "runas",
                    sys.executable,
                    " ".join(sys.argv),
                    None,
                    1
                )
                sys.exit()
            except:
                messagebox.showerror(
                    "Erro",
                    "Falha ao obter privilégios de administrador."
                )
                sys.exit(1)

    def log(self, mensagem):
        self.log_text.insert(tk.END, f"{mensagem}\n")
        self.log_text.see(tk.END)

    def selecionar_certificado(self):
        filetypes = [
            ('Certificados', '*.pfx;*.p12;*.cer;*.crt;*.der;*.pem'),
            ('Todos os arquivos', '*.*')
        ]
        
        arquivo = filedialog.askopenfilename(
            title="Selecione o certificado",
            filetypes=filetypes
        )
        
        if arquivo:
            self.certificado_selecionado = arquivo
            self.install_btn.config(state='normal')
            self.log(f"Certificado selecionado: {arquivo}")
            self.status_var.set(f"Certificado selecionado: {os.path.basename(arquivo)}")

    def instalar_certificado_selecionado(self):
        if not self.certificado_selecionado:
            return

        ext = Path(self.certificado_selecionado).suffix.lower()
        
        if ext in ['.pfx', '.p12']:

            dialog = tk.Toplevel(self)
            dialog.title("Senha do Certificado")
            dialog.geometry("300x150")
            dialog.transient(self)
            dialog.grab_set()
            dialog.iconbitmap(icone)

            ttk.Label(
                dialog,
                text="Digite a senha do certificado:",
                padding=10
            ).pack()

            senha_var = tk.StringVar()
            senha_entry = ttk.Entry(dialog, show="*", textvariable=senha_var)
            senha_entry.pack(padx=20, pady=10)

            def confirmar():
                senha = senha_var.get()
                dialog.destroy()
                self.executar_instalacao(self.certificado_selecionado, senha)

            ttk.Button(
                dialog,
                text="Confirmar",
                command=confirmar
            ).pack(pady=10)

        else:
            self.executar_instalacao(self.certificado_selecionado)

    def executar_instalacao(self, caminho_certificado, senha=None):
        def instalar():
            try:
                if senha:
                    comando = [
                        "certutil",
                        "-importPFX",
                        "-p", senha,
                        caminho_certificado
                    ]
                else:
                    comando = [
                        "certutil",
                        "-addstore",
                        "My",
                        caminho_certificado
                    ]

                processo = subprocess.run(
                    comando,
                    check=True,
                    capture_output=True,
                    text=True
                )

                self.log(f"Saída do comando:\n{processo.stdout}")
                
                if processo.returncode == 0:
                    messagebox.showinfo(
                        "Sucesso",
                        "Certificado instalado com sucesso!"
                    )
                    self.status_var.set("Certificado instalado com sucesso")
                else:
                    messagebox.showerror(
                        "Erro",
                        f"Erro ao instalar o certificado. Código: {processo.returncode}"
                    )
                    
            except subprocess.CalledProcessError as e:
                messagebox.showerror(
                    "Erro",
                    f"Erro ao instalar o certificado:\n{e.stderr}"
                )
                self.log(f"Erro: {e.stderr}")

        thread = threading.Thread(target=instalar)
        thread.start()

    def listar_certificados(self):
        def listar():
            try:
                processo = subprocess.run(
                    ["certutil", "-store", "My"],
                    check=True,
                    capture_output=True,
                    text=True
                )
                
                linhas = processo.stdout.split('\n')
                certificados = []
                cert_atual = {}
                
                for linha in linhas:
                    if "Número de Série:" in linha:
                        if cert_atual:
                            certificados.append(cert_atual)
                        cert_atual = {'serial': linha.split(': ')[1].strip()}
                    elif "Requerente:" in linha:
                        cert_atual['nome'] = linha.split(': ')[1].strip()
                    elif "NotAfter:" in linha:
                        cert_atual['validade'] = linha.split(': ')[1].strip()
                
                if cert_atual:
                    certificados.append(cert_atual)
                
                self.log_text.delete(1.0, tk.END)
                
                self.log("Certificados instalados:")
                for i, cert in enumerate(certificados, 1):
                    self.log(f"\nCertificado {i}:")
                    self.log(f"Nome: {cert.get('nome', 'N/A')}")
                    self.log(f"Número de Série: {cert.get('serial', 'N/A')}")
                    self.log(f"Validade: {cert.get('validade', 'N/A')}")
                
                self.status_var.set("Certificados listados com sucesso")
                
            except subprocess.CalledProcessError as e:
                messagebox.showerror(
                    "Erro",
                    f"Erro ao listar certificados:\n{e.stderr}"
                )
                self.log(f"Erro: {e.stderr}")

        thread = threading.Thread(target=listar)
        thread.start()

    def desinstalar_certificado(self):
        def selecionar_certificado():
            try:
                processo = subprocess.run(
                    ["certutil", "-store", "My"],
                    check=True,
                    capture_output=True,
                    text=True,
                    encoding='latin1'
                )
                
                if not processo.stdout:
                    messagebox.showerror("Erro", "Não foi possível listar os certificados")
                    return
                
                dialog = tk.Toplevel(self)
                dialog.title("Selecionar Certificado para Desinstalar")
                dialog.geometry("800x400")
                dialog.transient(self)
                dialog.grab_set()
                dialog.iconbitmap(icone)
                
                frame = ttk.Frame(dialog, padding="10")
                frame.pack(fill='both', expand=True)
                
                ttk.Label(frame, text="Selecione o certificado para desinstalar:").pack(pady=5)
                
                tree_frame = ttk.Frame(frame)
                tree_frame.pack(fill='both', expand=True, pady=5)
                
                tree = ttk.Treeview(tree_frame, columns=('serial', 'nome', 'validade'), show='headings')
                tree.heading('serial', text='Número de Série')
                tree.heading('nome', text='Nome')
                tree.heading('validade', text='Validade')
                
                tree.column('serial', width=150)
                tree.column('nome', width=400)
                tree.column('validade', width=150)
                
                vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
                hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
                tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
                
                tree.grid(column=0, row=0, sticky='nsew')
                vsb.grid(column=1, row=0, sticky='ns')
                hsb.grid(column=0, row=1, sticky='ew')
                
                tree_frame.grid_columnconfigure(0, weight=1)
                tree_frame.grid_rowconfigure(0, weight=1)
                
                try:
                    linhas = processo.stdout.split('\n')
                    cert_atual = None
                    
                    for linha in linhas:
                        linha = linha.strip()
                        if "================ Certificado" in linha:
                            if cert_atual:
                                tree.insert('', 'end', values=(
                                    cert_atual.get('serial', ''),
                                    cert_atual.get('nome', ''),
                                    cert_atual.get('validade', '')
                                ))
                            cert_atual = {}
                        elif cert_atual is not None:
                            if "Número de Série:" in linha:
                                cert_atual['serial'] = linha.split(': ')[1].strip()
                            elif "Requerente:" in linha:
                                cert_atual['nome'] = linha.split(': ')[1].strip()
                            elif "NotAfter:" in linha:
                                cert_atual['validade'] = linha.split(': ')[1].strip()
                    
                    if cert_atual:
                        tree.insert('', 'end', values=(
                            cert_atual.get('serial', ''),
                            cert_atual.get('nome', ''),
                            cert_atual.get('validade', '')
                        ))
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao processar certificados: {str(e)}")
                    return
                
                def confirmar_desinstalacao():
                    selection = tree.selection()
                    if not selection:
                        messagebox.showwarning("Aviso", "Selecione um certificado para desinstalar")
                        return
                    
                    item = tree.item(selection[0])
                    serial = item['values'][0]
                    nome = item['values'][1]
                    
                    if messagebox.askyesno("Confirmar", 
                        f"Tem certeza que deseja desinstalar o certificado?\n\n"
                        f"Nome: {nome}\n"
                        f"Número de Série: {serial}"):
                        dialog.destroy()
                        self.executar_desinstalacao(serial)
                
                button_frame = ttk.Frame(frame)
                button_frame.pack(pady=10, fill='x')
                
                ttk.Button(
                    button_frame,
                    text="Desinstalar",
                    command=confirmar_desinstalacao
                ).pack(side='left', padx=5)
                
                ttk.Button(
                    button_frame,
                    text="Cancelar",
                    command=dialog.destroy
                ).pack(side='right', padx=5)
                
            except subprocess.CalledProcessError as e:
                messagebox.showerror("Erro", f"Erro ao listar certificados: {str(e)}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro inesperado: {str(e)}")
        
        thread = threading.Thread(target=selecionar_certificado)
        thread.start()

    def executar_desinstalacao(self, serial):
        def desinstalar():
            try:
                comando = ["certutil", "-delstore", "My", serial]
                processo = subprocess.run(
                    comando,
                    check=True,
                    capture_output=True,
                    text=True
                )
                
                if processo.returncode == 0:
                    messagebox.showinfo("Sucesso", "Certificado desinstalado com sucesso!")
                    self.log(f"Certificado {serial} desinstalado com sucesso")
                    self.status_var.set("Certificado desinstalado com sucesso")
                else:
                    messagebox.showerror("Erro", "Erro ao desinstalar o certificado")
                    
            except subprocess.CalledProcessError as e:
                messagebox.showerror("Erro", f"Erro ao desinstalar o certificado:\n{e.stderr}")
                self.log(f"Erro: {e.stderr}")
        
        thread = threading.Thread(target=desinstalar)
        thread.start()

if __name__ == "__main__":
    app = CertificadoInstalador()
    app.mainloop()