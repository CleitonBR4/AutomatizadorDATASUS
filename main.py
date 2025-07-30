import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import re
from datetime import datetime
import logging
import threading
import queue
from ftplib import FTP, error_perm
import webbrowser
import subprocess
import requests

# --- Configurações Globais ---
FTP_SERVER = "arpoador.datasus.gov.br"
FTP_PATH_BPA = "/siasus/BPA/"
FTP_PATH_SIA = "/siasus/sia/"
FTP_PATH_FPO = "/siasus/fpo/"

# Diretórios padrão de instalação
DIR_BPA = "C:\\BPA"
DIR_FPO = "C:\\FPO"
DIR_SIA = "C:\\INSTSIA"
DIR_CNES = "C:\\CNES"

# Arquivos para controle de versão local
VERSION_FILE_BPA = os.path.join(DIR_BPA, "versao.txt")
VERSION_FILE_SIA = os.path.join(DIR_SIA, "versao.txt")
VERSION_FILE_FPO = os.path.join(DIR_FPO, "versao.txt")

# --- Configuração do Log de Atividades ---
logging.basicConfig(
    filename='log_automatizador_datasus.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    encoding='utf-8'
)

class App(tk.Tk):
    """Classe principal da aplicação com a interface gráfica."""
    def __init__(self):
        super().__init__()
        self.title("Automatizador de Faturamento DATASUS (BPA, SIA, FPO, CNES)")
        self.geometry("850x750")

        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.style.configure("TButton", padding=5, font=('Helvetica', 10))
        self.style.configure("TLabel", padding=2, font=('Helvetica', 10))
        self.style.configure("TLabelframe.Label", font=('Helvetica', 11, 'bold'))
        self.style.configure("Success.TButton", foreground="green", font=('Helvetica', 10, 'bold'))

        self.update_queue = queue.Queue()
        self.create_widgets()
        self.process_queue()
        self.log("Programa iniciado. Clique em 'Iniciar Verificação Geral' para começar.")

    def log(self, message, level="info"):
        """Registra uma mensagem na Central de Notificações e no arquivo de log."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, formatted_message + "\n")
        self.log_text.config(state=tk.DISABLED)
        self.log_text.see(tk.END)
        
        if level == "info": logging.info(message)
        elif level == "error": logging.error(message)
        elif level == "warning": logging.warning(message)

    def process_queue(self):
        """Verifica a fila de tarefas e atualiza a GUI de forma segura."""
        try:
            task = self.update_queue.get_nowait()
            task()
        except queue.Empty:
            pass
        finally:
            self.after(100, self.process_queue)

    def create_widgets(self):
        """Cria todos os componentes visuais da interface."""
        # --- Menu Superior ---
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        # Menu de Guias
        guides_menu = tk.Menu(menubar, tearoff=0)
        guides_menu.add_command(label="BPA", command=lambda: self.show_guide("bpa.pdf"))
        guides_menu.add_command(label="FPO", command=lambda: self.show_guide("fpo.pdf"))
        guides_menu.add_command(label="SIA", command=lambda: self.show_guide("sia.pdf"))
        menubar.add_cascade(label="Guias", menu=guides_menu)

        # Menu Firebird
        firebird_menu = tk.Menu(menubar, tearoff=0)
        firebird_menu.add_command(label="Verificar Versão", command=lambda: self.start_thread(self.check_firebird_version))
        firebird_menu.add_command(label="Download (1.5.5)", command=self.download_firebird)
        firebird_menu.add_separator()
        firebird_menu.add_command(label="Start (Iniciar Serviço)", command=lambda: self.start_thread(self.start_firebird_service))
        firebird_menu.add_command(label="Shutdown (Parar Serviço)", command=lambda: self.start_thread(self.stop_firebird_service))
        menubar.add_cascade(label="Firebird", menu=firebird_menu)

        # Menu de Avisos
        menubar.add_command(label="Avisos Importantes", command=self.show_warnings)

        # Menu de Links
        links_menu = tk.Menu(menubar, tearoff=0)
        links_menu.add_command(label="SIGTAP", command=lambda: self.open_link("http://sigtap.datasus.gov.br"))
        links_menu.add_command(label="CNES", command=lambda: self.open_link("http://cnes.datasus.gov.br"))
        links_menu.add_command(label="CNES Antigo", command=lambda: self.open_link("http://cnes2.datasus.gov.br"))
        menubar.add_cascade(label="Links Úteis", menu=links_menu)

        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Botão de Verificação Principal ---
        start_button_frame = ttk.Frame(main_frame)
        start_button_frame.pack(fill=tk.X, pady=10)
        self.start_button = ttk.Button(start_button_frame, text="▶ Iniciar Verificação Geral", command=self.initial_setup, style="Success.TButton")
        self.start_button.pack(pady=5)

        # --- Seção do Dashboard de Status ---
        dashboard_frame = ttk.LabelFrame(main_frame, text="Dashboard de Status", padding="10")
        dashboard_frame.pack(fill=tk.X, expand=False, pady=5)
        dashboard_frame.grid_columnconfigure(1, weight=1)

        # Status do BPA
        ttk.Label(dashboard_frame, text="BPA:", font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky="w", padx=5)
        self.bpa_status_var = tk.StringVar(value="Aguardando verificação...")
        self.bpa_status_display = ttk.Label(dashboard_frame, textvariable=self.bpa_status_var)
        self.bpa_status_display.grid(row=0, column=1, sticky="w")
        self.bpa_action_button = ttk.Button(dashboard_frame, text="Ação", state=tk.DISABLED, command=self.download_bpa)
        self.bpa_action_button.grid(row=0, column=2, padx=5, pady=2)
        ttk.Button(dashboard_frame, text="Abrir Pasta", command=lambda: self.open_directory(DIR_BPA)).grid(row=0, column=3, padx=5, pady=2)
        
        # Status do SIA
        ttk.Label(dashboard_frame, text="SIA:", font=("Helvetica", 10, "bold")).grid(row=1, column=0, sticky="w", padx=5)
        self.sia_status_var = tk.StringVar(value="Aguardando verificação...")
        self.sia_status_display = ttk.Label(dashboard_frame, textvariable=self.sia_status_var)
        self.sia_status_display.grid(row=1, column=1, sticky="w")
        self.sia_action_button = ttk.Button(dashboard_frame, text="Ação", state=tk.DISABLED, command=self.download_sia)
        self.sia_action_button.grid(row=1, column=2, padx=5, pady=2)
        ttk.Button(dashboard_frame, text="Abrir Pasta", command=lambda: self.open_directory(DIR_SIA)).grid(row=1, column=3, padx=5, pady=2)

        # Status do FPO
        ttk.Label(dashboard_frame, text="FPO:", font=("Helvetica", 10, "bold")).grid(row=2, column=0, sticky="w", padx=5)
        self.fpo_status_var = tk.StringVar(value="Aguardando verificação...")
        self.fpo_status_display = ttk.Label(dashboard_frame, textvariable=self.fpo_status_var)
        self.fpo_status_display.grid(row=2, column=1, sticky="w")
        self.fpo_action_button = ttk.Button(dashboard_frame, text="Ação", state=tk.DISABLED, command=self.download_fpo)
        self.fpo_action_button.grid(row=2, column=2, padx=5, pady=2)
        ttk.Button(dashboard_frame, text="Abrir Pasta", command=lambda: self.open_directory(DIR_FPO)).grid(row=2, column=3, padx=5, pady=2)

        # Status do CNES
        ttk.Label(dashboard_frame, text="CNES:", font=("Helvetica", 10, "bold")).grid(row=3, column=0, sticky="w", padx=5)
        self.cnes_status_var = tk.StringVar(value="Pronto para baixar")
        self.cnes_status_display = ttk.Label(dashboard_frame, textvariable=self.cnes_status_var)
        self.cnes_status_display.grid(row=3, column=1, sticky="w")
        self.cnes_action_button = ttk.Button(dashboard_frame, text="Baixar CNES", command=self.download_cnes)
        self.cnes_action_button.grid(row=3, column=2, padx=5, pady=2)
        ttk.Button(dashboard_frame, text="Abrir Pasta", command=lambda: self.open_directory(DIR_CNES)).grid(row=3, column=3, padx=5, pady=2)

        # --- Seção do BDSIA ---
        bdsia_frame = ttk.LabelFrame(main_frame, text="Download do BDSIA (Tabela Unificada)", padding="10")
        bdsia_frame.pack(fill=tk.X, expand=False, pady=10)
        bdsia_frame.grid_columnconfigure(0, weight=1)

        self.bdsia_labels = []
        self.bdsia_buttons = []
        for i in range(3):
            label = ttk.Label(bdsia_frame, text=f"Buscando versão {i+1}...")
            label.grid(row=i, column=0, sticky="w", padx=5, pady=2)
            button = ttk.Button(bdsia_frame, text="Baixar", state=tk.DISABLED)
            button.grid(row=i, column=1, padx=5, pady=2)
            self.bdsia_labels.append(label)
            self.bdsia_buttons.append(button)
        
        # --- Seção da Central de Notificações ---
        log_frame = ttk.LabelFrame(main_frame, text="Central de Notificações e Log de Atividades", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = tk.Text(log_frame, height=15, state=tk.DISABLED, wrap=tk.WORD, bg="#f0f0f0", font=('Courier New', 9))
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # --- Rodapé ---
        footer_text = "Este software é gratuito e de codigo aberto. Considere pagar um café para o desenvolvedor. PIX: cleiton.contato0@gmail.com"
        footer_label = ttk.Label(main_frame, text=footer_text, font=("Helvetica", 8), foreground="gray")
        footer_label.pack(side=tk.BOTTOM, pady=(5, 0))

    def start_thread(self, target_function, *args):
        """Cria e inicia uma nova thread para executar tarefas demoradas."""
        thread = threading.Thread(target=target_function, args=args)
        thread.daemon = True
        thread.start()

    def initial_setup(self):
        """Executa as verificações e configurações iniciais."""
        self.log("Iniciando verificação geral dos sistemas...")
        self.start_button.config(state=tk.DISABLED, text="Verificando...")
        self.start_thread(self.ensure_folders_exist)
        self.start_thread(self.check_bpa)
        self.start_thread(self.check_sia)
        self.start_thread(self.check_fpo)
        self.start_thread(self.check_bdsia)
        self.after(10000, lambda: self.start_button.config(state=tk.NORMAL, text="▶ Iniciar Verificação Geral"))

    def ensure_folders_exist(self):
        """Garante que os diretórios base e de exportação existam em C:\."""
        try:
            paths_to_create = [
                DIR_BPA, os.path.join(DIR_BPA, "EXPORTA"),
                DIR_FPO, os.path.join(DIR_FPO, "EXPORTA"),
                DIR_SIA, os.path.join(DIR_SIA, "IMPORTA"),
                DIR_CNES
            ]
            for path in paths_to_create:
                if not os.path.exists(path):
                    os.makedirs(path)
                    self.log(f"Diretório criado: {path}")
        except PermissionError:
            self.log("Erro de permissão ao criar pastas em C:\\. Execute como Administrador.", "error")
            self.update_queue.put(lambda: messagebox.showerror("Erro de Permissão", "Não foi possível criar as pastas necessárias em C:\\.\n\nPor favor, feche o programa e execute-o como Administrador."))
        except Exception as e:
            self.log(f"Erro inesperado ao criar diretórios: {e}", "error")

    def get_local_version(self, file_path):
        """Lê a versão salva localmente."""
        try:
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    return f.read().strip()
        except Exception as e:
            self.log(f"Erro ao ler arquivo de versão {file_path}: {e}", "error")
        return None

    def set_status(self, label_var, display_widget, message, color, button, button_text=""):
        """Atualiza a label de status e o botão de ação."""
        label_var.set(message)
        display_widget.config(foreground=color)
        if button_text:
            button.config(text=button_text, state=tk.NORMAL)
        else:
            button.config(state=tk.DISABLED)

    def list_ftp_files(self, ftp_path):
        """Conecta ao FTP e lista os arquivos em um diretório."""
        try:
            ftp = FTP(FTP_SERVER, timeout=20)
            ftp.login() # Login anônimo
            ftp.cwd(ftp_path)
            files = ftp.nlst()
            ftp.quit()
            return files
        except Exception as e:
            self.log(f"Falha ao conectar ou listar arquivos em {ftp_path}: {e}", "error")
            return None

    def open_directory(self, path):
        """Abre um diretório no explorador de arquivos."""
        try:
            if os.path.exists(path):
                os.startfile(path)
                self.log(f"Abrindo diretório: {path}")
            else:
                self.log(f"Tentativa de abrir diretório falhou. Caminho não existe: {path}", "warning")
                messagebox.showwarning("Diretório não encontrado", f"O diretório '{path}' ainda não existe.")
        except Exception as e:
            self.log(f"Erro ao abrir diretório {path}: {e}", "error")
            messagebox.showerror("Erro", f"Não foi possível abrir o diretório:\n{e}")

    def show_warnings(self):
        """Exibe uma caixa de diálogo com avisos importantes."""
        warnings = (
            "1. Importante: não extrair o arquivo BDSIA com o programa SIA aberto.\n\n"
            "2. Lembre-se de criar novas pastas da competência do mês.\n\n"
            "3. Sempre que for lançar um novo BPA consolidado, colocar a folha com número alto para evitar duplicidade."
        )
        messagebox.showinfo("Avisos Importantes", warnings)

    def open_link(self, url):
        """Abre uma URL no navegador padrão."""
        try:
            self.log(f"Abrindo link: {url}")
            webbrowser.open_new_tab(url)
        except Exception as e:
            self.log(f"Erro ao abrir o link {url}: {e}", "error")
            messagebox.showerror("Erro", f"Não foi possível abrir o link:\n{e}")
            
    def show_guide(self, guide_filename):
        """Abre um arquivo de guia PDF."""
        try:
            # Assume que os guias estão em uma pasta 'guias' ao lado do executável
            guide_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "guias", guide_filename)
            if os.path.exists(guide_path):
                self.log(f"Abrindo guia: {guide_path}")
                os.startfile(guide_path)
            else:
                self.log(f"Arquivo de guia não encontrado: {guide_path}", "error")
                messagebox.showerror("Arquivo não encontrado", f"O guia '{guide_filename}' não foi encontrado na pasta 'guias'.")
        except Exception as e:
            self.log(f"Erro ao abrir o guia: {e}", "error")
            messagebox.showerror("Erro", f"Não foi possível abrir o guia:\n{e}")

    # --- Lógica do BPA ---
    def check_bpa(self):
        try:
            self.log("Verificando BPA no servidor FTP...")
            files = self.list_ftp_files(FTP_PATH_BPA)
            if files is None: raise ConnectionError("Não foi possível listar arquivos do BPA.")

            bpa_files = [f for f in files if re.match(r'bpamag\d+\.exe', f, re.IGNORECASE)]
            if not bpa_files: raise FileNotFoundError("Nenhum instalador 'bpamag*.exe' encontrado.")
            
            latest_bpa_file = sorted(bpa_files, reverse=True)[0]
            self.latest_bpa_version = latest_bpa_file[:-4] # Remove .exe
            
            local_version = self.get_local_version(VERSION_FILE_BPA)
            
            def update_gui():
                if local_version == self.latest_bpa_version:
                    self.set_status(self.bpa_status_var, self.bpa_status_display, f"Instalado: {local_version} (Atualizado)", "green", self.bpa_action_button)
                else:
                    msg = f"Instalado: {local_version or 'Nenhum'}. Disponível: {self.latest_bpa_version}"
                    self.set_status(self.bpa_status_var, self.bpa_status_display, msg, "orange", self.bpa_action_button, "Baixar BPA")
                self.log(f"Verificação do BPA concluída. Versão online: {self.latest_bpa_version}")

            self.update_queue.put(update_gui)
        except Exception as e:
            self.update_queue.put(lambda: self.set_status(self.bpa_status_var, self.bpa_status_display, "Erro na verificação", "red", self.bpa_action_button))
            self.log(f"Erro ao verificar BPA: {e}", "error")

    def download_bpa(self):
        filename = self.latest_bpa_version + ".exe"
        self.handle_ftp_download_request(DIR_BPA, FTP_PATH_BPA, filename, version_file=VERSION_FILE_BPA, version_str=self.latest_bpa_version, callback=self.check_bpa)

    # --- Lógica do SIA ---
    def check_sia(self):
        try:
            self.log("Verificando SIA no servidor FTP...")
            files = self.list_ftp_files(FTP_PATH_SIA)
            if files is None: raise ConnectionError("Não foi possível listar arquivos do SIA.")

            sia_files = [f for f in files if re.match(r'instsia\d{4}\.exe', f, re.IGNORECASE)]
            if not sia_files: raise FileNotFoundError("Nenhum instalador 'instsia*.exe' encontrado.")
            
            self.latest_sia_file = sorted(sia_files, reverse=True)[0]
            self.latest_sia_version = self.latest_sia_file[:-4] # Remove .exe
            
            local_version = self.get_local_version(VERSION_FILE_SIA)

            def update_gui():
                if local_version == self.latest_sia_version:
                    self.set_status(self.sia_status_var, self.sia_status_display, f"Instalado: {local_version} (Atualizado)", "green", self.sia_action_button)
                else:
                    msg = f"Instalado: {local_version or 'Nenhum'}. Disponível: {self.latest_sia_version}"
                    self.set_status(self.sia_status_var, self.sia_status_display, msg, "orange", self.sia_action_button, "Baixar SIA")
                self.log(f"Verificação do SIA concluída. Versão online: {self.latest_sia_version}")

            self.update_queue.put(update_gui)
        except Exception as e:
            self.update_queue.put(lambda: self.set_status(self.sia_status_var, self.sia_status_display, "Erro na verificação", "red", self.sia_action_button))
            self.log(f"Erro ao verificar SIA: {e}", "error")

    def download_sia(self):
        self.handle_ftp_download_request(DIR_SIA, FTP_PATH_SIA, self.latest_sia_file, version_file=VERSION_FILE_SIA, version_str=self.latest_sia_version, callback=self.check_sia)

    # --- Lógica do FPO ---
    def check_fpo(self):
        try:
            self.log("Verificando FPO no servidor FTP...")
            files = self.list_ftp_files(FTP_PATH_FPO)
            if files is None: raise ConnectionError("Não foi possível listar arquivos do FPO.")

            fpo_installers = [f for f in files if "instalador" in f.lower() and f.endswith('.exe')]
            fpo_updates = [f for f in files if "instalador" not in f.lower() and f.endswith('.exe') and f.lower().startswith('fpo')]
            
            if not fpo_updates: raise FileNotFoundError("Nenhum arquivo de atualização do FPO encontrado.")
            
            self.fpo_installer_file = sorted(fpo_installers, reverse=True)[0] if fpo_installers else None
            self.latest_fpo_update_file = sorted(fpo_updates, reverse=True)[0]
            self.latest_fpo_version = self.latest_fpo_update_file[:-4]

            local_version = self.get_local_version(VERSION_FILE_FPO)

            def update_gui():
                if not os.path.exists(DIR_FPO) or not local_version:
                    msg = "Não instalado. É preciso baixar o instalador primeiro."
                    self.set_status(self.fpo_status_var, self.fpo_status_display, msg, "red", self.fpo_action_button, "Baixar Instalador FPO")
                    self.fpo_target_file = self.fpo_installer_file
                    self.fpo_target_version = "Instalador Base"
                elif local_version == self.latest_fpo_version:
                    self.set_status(self.fpo_status_var, self.fpo_status_display, f"Instalado: {local_version} (Atualizado)", "green", self.fpo_action_button)
                else:
                    msg = f"Instalado: {local_version}. Disponível: {self.latest_fpo_version}"
                    self.set_status(self.fpo_status_var, self.fpo_status_display, msg, "orange", self.fpo_action_button, "Baixar Atualização FPO")
                    self.fpo_target_file = self.latest_fpo_update_file
                    self.fpo_target_version = self.latest_fpo_version
                self.log(f"Verificação do FPO concluída.")

            self.update_queue.put(update_gui)
        except Exception as e:
            self.update_queue.put(lambda: self.set_status(self.fpo_status_var, self.fpo_status_display, "Erro na verificação", "red", self.fpo_action_button))
            self.log(f"Erro ao verificar FPO: {e}", "error")

    def download_fpo(self):
        if not hasattr(self, 'fpo_target_file') or not self.fpo_target_file:
            messagebox.showerror("Erro", "Não foi possível determinar o arquivo FPO para baixar. Tente verificar novamente.")
            return
        self.handle_ftp_download_request(DIR_FPO, FTP_PATH_FPO, self.fpo_target_file, version_file=VERSION_FILE_FPO, version_str=self.fpo_target_version, callback=self.check_fpo)

    # --- Lógica do BDSIA ---
    def check_bdsia(self):
        try:
            self.log("Listando versões do BDSIA...")
            files = self.list_ftp_files(FTP_PATH_SIA)
            if files is None: raise ConnectionError("Não foi possível listar arquivos do SIA/BDSIA.")
            
            bdsia_files = [f for f in files if re.match(r'BDSIA\d{6}[a-zA-Z]\.exe', f, re.IGNORECASE)]
            top_3_bdsia = sorted(bdsia_files, reverse=True)[:3]

            def update_gui():
                self.log(f"{len(top_3_bdsia)} versões recentes do BDSIA encontradas.")
                for i in range(3):
                    if i < len(top_3_bdsia):
                        filename = top_3_bdsia[i]
                        self.bdsia_labels[i].config(text=filename)
                        self.bdsia_buttons[i].config(state=tk.NORMAL, command=lambda f=filename: self.handle_ftp_download_request(DIR_SIA, FTP_PATH_SIA, f))
                    else:
                        self.bdsia_labels[i].config(text="-"*20)
                        self.bdsia_buttons[i].config(state=tk.DISABLED, command=None)

            self.update_queue.put(update_gui)
        except Exception as e:
            self.log(f"Erro ao buscar versões do BDSIA: {e}", "error")
            def clear_bdsia_gui():
                for i in range(3):
                    self.bdsia_labels[i].config(text="Erro ao buscar versões.")
                    self.bdsia_buttons[i].config(state=tk.DISABLED)
            self.update_queue.put(clear_bdsia_gui)
    
    # --- Lógica do CNES ---
    def download_cnes(self):
        url = "https://cnes.datasus.gov.br/EstatisticasServlet?path=SCNES4700-COMPLETA.ZIP"
        filename = "SCNES4700-COMPLETA.ZIP"
        self.handle_http_download_request(DIR_CNES, url, filename)

    # --- Lógica do Firebird ---
    def check_firebird_version(self):
        self.log("Verificando versão do Firebird...")
        try:
            # Caminhos comuns para o isql.exe do Firebird 1.5
            possible_paths = [
                "C:\\Program Files\\Firebird\\Firebird_1_5\\bin\\isql.exe",
                "C:\\Program Files (x86)\\Firebird\\Firebird_1_5\\bin\\isql.exe"
            ]
            isql_path = None
            for path in possible_paths:
                if os.path.exists(path):
                    isql_path = path
                    break
            
            if not isql_path:
                raise FileNotFoundError("O executável 'isql.exe' do Firebird não foi encontrado nos caminhos padrão.")

            # Usar subprocess para obter a versão
            result = subprocess.run([isql_path, "-z"], capture_output=True, text=True, check=True, creationflags=subprocess.CREATE_NO_WINDOW)
            version_line = result.stdout.strip().splitlines()[0]
            self.log(f"Versão do Firebird encontrada: {version_line}")
            self.update_queue.put(lambda: messagebox.showinfo("Versão do Firebird", version_line))
        except Exception as e:
            self.log(f"Não foi possível verificar a versão do Firebird: {e}", "error")
            self.update_queue.put(lambda: messagebox.showerror("Erro na Verificação", f"Não foi possível verificar a versão do Firebird.\nVerifique se ele está instalado ou se o programa tem permissão.\n\nErro: {e}"))

    def download_firebird(self):
        url = "https://cnes.datasus.gov.br/EstatisticasServlet?path=INSTALADORFIREBIRD-155.ZIP"
        filename = "INSTALADORFIREBIRD-155.ZIP"
        dest_dir = filedialog.askdirectory(title="Selecione uma pasta para salvar o instalador do Firebird")
        if dest_dir:
            self.handle_http_download_request(dest_dir, url, filename)

    def _manage_firebird_service(self, action):
        command = ["sc", action, "FirebirdServerDefaultInstance"]
        action_text = "iniciar" if action == "start" else "parar"
        self.log(f"Tentando {action_text} o serviço Firebird...")
        try:
            result = subprocess.run(command, capture_output=True, text=True, check=False, creationflags=subprocess.CREATE_NO_WINDOW)
            
            success = False
            if result.returncode == 0:
                success = True
            # Código 1062: o serviço não foi iniciado (sucesso ao tentar parar)
            # Código 1056: uma instância do serviço já está em execução (sucesso ao tentar iniciar)
            elif action == "stop" and "1062" in result.stderr:
                success = True
            elif action == "start" and "1056" in result.stderr:
                success = True

            if success:
                msg = f"Serviço Firebird {'iniciado' if action == 'start' else 'parado'} com sucesso."
                self.log(msg)
                self.update_queue.put(lambda: messagebox.showinfo("Sucesso", msg))
            else:
                raise OSError(result.stderr or result.stdout)
        except Exception as e:
            msg = f"Falha ao {action_text} o serviço Firebird. Tente executar o programa como Administrador."
            self.log(f"{msg}\nErro: {e}", "error")
            self.update_queue.put(lambda: messagebox.showerror("Erro de Permissão", f"{msg}\n\nDetalhes: {e}"))

    def start_firebird_service(self):
        self._manage_firebird_service("start")

    def stop_firebird_service(self):
        self._manage_firebird_service("stop")

    # --- Funções de Download ---
    def handle_ftp_download_request(self, dest_dir, ftp_path, filename, version_file=None, version_str=None, callback=None):
        """Pede confirmação e inicia a thread de download FTP."""
        save_path = os.path.join(dest_dir, filename)
        if messagebox.askyesno("Confirmação de Download", f"O arquivo '{filename}' será salvo em:\n'{dest_dir}'\n\nDeseja continuar?"):
            self.start_thread(self._ftp_download_worker, ftp_path, filename, save_path, version_file, version_str, callback)
        else:
            self.log(f"Download de {filename} cancelado pelo usuário.", "warning")

    def handle_http_download_request(self, dest_dir, url, filename, callback=None):
        """Pede confirmação e inicia a thread de download HTTP."""
        save_path = os.path.join(dest_dir, filename)
        if messagebox.askyesno("Confirmação de Download", f"O arquivo '{filename}' será salvo em:\n'{dest_dir}'\n\nDeseja continuar?"):
            self.start_thread(self._http_download_worker, url, save_path, callback)
        else:
            self.log(f"Download de {filename} cancelado pelo usuário.", "warning")

    def _ftp_download_worker(self, ftp_path, filename, save_path, version_file=None, version_str=None, callback=None):
        """Worker que executa o download FTP em uma thread separada."""
        try:
            self.log(f"Iniciando download de {filename} via FTP...")
            ftp = FTP(FTP_SERVER, timeout=30)
            ftp.login()
            ftp.cwd(ftp_path)
            
            with open(save_path, 'wb') as f:
                ftp.retrbinary(f'RETR {filename}', f.write)
            
            ftp.quit()
            self.log(f"Download de {filename} concluído com sucesso!", "info")
            self.update_queue.put(lambda: self.post_download_action(filename, save_path, version_file, version_str, callback))
        except Exception as e:
            self.log(f"Falha no download de {filename}: {e}", "error")
            self.update_queue.put(lambda: messagebox.showerror("Erro de Download", f"Ocorreu um erro no download FTP: {e}"))

    def _http_download_worker(self, url, save_path, callback=None):
        """Worker que executa o download HTTP em uma thread separada."""
        filename = os.path.basename(save_path)
        try:
            self.log(f"Iniciando download de {filename} via HTTP...")
            response = requests.get(url, stream=True, timeout=30)
            response.raise_for_status()
            with open(save_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            self.log(f"Download de {filename} concluído com sucesso!", "info")
            self.update_queue.put(lambda: self.post_download_action(filename, save_path, callback=callback))
        except Exception as e:
            self.log(f"Falha no download de {filename}: {e}", "error")
            self.update_queue.put(lambda: messagebox.showerror("Erro de Download", f"Ocorreu um erro no download HTTP: {e}"))

    def post_download_action(self, filename, save_path, version_file=None, version_str=None, callback=None):
        """Ações a serem executadas na thread principal após o download."""
        is_zip = filename.lower().endswith('.zip')
        if is_zip:
            if messagebox.askyesno("Sucesso", f"{filename} baixado com sucesso!\nDeseja extrair o conteúdo agora?"):
                self.extract_zip(save_path)
        else: # É .exe
            messagebox.showinfo("Sucesso", f"Instalador '{filename}' baixado com sucesso!\n\nAgora, execute o arquivo para instalar ou atualizar o programa.")
            if version_file and version_str:
                with open(version_file, 'w') as f:
                    f.write(version_str)
                self.log(f"Versão local atualizada para {version_str}.")
        
        if callback:
            self.start_thread(callback)

    def extract_zip(self, zip_path):
        """Extrai um arquivo zip para um diretório escolhido pelo usuário."""
        try:
            self.log(f"Solicitando local para extrair {os.path.basename(zip_path)}...")
            extract_dir = filedialog.askdirectory(title=f"Escolha onde extrair {os.path.basename(zip_path)}")
            if not extract_dir:
                self.log("Extração cancelada.", "warning")
                return

            self.log(f"Extraindo para {extract_dir}...")
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
            
            self.log("Extração concluída com sucesso.", "info")
            messagebox.showinfo("Sucesso", f"Arquivos extraídos com sucesso para:\n{extract_dir}")
        except Exception as e:
            self.log(f"Falha ao extrair arquivo zip: {e}", "error")
            messagebox.showerror("Erro de Extração", f"Não foi possível extrair o arquivo: {e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
