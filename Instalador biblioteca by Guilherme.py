import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
import subprocess
import sys
import os
import ctypes  
import time

def is_admin():
    
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()  
    except Exception:
       return False

if not is_admin():
    messagebox.showerror("Erro", "Abra o CMD como administrador e execute o script pelo CMD.")
    sys.exit()

class LibraryInstaller:
    def __init__(self, master):
        self.master = master
        self.master.title("Instalador de Bibliotecas Python")
        self.master.geometry("500x450")
        self.master.configure(bg="#eaeaea")

        self.libraries = []
        self.python_version = ""
        self.pip_status = ""

        
        self.title_label = tk.Label(master, text="Instalador de Bibliotecas", font=("Helvetica", 18, "bold"), bg="#eaeaea")
        self.title_label.pack(pady=10)

        
        self.instruction_label = tk.Label(master, text="Adicione as bibliotecas que deseja instalar:", bg="#eaeaea")
        self.instruction_label.pack()

        
        self.status_label = tk.Label(master, text="", bg="#eaeaea", font=("Arial", 10))
        self.status_label.pack(pady=5)

        
        self.listbox = tk.Listbox(master, width=60, height=10, bg="#ffffff", selectmode=tk.SINGLE)
        self.listbox.pack(pady=5)

        
        self.button_frame = tk.Frame(master, bg="#eaeaea")
        self.button_frame.pack(pady=10)

        self.add_button = tk.Button(self.button_frame, text="+ Adicionar Biblioteca", command=self.add_library, bg="#4CAF50", fg="white", font=("Arial", 12))
        self.add_button.pack(side=tk.LEFT, padx=5)

        self.add_multiple_button = tk.Button(self.button_frame, text="Adicionar Múltiplas", command=self.add_multiple_libraries, bg="#FF9800", fg="white", font=("Arial", 12))
        self.add_multiple_button.pack(side=tk.LEFT, padx=5)

        self.install_button = tk.Button(self.button_frame, text="Instalar Bibliotecas", command=self.install_libraries, bg="#2196F3", fg="white", font=("Arial", 12))
        self.install_button.pack(side=tk.LEFT, padx=5)

        
        self.progress = ttk.Progressbar(master, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

        
        self.footer_label = tk.Label(master, text="Desenvolvido por Guilherme Silva", bg="#eaeaea", font=("Arial", 8))
        self.footer_label.pack(side=tk.BOTTOM, pady=10)

        
        self.check_python_and_pip()

        
        self.log_file = "install_libs_log.txt"
        self.log("Iniciando o instalador de bibliotecas.")

    def normalize_message(message):
        replacements = {
            'ç': 'c',
            'á': 'a',
            'é': 'e',
            'í': 'i',
            'ó': 'o',
            'ú': 'u',
            'ã': 'a',
            'õ': 'o',
           
        }
        for original, replacement in replacements.items():
            message = message.replace(original, replacement)
        return message

    def log(self, message):
        try:
            with open(self.log_file, "a", encoding="utf-8") as log:
                log.write(f"{message}\n")
        except Exception as e:
            # Normaliza a mensagem apenas em caso de erro
            normalized_message = normalize_message(message)
            print(f"Erro ao escrever no log: {e}. Tentando registrar no formato sem acentuacao.")
            with open(self.log_file, "a", encoding="utf-8") as log:
                log.write(f"{normalized_message}\n")


    def check_python_and_pip(self):
        
        try:
            self.python_version = subprocess.check_output([sys.executable, "--version"]).decode().strip()
            self.pip_status = "pip está disponível."
        except Exception:
            user_profile = os.getenv("USERPROFILE")
            python_path = os.path.join(user_profile, "AppData", "Local", "Programs", "Python", "Python312", "python.exe")
            pip_path = os.path.join(user_profile, "AppData", "Local", "Programs", "Python", "Python312", "Scripts", "pip.exe")

            if os.path.exists(python_path):
                self.python_version = f"Python encontrado: {python_path}"
                self.pip_status = "pip encontrado: " + pip_path if os.path.exists(pip_path) else "pip não encontrado."
            else:
                self.python_version = "Python não encontrado."
                self.pip_status = "pip não encontrado."

        
        self.update_status(f"{self.python_version}\n{self.pip_status}")

    def update_status(self, message):
        self.status_label.config(text=message)

    def add_library(self):
        library = simpledialog.askstring("Adicionar Biblioteca", "Nome da biblioteca:")
        if library:
            library = library.strip()
            if library and library not in self.libraries:
                self.libraries.append(library)
                self.listbox.insert(tk.END, library)
                self.update_status(f"Biblioteca '{library}' adicionada.")
                self.log(f"Biblioteca '{library}' adicionada.")
            else:
                self.update_status(f"Biblioteca '{library}' já está na lista.")

    def add_multiple_libraries(self):
        libraries_input = simpledialog.askstring("Adicionar Múltiplas Bibliotecas", "Nomes das bibliotecas (separe por vírgula ou espaço):")
        if libraries_input:
            libraries = [lib.strip() for lib in libraries_input.replace(',', ' ').split()]
            for library in libraries:
                if library and library not in self.libraries:
                    self.libraries.append(library)
                    self.listbox.insert(tk.END, library)
                    self.update_status(f"Biblioteca '{library}' adicionada.")
                    self.log(f"Biblioteca '{library}' adicionada.")
                elif library in self.libraries:
                    self.update_status(f"Biblioteca '{library}' já está na lista.")

    def run_command(self, command):
        
        try:
            
            result = subprocess.run(command, shell=True, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            return result
        except subprocess.CalledProcessError as e:
            return e

    def install_libraries(self):
        self.stop_service("vpnagent")
    
        self.progress.start()
        total_libraries = len(self.libraries)
        installed_libraries = []
        failed_libraries = []

        for index, library in enumerate(self.libraries):
            result = self.run_command(f'"{sys.executable}" -m pip install {library}')
        
            if result.returncode == 0:
                installed_libraries.append(library)
                status_message = f"Instalação de '{library}' concluída."
                self.update_status(status_message)
                self.log(status_message)  # Log normalizado
            else:
                stderr_message = result.stderr.decode('latin-1', errors='replace').strip()  # Mantenha a decodificação aqui
                failed_libraries.append((library, "Erro na instalação"))
                error_message = f"Falha ao instalar '{library}': {stderr_message}"
                self.update_status(f"Falha ao instalar '{library}'.")
                self.log(error_message)  # Log normalizado

            self.progress['value'] = (index + 1) / total_libraries * 100
            self.master.update_idletasks()

        self.progress.stop()

        if installed_libraries:
            messagebox.showinfo("Instalação Completa", f"Bibliotecas instaladas com sucesso: {', '.join(installed_libraries)}")

        if failed_libraries:
            failed_messages = "\n".join([f"{lib}: {reason}" for lib, reason in failed_libraries])
            messagebox.showwarning("Instalação Incompleta", f"Algumas bibliotecas não foram instaladas corretamente:\n{failed_messages}")

        self.start_service("vpnagent")
        time.sleep(10)
        self.run_vpn()
        
    def stop_service(self, service_name):
        subprocess.run(["sc", "stop", service_name])
        if not self.is_service_running(service_name):
            messagebox.showinfo("Serviço Parado", f"O serviço {service_name} foi parado com sucesso.")

    def start_service(self, service_name):
        subprocess.run(["sc", "start", service_name])
        messagebox.showinfo("Serviço Iniciado", f"O serviço {service_name} foi iniciado.")

    def run_vpn(self):
        subprocess.Popen([r"C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"])
        self.update_status("VPN iniciado.")

    def is_service_running(self, service_name):
        result = subprocess.run(["sc", "query", service_name], capture_output=True, text=True)
        return "RUNNING" in result.stdout

if __name__ == "__main__":
    root = tk.Tk()
    app = LibraryInstaller(root)
    root.mainloop()
