import tkinter as tk
from tkinter import ttk, messagebox
import threading
import os
import sys
from PIL import Image, ImageTk
import customtkinter as ctk
import webbrowser
from datetime import datetime, timedelta
try:
    from auto_455 import main as auto_main
except ImportError:
    def auto_main():
        # Função placeholder para caso o módulo não esteja disponível durante testes
        import time
        print("Simulando automação...")
        time.sleep(5)
        return True

class ModernTooltip:
    """Classe para criar tooltips elegantes para elementos da interface"""
    
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)
    
    def show_tooltip(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        
        self.tooltip = ctk.CTkToplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        
        label = ctk.CTkLabel(
            self.tooltip, 
            text=self.text,
            fg_color="#303030",
            corner_radius=6,
            padx=10,
            pady=6
        )
        label.pack()
    
    def hide_tooltip(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

class LogDisplay(ctk.CTkScrollableFrame):
    """Widget para exibir logs da automação"""
    
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.logs = []
        self.max_logs = 100
        
    def add_log(self, message, log_type="info"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        # Definir cores com base no tipo de log
        colors = {
            "info": "#4CAF50",
            "warning": "#FF9800",
            "error": "#F44336",
            "success": "#2196F3"
        }
        color = colors.get(log_type, "#FFFFFF")
        
        # Criar frame para cada entrada de log
        log_frame = ctk.CTkFrame(self, fg_color="transparent")
        log_frame.pack(fill="x", padx=5, pady=2)
        
        # Adicionar timestamp
        time_label = ctk.CTkLabel(
            log_frame, 
            text=timestamp,
            width=60,
            font=("Roboto", 10),
            text_color="#888888"
        )
        time_label.pack(side="left", padx=(0, 10))
        
        # Adicionar mensagem com cor apropriada
        msg_label = ctk.CTkLabel(
            log_frame, 
            text=message,
            font=("Roboto", 11),
            text_color=color,
            anchor="w"
        )
        msg_label.pack(side="left", fill="x", expand=True)
        
        # Manter o número de logs limitado
        self.logs.append(log_frame)
        if len(self.logs) > self.max_logs:
            old_log = self.logs.pop(0)
            old_log.destroy()
        
        # Rolar para o último log
        self.after(100, lambda: self._parent_canvas.yview_moveto(1.0))


class AutomacaoApp:
    def __init__(self, root):
        # Configuração do CustomTkinter
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.root = root
        self.root.title("Automação SSW")
        self.root.geometry("700x500")
        self.root.minsize(600, 450)
        
        # Variáveis de controle
        self.running = False
        self.thread = None
        self.auto_count = 0
        self.start_time = None
        
        # Criar interface principal
        self.create_interface()
        
        # Protocolo de fechamento da janela
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Iniciar a atualização do cronômetro
        self.update_runtime()
        
        # Iniciar com um log de boas-vindas
        self.log_display.add_log("Sistema de automação inicializado com sucesso", "success")
        self.log_display.add_log("Aguardando comando para iniciar...", "info")

    def create_interface(self):
        # Container principal
        main_container = ctk.CTkFrame(self.root)
        main_container.pack(fill="both", expand=True, padx=15, pady=15)
        
        # Painel superior - Informações de status
        self.status_panel = ctk.CTkFrame(main_container)
        self.status_panel.pack(fill="x", padx=10, pady=10)
        
        # Título do aplicativo
        title_label = ctk.CTkLabel(
            self.status_panel, 
            text="AUTOMAÇÃO SSW",
            font=ctk.CTkFont(family="Roboto", size=20, weight="bold")
        )
        title_label.pack(pady=(10, 20))
        
        # Grid de estatísticas
        stats_frame = ctk.CTkFrame(self.status_panel, fg_color="transparent")
        stats_frame.pack(fill="x", padx=20, pady=5)
        stats_frame.grid_columnconfigure((0, 1, 2), weight=1)
        
        # Estatística 1: Status
        status_frame = ctk.CTkFrame(stats_frame, fg_color="#2a2d2e", corner_radius=10)
        status_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        
        ctk.CTkLabel(
            status_frame, 
            text="STATUS",
            font=ctk.CTkFont(family="Roboto", size=12)
        ).pack(pady=(10, 0))
        
        self.status_label = ctk.CTkLabel(
            status_frame, 
            text="PARADO",
            font=ctk.CTkFont(family="Roboto", size=16, weight="bold"),
            text_color="#FF5252"
        )
        self.status_label.pack(pady=(0, 10))
        
        # Estatística 2: Tempo de Execução
        time_frame = ctk.CTkFrame(stats_frame, fg_color="#2a2d2e", corner_radius=10)
        time_frame.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        
        ctk.CTkLabel(
            time_frame, 
            text="TEMPO DE EXECUÇÃO",
            font=ctk.CTkFont(family="Roboto", size=12)
        ).pack(pady=(10, 0))
        
        self.time_label = ctk.CTkLabel(
            time_frame, 
            text="00:00:00",
            font=ctk.CTkFont(family="Roboto", size=16, weight="bold")
        )
        self.time_label.pack(pady=(0, 10))
        
        # Estatística 3: Execuções
        count_frame = ctk.CTkFrame(stats_frame, fg_color="#2a2d2e", corner_radius=10)
        count_frame.grid(row=0, column=2, padx=10, pady=10, sticky="ew")
        
        ctk.CTkLabel(
            count_frame, 
            text="EXECUÇÕES",
            font=ctk.CTkFont(family="Roboto", size=12)
        ).pack(pady=(10, 0))
        
        self.count_label = ctk.CTkLabel(
            count_frame, 
            text="0",
            font=ctk.CTkFont(family="Roboto", size=16, weight="bold")
        )
        self.count_label.pack(pady=(0, 10))
        
        # Painel de botões de controle
        control_panel = ctk.CTkFrame(main_container)
        control_panel.pack(fill="x", padx=10, pady=(5, 10))
        
        # Botão iniciar/parar grande
        self.toggle_button = ctk.CTkButton(
            control_panel,
            text="INICIAR AUTOMAÇÃO",
            command=self.toggle_automation,
            height=40,
            fg_color="#4CAF50",
            hover_color="#388E3C",
            font=ctk.CTkFont(family="Roboto", size=14, weight="bold")
        )
        self.toggle_button.pack(side="left", padx=10, pady=10, fill="x", expand=True)
        
        # Botão de configurações
        self.settings_button = ctk.CTkButton(
            control_panel,
            text="",
            width=40,
            height=40,
            fg_color="#555555",
            hover_color="#666666",
            command=self.show_settings
        )
        self.settings_button.pack(side="right", padx=(5, 10), pady=10)
        ModernTooltip(self.settings_button, "Configurações")
        
        # Botão de ajuda
        self.help_button = ctk.CTkButton(
            control_panel,
            text="",
            width=40,
            height=40,
            fg_color="#555555",
            hover_color="#666666",
            command=self.show_help
        )
        self.help_button.pack(side="right", padx=(5, 0), pady=10)
        ModernTooltip(self.help_button, "Ajuda")
        
        # Botão de agendamento
        self.schedule_button = ctk.CTkButton(
            control_panel,
            text="🕒",
            width=40,
            height=40,
            fg_color="#555555",
            hover_color="#666666",
            command=self.show_schedule
        )
        self.schedule_button.pack(side="right", padx=(5, 0), pady=10)
        ModernTooltip(self.schedule_button, "Agendar Início")
        
        # Área de log
        log_container = ctk.CTkFrame(main_container)
        log_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        log_header = ctk.CTkFrame(log_container, fg_color="transparent")
        log_header.pack(fill="x", padx=10, pady=(5, 0))
        
        ctk.CTkLabel(
            log_header, 
            text="LOG DE ATIVIDADES",
            font=ctk.CTkFont(family="Roboto", size=12, weight="bold")
        ).pack(side="left")
        
        clear_log_btn = ctk.CTkButton(
            log_header,
            text="Limpar",
            font=ctk.CTkFont(size=11),
            width=60,
            height=25,
            command=self.clear_logs
        )
        clear_log_btn.pack(side="right")
        
        # Widget de logs
        self.log_display = LogDisplay(
            log_container,
            width=300,
            corner_radius=10,
            fg_color="#1a1a1a",
            border_width=1,
            border_color="#333333"
        )
        self.log_display.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Rodapé
        footer = ctk.CTkFrame(main_container, fg_color="transparent")
        footer.pack(fill="x", padx=10, pady=(0, 5))
        
        version_label = ctk.CTkLabel(
            footer, 
            text="v1.2.0",
            font=("Roboto", 10),
            text_color="#888888"
        )
        version_label.pack(side="right")
        
        # Substituir os textos por ícones (simulando utilização de ícones)
        # Em um cenário real, você usaria imagens reais para os ícones
        self.settings_button.configure(text="⚙️")
        self.help_button.configure(text="❓")

    def toggle_automation(self):
        if not self.running:
            # Iniciar automação
            self.running = True
            self.start_time = datetime.now()
            self.toggle_button.configure(
                text="PARAR AUTOMAÇÃO", 
                fg_color="#F44336",
                hover_color="#D32F2F"
            )
            self.status_label.configure(text="EM EXECUÇÃO", text_color="#4CAF50")
            
            # Log
            self.log_display.add_log("Automação iniciada", "success")
            
            # Iniciar thread da automação
            self.thread = threading.Thread(target=self.run_automation)
            self.thread.daemon = True
            self.thread.start()
        else:
            # Parar automação
            self.running = False
            self.toggle_button.configure(
                text="INICIAR AUTOMAÇÃO", 
                fg_color="#4CAF50",
                hover_color="#388E3C"
            )
            self.status_label.configure(text="PARADO", text_color="#FF5252")
            
            # Log
            self.log_display.add_log("Automação interrompida pelo usuário", "warning")
            
            if self.thread:
                self.thread.join(timeout=1)

    def run_automation(self):
        """Executa a automação uma única vez"""
        try:
            # Executar a automação
            self.log_display.add_log("Iniciando automação...", "info")
            
            # Chamar a função de automação
            result = auto_main()
            
            # Incrementar contador
            self.auto_count += 1
            self.root.after(0, lambda: self.count_label.configure(text=str(self.auto_count)))
            
            # Log de sucesso
            self.log_display.add_log("Automação completada com sucesso", "success")
            
            # Atualizar interface após conclusão
            self.running = False
            self.root.after(0, lambda: self.toggle_button.configure(
                text="INICIAR AUTOMAÇÃO", 
                fg_color="#4CAF50",
                hover_color="#388E3C"
            ))
            self.root.after(0, lambda: self.status_label.configure(
                text="PARADO", 
                text_color="#FF5252"
            ))
            
        except Exception as e:
            error_msg = f"Erro na automação: {str(e)}"
            self.log_display.add_log(error_msg, "error")
            self.running = False
            self.root.after(0, self.update_ui_after_error)

    def update_ui_after_error(self):
        """Atualiza a UI após um erro na automação"""
        self.toggle_button.configure(
            text="INICIAR AUTOMAÇÃO", 
            fg_color="#4CAF50",
            hover_color="#388E3C"
        )
        self.status_label.configure(text="ERRO", text_color="#F44336")
        self.running = False
        
        # Exibir popup de erro
        messagebox.showerror(
            "Erro de Automação",
            "Ocorreu um erro durante a execução da automação.\nConsulte o log para mais detalhes."
        )

    def update_runtime(self):
        """Atualiza o tempo de execução exibido"""
        if self.running and self.start_time:
            # Calcular tempo decorrido
            elapsed = datetime.now() - self.start_time
            # Formatar como HH:MM:SS
            hours, remainder = divmod(elapsed.seconds, 3600)
            minutes, seconds = divmod(remainder, 60)
            time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
            self.time_label.configure(text=time_str)
        
        # Agendar próxima atualização
        self.root.after(1000, self.update_runtime)

    def clear_logs(self):
        """Limpa os logs exibidos"""
        for log in self.log_display.logs:
            if log.winfo_exists():
                log.destroy()
        self.log_display.logs = []
        self.log_display.add_log("Log limpo pelo usuário", "info")

    def show_settings(self):
        """Exibe janela de configurações"""
        settings_window = ctk.CTkToplevel(self.root)
        settings_window.title("Configurações")
        settings_window.geometry("400x300")
        settings_window.resizable(False, False)
        settings_window.grab_set()  # Modal
        
        # Adicionar conteúdo às configurações
        frame = ctk.CTkFrame(settings_window)
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(
            frame, 
            text="Configurações",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=(0, 20))
        
        # Tema
        theme_frame = ctk.CTkFrame(frame)
        theme_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(
            theme_frame,
            text="Tema:",
            anchor="w"
        ).pack(side="left", padx=10)
        
        theme_var = ctk.StringVar(value="dark")
        theme_cb = ctk.CTkComboBox(
            theme_frame,
            values=["dark", "light", "system"],
            variable=theme_var,
            command=lambda choice: ctk.set_appearance_mode(choice)
        )
        theme_cb.pack(side="right", padx=10)
        
        # Outros controles de configuração podem ser adicionados aqui
        
        # Botão de fechar
        ctk.CTkButton(
            frame,
            text="Salvar e Fechar",
            command=settings_window.destroy
        ).pack(side="bottom", pady=20)

    def show_help(self):
        """Exibe janela de ajuda"""
        help_window = ctk.CTkToplevel(self.root)
        help_window.title("Ajuda")
        help_window.geometry("500x400")
        help_window.grab_set()  # Modal
        
        # Adicionar conteúdo de ajuda
        frame = ctk.CTkScrollableFrame(help_window)
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(
            frame, 
            text="Ajuda - Automação SSW",
            font=ctk.CTkFont(size=18, weight="bold")
        ).pack(pady=(0, 20))
        
        # Conteúdo da ajuda
        help_topics = [
            {
                "title": "Iniciando uma automação",
                "content": "Para iniciar a automação, clique no botão verde 'INICIAR AUTOMAÇÃO'. A automação será executada em segundo plano e você poderá acompanhar o progresso através do painel de logs."
            },
            {
                "title": "Parando uma automação",
                "content": "Para interromper uma automação em andamento, clique no botão vermelho 'PARAR AUTOMAÇÃO'. O sistema encerrará o ciclo atual e interromperá o processo."
            },
            {
                "title": "Entendendo os logs",
                "content": "O painel de logs exibe informações sobre a execução da automação. Diferentes cores indicam diferentes tipos de mensagens: verde para sucesso, vermelho para erros, laranja para avisos e azul para informações gerais."
            },
            {
                "title": "Configurações",
                "content": "No menu de configurações você pode ajustar o tema da interface e outras preferências do sistema de automação."
            },
            {
                "title": "Suporte",
                "content": "Em caso de problemas ou dúvidas, entre em contato com o suporte técnico pelo email suporte@sswautomacao.com.br"
            }
        ]
        
        for i, topic in enumerate(help_topics):
            topic_frame = ctk.CTkFrame(frame)
            topic_frame.pack(fill="x", pady=10)
            
            ctk.CTkLabel(
                topic_frame,
                text=topic["title"],
                font=ctk.CTkFont(size=14, weight="bold"),
                anchor="w"
            ).pack(fill="x", padx=10, pady=(10, 5))
            
            ctk.CTkLabel(
                topic_frame,
                text=topic["content"],
                anchor="w",
                justify="left",
                wraplength=440
            ).pack(fill="x", padx=15, pady=(0, 10))
        
        # Botão de fechar
        ctk.CTkButton(
            help_window,
            text="Fechar",
            command=help_window.destroy
        ).pack(side="bottom", pady=10)

    def show_schedule(self):
        """Exibe janela de agendamento"""
        schedule_window = ctk.CTkToplevel(self.root)
        schedule_window.title("Agendar Início")
        schedule_window.geometry("300x200")
        schedule_window.resizable(False, False)
        schedule_window.grab_set()  # Modal

        frame = ctk.CTkFrame(schedule_window)
        frame.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            frame, 
            text="Agendar Início da Automação",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=(0, 20))

        # Frame para hora e minuto
        time_frame = ctk.CTkFrame(frame, fg_color="transparent")
        time_frame.pack(fill="x", pady=10)

        # Spinbox para hora
        hour_var = tk.StringVar(value="00")
        hour_spinbox = ctk.CTkEntry(
            time_frame,
            textvariable=hour_var,
            width=50,
            justify="center"
        )
        hour_spinbox.pack(side="left", padx=5)

        ctk.CTkLabel(time_frame, text=":").pack(side="left")

        # Spinbox para minuto
        minute_var = tk.StringVar(value="00")
        minute_spinbox = ctk.CTkEntry(
            time_frame,
            textvariable=minute_var,
            width=50,
            justify="center"
        )
        minute_spinbox.pack(side="left", padx=5)

        def schedule_start():
            try:
                hour = int(hour_var.get())
                minute = int(minute_var.get())
                
                if hour < 0 or hour > 23 or minute < 0 or minute > 59:
                    raise ValueError("Horário inválido")

                now = datetime.now()
                schedule_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
                
                # Se o horário já passou hoje, agenda para amanhã
                if schedule_time <= now:
                    schedule_time += timedelta(days=1)

                delay_seconds = (schedule_time - now).total_seconds()
                
                self.log_display.add_log(
                    f"Automação agendada para {schedule_time.strftime('%H:%M')}",
                    "info"
                )
                
                # Agenda o início da automação
                self.root.after(int(delay_seconds * 1000), self.toggle_automation)
                schedule_window.destroy()
                
            except ValueError as e:
                messagebox.showerror(
                    "Erro",
                    "Por favor, insira um horário válido (HH:MM)"
                )

        # Botão de agendar
        ctk.CTkButton(
            frame,
            text="Agendar",
            command=schedule_start
        ).pack(side="bottom", pady=20)

    def on_closing(self):
        """Manipula o evento de fechamento da janela"""
        if self.running:
            # Perguntar se deseja realmente fechar
            if messagebox.askyesno("Confirmação", "A automação está em execução. Deseja realmente sair?"):
                self.running = False
                if self.thread:
                    self.thread.join(timeout=1)
                self.root.destroy()
                sys.exit(0)
        else:
            self.root.destroy()
            sys.exit(0)

if __name__ == "__main__":
    # Verificar se o módulo customtkinter está instalado
    try:
        import customtkinter
    except ImportError:
        root = tk.Tk()
        root.withdraw()
        if messagebox.askyesno(
            "Dependência não encontrada", 
            "Esta aplicação requer o módulo 'customtkinter' para uma interface moderna.\n\n"
            "Deseja instalar agora? (Requer conexão com internet)"
        ):
            import subprocess
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", "customtkinter", "pillow"])
                messagebox.showinfo(
                    "Instalação concluída", 
                    "Instalação concluída com sucesso. O aplicativo será iniciado."
                )
                # Reiniciar o aplicativo
                os.execv(sys.executable, ['python'] + sys.argv)
            except Exception as e:
                messagebox.showerror(
                    "Erro na instalação", 
                    f"Ocorreu um erro ao instalar as dependências: {str(e)}"
                )
                sys.exit(1)
        else:
            sys.exit(1)
    
    root = ctk.CTk()
    app = AutomacaoApp(root)
    root.mainloop()