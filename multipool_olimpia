import sys
import os
import socket
import getpass
import sqlite3
import datetime
import shutil
from collections import defaultdict
from PyQt5 import QtWidgets, QtCore, QtGui
from openpyxl import Workbook, load_workbook
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt

# ---- Caminho do banco robusto para .py e .exe (PyInstaller) ----
# CONFIG_FILE will be set after resource_path is defined

def ler_config_str(chave, padrao=""):
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                for linha in f:
                    if "=" in linha:
                        k, v = linha.strip().split("=", 1)
                        if k.strip().upper() == chave.upper():
                            return v.strip()
    except Exception:
        pass
    return padrao

def app_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))





# --- Connection helper for consistent PRAGMAs and timeouts (added by review) ---
def get_conn(db_path):
    conn = sqlite3.connect(db_path, timeout=7.0, isolation_level=None)
    cur = conn.cursor()
    try:
        cur.execute("PRAGMA journal_mode=WAL;")
    except Exception:
        pass
    cur.execute("PRAGMA busy_timeout=7000;")
    cur.execute("PRAGMA synchronous=NORMAL;")
    cur.execute("PRAGMA foreign_keys=ON;")
    return conn

def resource_path(relative_path):
    """Retorna caminho absoluto para arquivo, mesmo no executável."""
    if hasattr(sys, '_MEIPASS'):  # Quando rodando pelo PyInstaller
        base_path = sys._MEIPASS
    else:
        base_path = app_base_dir()
    return os.path.join(base_path, relative_path)

# Suprimir avisos do Qt sobre fontes
os.environ['QT_LOGGING_RULES'] = 'qt.qpa.plugin=false'
os.environ['QT_FONT_DPI'] = '96'

LOG_DIR = "logs"
BACKUP_DIR = "backups"
CONFIG_FILE = resource_path("config.txt")
DB_FILE = (ler_config_str("DB_PATH") or os.path.join(app_base_dir(), "dados", "multipool.db"))
os.makedirs(os.path.join(app_base_dir(), "dados"), exist_ok=True)
ONEDRIVE_FILE = "onedrive_path.txt"
LOGO_PATH = resource_path("logo.png")

def garantir_diretorio(path):
    os.makedirs(path, exist_ok=True)

def formatar_data_display(data_str):
    if not data_str:
        return "–"
    for fmt in ["%Y-%m-%d", "%d/%m/%Y"]:
        try:
            return datetime.datetime.strptime(data_str[:10], fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass
    return data_str

def normalizar_data(celula):
    if isinstance(celula, (datetime.datetime, datetime.date)):
        return celula.strftime("%Y-%m-%d")
    if not celula:
        return ""
    texto = str(celula)
    for fmt in ["%d/%m/%Y", "%Y-%m-%d"]:
        try:
            return datetime.datetime.strptime(texto[:10], fmt).strftime("%Y-%m-%d")
        except:
            pass
    return texto

def backup_banco():
    garantir_diretorio(BACKUP_DIR)
    if os.path.exists(DB_FILE):
        destino = os.path.join(BACKUP_DIR, f"backup_{datetime.datetime.now():%Y-%m-%d_%H%M%S}.db")
        shutil.copy2(DB_FILE, destino)

def registrar_log(acao, dados):
    garantir_diretorio(LOG_DIR)
    log_path = os.path.join(LOG_DIR, f"log_{datetime.date.today()}.txt")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.datetime.now()}] {acao}: {dados}\n")

def exportar_para_excel(dados, nome_arquivo):
    garantir_diretorio("exportacoes")
    caminho = os.path.join("exportacoes", f"{nome_arquivo}.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Datas"
    # Cabeçalhos incluindo campos internos
    ws.append([
        "Cotista", "Contato", "Empreendimento", "Entrada", "Saída", "Dormitório", 
        "Valor", "Disponível", "Fonte", "Nº da Cota", "Nº Apartamento", "Torre", 
        "Letra de Prioridade (HBS-Royal)"
    ])
    for linha in dados:
        # Incluir todos os campos, exceto ID e timestamp
        linha_export = linha[1:]  # Remove ID
        if len(linha_export) > 13:  # Remove timestamp se existir
            linha_export = linha_export[:13]
        ws.append(linha_export)
    wb.save(caminho)
    return caminho

def salvar_config(aba, criterio):
    with open(CONFIG_FILE, "w") as f:
        f.write(f"{aba}\n{criterio}")

def carregar_config():
    if not os.path.exists(CONFIG_FILE):
        return 0, "ENTRADA"
    try:
        with open(CONFIG_FILE, "r") as f:
            linhas = f.readlines()
            aba = int(linhas[0].strip()) if linhas else 0
            criterio = linhas[1].strip() if len(linhas) > 1 else "ENTRADA"
            return aba, criterio
    except:
        return 0, "ENTRADA"

class DatabaseManager:
    def __init__(self, db_file):
        self.db_file = db_file
        self.init_db()

    def init_db(self):
        # Abre a conexão com o banco
        with get_conn(self.db_file) as conn:
            # 1. Ativar WAL para mais segurança em múltiplos acessos
            conn.execute("PRAGMA journal_mode=WAL;")

            # 2. Criar a tabela se não existir
            conn.execute("""
                CREATE TABLE IF NOT EXISTS registros (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    cotista TEXT NOT NULL,
                    contato TEXT,
                    empreendimento TEXT,
                    entrada TEXT NOT NULL,
                    saida TEXT,
                    dormitorio TEXT,
                    valor TEXT,
                    disponivel TEXT DEFAULT 'Sim',
                    fonte TEXT DEFAULT 'Cliente',
                    numero_cota TEXT,
                    numero_apartamento TEXT,
                    torre TEXT,
                    letra_prioridade TEXT,
                    criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            # Indexes added by review for performance (separate executes)
            conn.execute("CREATE INDEX IF NOT EXISTS idx_registros_entrada ON registros(entrada);")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_registros_cotista ON registros(cotista COLLATE NOCASE);")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_registros_emp ON registros(empreendimento COLLATE NOCASE);")
            # 3. Garantir que as colunas novas existam (compatibilidade com bancos antigos)
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(registros)")
            colunas = [coluna[1] for coluna in cursor.fetchall()]

            # Colunas principais
            if 'disponivel' not in colunas:
                conn.execute("ALTER TABLE registros ADD COLUMN disponivel TEXT DEFAULT 'Sim'")
            if 'fonte' not in colunas:
                conn.execute("ALTER TABLE registros ADD COLUMN fonte TEXT DEFAULT 'Cliente'")

            # Campos internos adicionais
            if 'numero_cota' not in colunas:
                conn.execute("ALTER TABLE registros ADD COLUMN numero_cota TEXT")
            if 'numero_apartamento' not in colunas:
                conn.execute("ALTER TABLE registros ADD COLUMN numero_apartamento TEXT")
            if 'torre' not in colunas:
                conn.execute("ALTER TABLE registros ADD COLUMN torre TEXT")
            if 'letra_prioridade' not in colunas:
                conn.execute("ALTER TABLE registros ADD COLUMN letra_prioridade TEXT")

            conn.commit()

    def inserir(self, dados):
        # Garantir que temos 13 elementos (incluindo todos os campos)
        while len(dados) < 13:
            if len(dados) == 7:
                dados.append("Sim")  # disponivel
            elif len(dados) == 8:
                dados.append("Cliente")  # fonte
            elif len(dados) == 9:
                dados.append("")  # numero_cota
            elif len(dados) == 10:
                dados.append("")  # numero_apartamento
            elif len(dados) == 11:
                dados.append("")  # torre
            elif len(dados) == 12:
                dados.append("")  # letra_prioridade

        with get_conn(self.db_file) as conn:
            conn.execute("""
                INSERT INTO registros (cotista, contato, empreendimento, entrada, saida, dormitorio, valor, disponivel, fonte, numero_cota, numero_apartamento, torre, letra_prioridade)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, dados)
            registrar_log("INSERIR", f"Cotista: {dados[0]}, Entrada: {dados[3]}")

    def buscar_ordenado(self, criterio="ENTRADA"):
        allowed = {"ENTRADA": "entrada", "COTISTA": "cotista"}
        col = allowed.get(str(criterio).upper(), "entrada")
        with get_conn(self.db_file) as conn:
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT id, cotista, contato, empreendimento, entrada, saida, dormitorio, valor,
                        disponivel, fonte, numero_cota, numero_apartamento, torre, letra_prioridade
                FROM registros
                ORDER BY """ + col + " COLLATE NOCASE"
            )
            return cursor.fetchall()
    def atualizar(self, id_registro, dados):
        # Garantir que temos 13 elementos
        while len(dados) < 13:
            if len(dados) == 7:
                dados.append("Sim")
            elif len(dados) == 8:
                dados.append("Cliente")
            elif len(dados) == 9:
                dados.append("")
            elif len(dados) == 10:
                dados.append("")
            elif len(dados) == 11:
                dados.append("")
            elif len(dados) == 12:
                dados.append("")

        with get_conn(self.db_file) as conn:
            conn.execute("""
                UPDATE registros 
                SET cotista=?, contato=?, empreendimento=?, entrada=?, saida=?, dormitorio=?, valor=?, disponivel=?, fonte=?, numero_cota=?, numero_apartamento=?, torre=?, letra_prioridade=?
                WHERE id=?
            """, dados + [id_registro])
            registrar_log("ATUALIZAR", f"ID: {id_registro}, Cotista: {dados[0]}")

    def buscar_por_id(self, id_registro):
        with get_conn(self.db_file) as conn:
            cursor = conn.cursor()
            # Buscar apenas os campos necessários, excluindo o timestamp
            cursor.execute("""
                SELECT id, cotista, contato, empreendimento, entrada, saida, dormitorio, valor, 
                        disponivel, fonte, numero_cota, numero_apartamento, torre, letra_prioridade
                FROM registros WHERE id=?
            """, (id_registro,))
            return cursor.fetchone()

def existe_duplicata(self, cotista, entrada, empreendimento):
    with get_conn(self.db_file) as conn:
        cursor = conn.cursor()
        cursor.execute("""SELECT COUNT(*) FROM registros
                          WHERE cotista=? AND entrada=? AND COALESCE(empreendimento,'')=COALESCE(?, '')""",
                       (cotista, entrada, empreendimento))
        return cursor.fetchone()[0] > 0


    def excluir(self, id_registro):
        with get_conn(self.db_file) as conn:
            cur = conn.cursor()
            cur.execute("SELECT cotista FROM registros WHERE id=?", (id_registro,))
            row = cur.fetchone()
            cotista = row[0] if row else None
            conn.execute("DELETE FROM registros WHERE id=?", (id_registro,))
            registrar_log("EXCLUIR", f"ID: {id_registro}, Cotista: {cotista or 'N/D'}")
            return cotista
class MplCanvas(FigureCanvas):
    def __init__(self, width=12, height=8, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi, facecolor='#2d2d2d')
        super().__init__(self.fig)
        self.setParent(None)

class Proximos7DiasDialog(QtWidgets.QDialog):
    def __init__(self, proximos):
        super().__init__()
        self.setWindowTitle("Próximos 7 Dias")
        self.setMinimumSize(600, 400)
        self.setModal(True)
        
        layout = QtWidgets.QVBoxLayout()
        
        # Label com contador
        label = QtWidgets.QLabel(f"Encontrados {len(proximos)} registros nos próximos 7 dias:")
        label.setStyleSheet("font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(label)
        
        # Text edit para mostrar os dados
        text_edit = QtWidgets.QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText("\n\n".join(proximos))
        text_edit.setStyleSheet("font-family: 'Consolas', 'Monaco', monospace; font-size: 12px;")
        layout.addWidget(text_edit)
        
        # Botão fechar
        btn = QtWidgets.QPushButton("Fechar")
        btn.clicked.connect(self.accept)
        btn.setStyleSheet("QPushButton { padding: 8px 20px; }")
        layout.addWidget(btn)
        
        self.setLayout(layout)

class EditDialog(QtWidgets.QDialog):
    def __init__(self, dados=None):
        super().__init__()
        self.setWindowTitle("Adicionar/Editar Registro")
        self.setMinimumSize(550, 650)
        self.setModal(True)
        
        # Layout principal com scroll
        main_layout = QtWidgets.QVBoxLayout()
        
        # Área com scroll para os campos
        scroll_area = QtWidgets.QScrollArea()
        scroll_widget = QtWidgets.QWidget()
        layout = QtWidgets.QFormLayout()
        layout.setSpacing(10)
        
        self.inputs = {}
        campos = [
            ("COTISTA", QtWidgets.QLineEdit()),
            ("CONTATO", QtWidgets.QLineEdit()),
            ("EMPREENDIMENTO", QtWidgets.QLineEdit()),
            ("ENTRADA", QtWidgets.QDateEdit()),
            ("SAIDA", QtWidgets.QDateEdit()),
            ("DORMITORIO", QtWidgets.QLineEdit()),
            ("VALOR_CLIENTE", QtWidgets.QLineEdit()),
            ("DISPONIVEL", QtWidgets.QComboBox()),
            ("FONTE", QtWidgets.QComboBox()),
            # Campos internos adicionais
            ("NUMERO_COTA", QtWidgets.QLineEdit()),
            ("NUMERO_APARTAMENTO", QtWidgets.QLineEdit()),
            ("TORRE", QtWidgets.QLineEdit()),
            ("LETRA_PRIORIDADE", QtWidgets.QLineEdit())
        ]
        
        for nome, widget in campos:
            self.inputs[nome] = widget
            widget.setStyleSheet("padding: 5px; margin: 2px;")

            # Validações específicas
            if nome in ("ENTRADA", "SAIDA"):
                widget.setDate(QtCore.QDate.currentDate())
                widget.setCalendarPopup(True)
                widget.setDisplayFormat("dd/MM/yyyy")
                # Não usar setReadOnly em QDateEdit. Para impedir digitação, usamos a política de foco:
                widget.setFocusPolicy(QtCore.Qt.StrongFocus)
            elif nome == "CONTATO":
                # Validador para telefone: (11) 1111-1111 ou (11) 11111-1111
                regex_telefone = QtCore.QRegularExpression(r"^\(?\d{2}\)?\s?\d{4,5}-\d{4}$")
                validator_telefone = QtGui.QRegularExpressionValidator(regex_telefone)
                widget.setValidator(validator_telefone)
                widget.setPlaceholderText("Ex: (17) 3281-1234 ou (17) 99624-5935")
            elif nome == "DISPONIVEL":
                widget.addItems(["Sim", "Não"])
            elif nome == "FONTE":
                widget.addItems(["Cliente", "Lead Internet", "Terceiros"])
            elif nome == "LETRA_PRIORIDADE":
                widget.setPlaceholderText("Ex: HBS-Royal")

            # Labels personalizados
            label_mapping = {
                "COTISTA": "Cotista:",
                "CONTATO": "Contato:",
                "EMPREENDIMENTO": "Empreendimento:",
                "ENTRADA": "Data Entrada:",
                "SAIDA": "Data Saída:",
                "DORMITORIO": "Dormitório:",
                "VALOR_CLIENTE": "Valor:",
                "DISPONIVEL": "Disponível:",
                "FONTE": "Fonte:",
                "NUMERO_COTA": "Nº da Cota:",
                "NUMERO_APARTAMENTO": "Nº Apartamento:",
                "TORRE": "Torre:",
                "LETRA_PRIORIDADE": "Letra de Prioridade (HBS-Royal):"
            }
            
            label_text = label_mapping.get(nome, nome.replace("_", " ").title() + ":")
            
            # Separador visual antes dos campos internos
            if nome == "NUMERO_COTA":
                separador = QtWidgets.QLabel("━━━━━━━ INFORMAÇÕES INTERNAS ━━━━━━━")
                separador.setStyleSheet("color: #4CAF50; font-weight: bold; margin: 10px 0;")
                separador.setAlignment(QtCore.Qt.AlignCenter)
                layout.addRow("", separador)
            
            layout.addRow(label_text, widget)
        
        # Configurar scroll area
        scroll_widget.setLayout(layout)
        scroll_area.setWidget(scroll_widget)
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("QScrollArea { border: none; }")
        
        main_layout.addWidget(scroll_area)
        
        # Preencher dados se editando
        if dados:
            self.inputs["COTISTA"].setText(dados[1] or "")
            self.inputs["CONTATO"].setText(dados[2] or "")
            self.inputs["EMPREENDIMENTO"].setText(dados[3] or "")
            
            # Configurar datas
            try:
                entrada_date = QtCore.QDate.fromString(dados[4][:10], "yyyy-MM-dd")
                if entrada_date.isValid():
                    self.inputs["ENTRADA"].setDate(entrada_date)
            except:
                pass
            
            try:
                saida_date = QtCore.QDate.fromString(dados[5][:10], "yyyy-MM-dd")
                if saida_date.isValid():
                    self.inputs["SAIDA"].setDate(saida_date)
            except:
                pass
            
            self.inputs["DORMITORIO"].setText(dados[6] or "")
            self.inputs["VALOR_CLIENTE"].setText(dados[7] or "")
            
            # Configurar disponível
            disponivel = dados[8] if len(dados) > 8 and dados[8] else "Sim"
            index = self.inputs["DISPONIVEL"].findText(disponivel)
            if index >= 0:
                self.inputs["DISPONIVEL"].setCurrentIndex(index)
            
            # Configurar fonte
            fonte = dados[9] if len(dados) > 9 and dados[9] else "Cliente"
            index = self.inputs["FONTE"].findText(fonte)
            if index >= 0:
                self.inputs["FONTE"].setCurrentIndex(index)
            
            # Campos internos - já vêm sem timestamp
            self.inputs["NUMERO_COTA"].setText(dados[10] if len(dados) > 10 and dados[10] and not str(dados[10]).count(':') else "")
            self.inputs["NUMERO_APARTAMENTO"].setText(dados[11] if len(dados) > 11 and dados[11] and not str(dados[11]).count(':') else "")
            self.inputs["TORRE"].setText(dados[12] if len(dados) > 12 and dados[12] and not str(dados[12]).count(':') else "")
            self.inputs["LETRA_PRIORIDADE"].setText(dados[13] if len(dados) > 13 and dados[13] and not str(dados[13]).count(':') else "")
        
        # Botões
        btn_layout = QtWidgets.QHBoxLayout()
        btn_ok = QtWidgets.QPushButton("OK")
        btn_cancel = QtWidgets.QPushButton("Cancelar")
        
        btn_ok.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; padding: 8px 20px; border: none; border-radius: 4px; }")
        btn_cancel.setStyleSheet("QPushButton { background-color: #f44336; color: white; padding: 8px 20px; border: none; border-radius: 4px; }")
        
        btn_ok.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)
        
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        
        btn_widget = QtWidgets.QWidget()
        btn_widget.setLayout(btn_layout)
        main_layout.addWidget(btn_widget)
        
        self.setLayout(main_layout)

    def get_dados(self):
        return [
            self.inputs["COTISTA"].text().strip(),
            self.inputs["CONTATO"].text().strip(),
            self.inputs["EMPREENDIMENTO"].text().strip(),
            self.inputs["ENTRADA"].date().toString("yyyy-MM-dd"),
            self.inputs["SAIDA"].date().toString("yyyy-MM-dd"),
            self.inputs["DORMITORIO"].text().strip(),
            self.inputs["VALOR_CLIENTE"].text().strip(),
            self.inputs["DISPONIVEL"].currentText(),
            self.inputs["FONTE"].currentText(),
            self.inputs["NUMERO_COTA"].text().strip(),
            self.inputs["NUMERO_APARTAMENTO"].text().strip(),
            self.inputs["TORRE"].text().strip(),
            self.inputs["LETRA_PRIORIDADE"].text().strip()
        ]

class MultipoolOlimpiaApp(QtWidgets.QMainWindow):

    LOCK_FILE = os.path.join(os.path.dirname(DB_FILE), "db.lock")

    
    def check_lock(self):
        """
        Verifica se existe um lock ativo.
        Se o arquivo existir, testa se o banco ESTÁ realmente bloqueado para escrita.
        Se conseguir um BEGIN IMMEDIATE, considera lock obsoleto e remove.
        """
        if not os.path.exists(self.LOCK_FILE):
            return False

        # Tenta detectar lock "de verdade" no SQLite
        try:
            with get_conn(DB_FILE) as conn:
                try:
                    conn.execute("BEGIN IMMEDIATE;")
                    conn.execute("ROLLBACK;")
                    # Se conseguimos travar e soltar, o DB não está em uso. Lock file é obsoleto.
                    try:
                        os.remove(self.LOCK_FILE)
                    except Exception:
                        pass
                    return False
                except sqlite3.OperationalError as e:
                    if "locked" in str(e).lower():
                        return True
        except Exception:
            # Qualquer erro inesperado, trate como bloqueado para ser conservador
            return True

        # Como fallback, mantém bloqueio
        return True

    
    def create_lock(self):
        """Cria o arquivo de lock no mesmo diretório do DB com PID/host/timestamp."""
        try:
            info = (
                f"pid={os.getpid()} "
                f"user={getpass.getuser()} "
                f"host={socket.gethostname()} "
                f"at {datetime.datetime.now():%Y-%m-%d %H:%M:%S}"
            )
            with open(self.LOCK_FILE, "w", encoding="utf-8") as f:
                f.write(info)
                f.flush()
                os.fsync(f.fileno())
        except Exception as e:
            QtWidgets.QMessageBox.warning(
                self,
                "Aviso",
                f"Não foi possível criar o lock:\n{str(e)}\n"
                "O sistema continuará, mas poderá haver risco de conflito."
            )

    def remove_lock(self):
        """Remove o lock ao sair."""
        try:
            if os.path.exists(self.LOCK_FILE):
                os.remove(self.LOCK_FILE)
        except:
            pass
        

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Multipool Olímpia - Sistema de Gestão")
        # Verificar se já existe um lock criado
        if self.check_lock():
            QtWidgets.QMessageBox.warning(
                self,
                "Modo Somente Leitura",
                "Outro usuário já está usando o sistema para editar os dados.\n\n"
                "Este computador entrará em MODO LEITURA: você pode consultar, mas não pode alterar."
            )
            self.read_only = True
        else:
            # Cria o lock para indicar que este usuário está editando
            self.create_lock()
            self.read_only = False
        self.session_dirty = False  # flag de alterações não exportadas

        # Configurar ícone se existir
        if os.path.exists(LOGO_PATH):
            self.setWindowIcon(QtGui.QIcon(LOGO_PATH))

        # Inicializar banco e variáveis
        backup_banco()
        self.db = DatabaseManager(DB_FILE)
        self.ultimo_excluido = None
        _, criterio = carregar_config()
        self.criterio_ordenacao = criterio

        # Criar interface
        self.setup_ui()
        self.load_data()
        self.criar_toolbar()
        self.setup_shortcuts()
        self.atualizar_indicador_ordenacao()  # Inicializar indicador

    def setup_ui(self):
        # Widget principal
        main_widget = QtWidgets.QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QtWidgets.QVBoxLayout(main_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # Barra de pesquisa global
        self.setup_search_bar(main_layout)
        
        # Abas
        self.tabs = QtWidgets.QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Estilo para as abas
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #3d3d3d;
                background-color: #2d2d2d;
            }
            QTabBar::tab {
                background-color: #3d3d3d;
                color: white;
                padding: 8px 16px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: #4CAF50;
            }
        """)

        # Criar abas
        self.future_tab = QtWidgets.QWidget()
        self.past_tab = QtWidgets.QWidget()
        self.logs_tab = QtWidgets.QWidget()
        self.accounting_tab = QtWidgets.QWidget()

        self.tabs.addTab(self.future_tab, "📅 Datas Futuras")
        self.tabs.addTab(self.past_tab, "📋 Datas Passadas")
        self.tabs.addTab(self.logs_tab, "📝 Histórico")
        self.tabs.addTab(self.accounting_tab, "📊 Contabilidade")

        # Configurar cada aba
        self.setup_data_tab(self.future_tab, "future")
        self.setup_data_tab(self.past_tab, "past")
        self.setup_logs_tab()
        self.setup_accounting_tab()

        self.tabs.setCurrentIndex(0)

    def setup_search_bar(self, parent_layout):
        """Configurar barra de pesquisa global"""
        search_layout = QtWidgets.QHBoxLayout()
        search_layout.setSpacing(10)
        
        search_label = QtWidgets.QLabel("🔍 Pesquisar:")
        search_label.setStyleSheet("color: white; font-weight: bold; font-size: 14px;")
        
        self.search_input = QtWidgets.QLineEdit()
        self.search_input.setPlaceholderText("Digite nome, data, nº cota, apartamento, torre...")
        self.search_input.setMinimumHeight(40)
        self.search_input.setStyleSheet("""
            QLineEdit {
                padding: 10px;
                border: 2px solid #4CAF50;
                border-radius: 6px;
                background-color: #2d2d2d;
                color: white;
                font-size: 14px;
            }
            QLineEdit:focus {
                border-color: #66BB6A;
                background-color: #3d3d3d;
            }
        """)
        handler = getattr(self, 'filtrar_dados', None)
        if handler is None:
            def _noop():
                try:
                    self.load_data()
                except Exception:
                    pass
            handler = _noop
            self.filtrar_dados = handler
        self.search_input.textChanged.connect(handler)
        self.search_input.returnPressed.connect(handler)
        
        self.btn_limpar_pesquisa = QtWidgets.QPushButton("✖ Limpar")
        self.btn_limpar_pesquisa.setMinimumHeight(40)
        self.btn_limpar_pesquisa.setMinimumWidth(80)
        self.btn_limpar_pesquisa.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 10px 15px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
        """)
        self.btn_limpar_pesquisa.clicked.connect(self.limpar_pesquisa)
        
        # Tooltips
        self.search_input.setToolTip("Pesquisar em todos os campos, incluindo informações internas (Ctrl+F)")
        self.btn_limpar_pesquisa.setToolTip("Limpar pesquisa e mostrar todos os registros (Escape)")
        
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_input, 1)  # Expandir o campo de pesquisa
        search_layout.addWidget(self.btn_limpar_pesquisa)
        
        parent_layout.addLayout(search_layout)

    def setup_data_tab(self, tab, tab_type):
        layout = QtWidgets.QVBoxLayout()
        layout.setSpacing(10)
        
        # Barra de botões de ordenação
        buttons_layout = QtWidgets.QHBoxLayout()
        
        # Botões de ordenação
        btn_data_proxima = QtWidgets.QPushButton("📅 Data Mais Próxima")
        btn_alfabetica = QtWidgets.QPushButton("🔤 Ordem Alfabética")
        btn_salvar = QtWidgets.QPushButton("💾 Salvar Configuração")
        
        # Estilo dos botões
        button_style = """
            QPushButton {
                background-color: #4d4d4d;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #5d5d5d;
            }
            QPushButton:pressed {
                background-color: #4CAF50;
            }
        """
        
        btn_data_proxima.setStyleSheet(button_style)
        btn_alfabetica.setStyleSheet(button_style)
        btn_salvar.setStyleSheet(button_style)
        
        # Tooltips com atalhos
        btn_data_proxima.setToolTip("Ordenar por data mais próxima (Ctrl+D)")
        btn_alfabetica.setToolTip("Ordenar alfabeticamente por cotista (Ctrl+A)")
        btn_salvar.setToolTip("Salvar configuração atual (Ctrl+Shift+S)")
        
        # Conectar eventos
        btn_data_proxima.clicked.connect(lambda: self.ordenar_por_data())
        btn_alfabetica.clicked.connect(lambda: self.ordenar_alfabeticamente())
        btn_salvar.clicked.connect(lambda: self.salvar_configuracao())
        
        buttons_layout.addWidget(btn_data_proxima)
        buttons_layout.addWidget(btn_alfabetica)
        buttons_layout.addWidget(btn_salvar)
        
        # Indicador de ordenação atual
        self.label_ordenacao = QtWidgets.QLabel("📊 Ordenação: Data mais próxima")
        self.label_ordenacao.setStyleSheet("""
            QLabel {
                color: #4CAF50;
                font-weight: bold;
                padding: 8px;
                background-color: #3d3d3d;
                border-radius: 4px;
            }
        """)
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.label_ordenacao)
        
        layout.addLayout(buttons_layout)
        
        # Criar tabela
        table = QtWidgets.QTableWidget()
        table.setColumnCount(10)
        table.setHorizontalHeaderLabels([
            "ID", "Cotista", "Contato", "Empreendimento", 
            "Entrada", "Saída", "Dormitório", "Valor", "Disponível", "Fonte"
        ])
        
        # Configurar tabela
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        table.setSortingEnabled(False)
        table.setStyleSheet("""
            QTableWidget {
                gridline-color: #3d3d3d;
                selection-background-color: #4CAF50;
            }
            QHeaderView::section {
                background-color: #3d3d3d;
                color: white;
                padding: 8px;
                border: 1px solid #2d2d2d;
            }
        """)
        
        # Configurar header
        header = table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)  # Cotista
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.Stretch)  # Empreendimento
        
        # Definir larguras mínimas
        table.setColumnWidth(0, 50)   # ID
        table.setColumnWidth(1, 200)  # Cotista
        table.setColumnWidth(2, 150)  # Contato
        table.setColumnWidth(3, 200)  # Empreendimento
        table.setColumnWidth(4, 100)  # Entrada
        table.setColumnWidth(5, 100)  # Saída
        table.setColumnWidth(6, 80)   # Dormitório
        table.setColumnWidth(7, 100)  # Valor
        table.setColumnWidth(8, 100)  # Disponível
        table.setColumnWidth(9, 120)  # Fonte
        
        # Ocultar coluna ID
        table.setColumnHidden(0, True)
        
        # Conectar eventos
        table.cellDoubleClicked.connect(self.editar)
        
        layout.addWidget(table)
        tab.setLayout(layout)
        
        # Armazenar referência
        if tab_type == "future":
            self.future_table = table
        else:
            self.past_table = table

    def setup_logs_tab(self):
        layout = QtWidgets.QVBoxLayout()
        layout.setSpacing(10)
        
        # Botão refresh
        refresh_btn = QtWidgets.QPushButton("🔄 Atualizar Logs")
        refresh_btn.clicked.connect(self.carregar_logs)
        refresh_btn.setStyleSheet("QPushButton { padding: 8px 16px; }")
        layout.addWidget(refresh_btn)
        
        # Text area para logs
        self.logs_text = QtWidgets.QTextEdit()
        self.logs_text.setReadOnly(True)
        self.logs_text.setStyleSheet("""
            QTextEdit {
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: 11px;
                background-color: #1e1e1e;
                color: #d4d4d4;
                border: 1px solid #3d3d3d;
            }
        """)
        layout.addWidget(self.logs_text)
        
        self.logs_tab.setLayout(layout)

    def setup_accounting_tab(self):
        layout = QtWidgets.QVBoxLayout()
        layout.setSpacing(10)

        # ====== filtros ======
        filters_layout = QtWidgets.QHBoxLayout()

        # Data inicial
        self.filter_start = QtWidgets.QDateEdit()
        self.filter_start.setDate(QtCore.QDate.currentDate().addMonths(-6))
        self.filter_start.setCalendarPopup(True)
        self.filter_start.setDisplayFormat("dd/MM/yyyy")
        filters_layout.addWidget(QtWidgets.QLabel("Data inicial:"))
        filters_layout.addWidget(self.filter_start)

        # Data final
        self.filter_end = QtWidgets.QDateEdit()
        self.filter_end.setDate(QtCore.QDate.currentDate())
        self.filter_end.setCalendarPopup(True)
        self.filter_end.setDisplayFormat("dd/MM/yyyy")
        filters_layout.addWidget(QtWidgets.QLabel("Data final:"))
        filters_layout.addWidget(self.filter_end)

        # Empreendimento
        self.filter_empreendimento = QtWidgets.QLineEdit()
        self.filter_empreendimento.setPlaceholderText("Empreendimento (opcional)")
        self.filter_empreendimento.setMaximumWidth(250)
        filters_layout.addWidget(QtWidgets.QLabel("Empreendimento:"))
        filters_layout.addWidget(self.filter_empreendimento)

        layout.addLayout(filters_layout)
        # ====== fim filtros ======

        # Botão atualizar
        update_btn = QtWidgets.QPushButton("📈 Atualizar Gráficos")
        update_btn.clicked.connect(self.atualizar_graficos)
        update_btn.setStyleSheet("QPushButton { padding: 8px 16px; }")
        layout.addWidget(update_btn)

        # Canvas para gráficos
        self.canvas = MplCanvas(width=14, height=10)
        layout.addWidget(self.canvas)

        self.accounting_tab.setLayout(layout)

    def criar_toolbar(self):
        toolbar = self.addToolBar("Barra Principal")
        toolbar.setMovable(False)
        toolbar.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
        
        # Estilo da toolbar
        toolbar.setStyleSheet("""
            QToolBar {
                background-color: #3d3d3d;
                border: none;
                spacing: 5px;
                padding: 5px;
            }
            QToolButton {
                background-color: #4d4d4d;
                color: white;
                border: none;
                padding: 8px 12px;
                border-radius: 4px;
                margin: 2px;
            }
            QToolButton:hover {
                background-color: #5d5d5d;
            }
            QToolButton:pressed {
                background-color: #4CAF50;
            }
        """)
        
        # Ações CRUD
        add_action = QtWidgets.QAction("➕ Adicionar", self)
        add_action.setShortcut("Ctrl+N")
        add_action.triggered.connect(self.adicionar)
        toolbar.addAction(add_action)
        
        edit_action = QtWidgets.QAction("✏️ Editar", self)
        edit_action.setShortcut("F2")
        edit_action.triggered.connect(self.editar)
        toolbar.addAction(edit_action)
        
        delete_action = QtWidgets.QAction("🗑️ Excluir", self)
        delete_action.setShortcut("Delete")
        delete_action.triggered.connect(self.excluir)
        toolbar.addAction(delete_action)
        
        toolbar.addSeparator()

        # Desabilitar ações de escrita no modo leitura
        for act in (add_action, edit_action, delete_action):
            act.setEnabled(not self.read_only)
        
        # Ações de arquivo
        export_action = QtWidgets.QAction("📊 Exportar", self)
        export_action.setShortcut("Ctrl+S")
        export_action.triggered.connect(self.exportar_excel)
        toolbar.addAction(export_action)
        
        import_action = QtWidgets.QAction("📥 Importar", self)
        import_action.setShortcut("Ctrl+O")
        import_action.triggered.connect(self.importar_excel)
        toolbar.addAction(import_action)
        
        toolbar.addSeparator()
        
        # Ações de análise
        stats_action = QtWidgets.QAction("📈 Estatísticas", self)
        stats_action.triggered.connect(self.mostrar_estatisticas)
        toolbar.addAction(stats_action)
        
        alert_action = QtWidgets.QAction("⚠️ Próximos 7 Dias", self)
        alert_action.triggered.connect(self.mostrar_alerta_proximos_7dias)
        toolbar.addAction(alert_action)

    def setup_shortcuts(self):
        # Atalhos adicionais
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Q"), self, self.close)
        QtWidgets.QShortcut(QtGui.QKeySequence("F5"), self, self.load_data)
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+1"), self, lambda: self.tabs.setCurrentIndex(0))
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+2"), self, lambda: self.tabs.setCurrentIndex(1))
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+3"), self, lambda: self.tabs.setCurrentIndex(2))
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+4"), self, lambda: self.tabs.setCurrentIndex(3))
        
        # Atalhos para ordenação e salvar
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+D"), self, self.ordenar_por_data)
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+A"), self, self.ordenar_alfabeticamente)
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+Shift+S"), self, self.salvar_configuracao)
        
        # Atalhos para pesquisa
        QtWidgets.QShortcut(QtGui.QKeySequence("Ctrl+F"), self, self.focar_pesquisa)
        QtWidgets.QShortcut(QtGui.QKeySequence("Escape"), self, self.limpar_pesquisa)

    def closeEvent(self, event):
        # Sempre confirmar antes de fechar (mesmo sem alterações)
        if getattr(self, "read_only", False):
            reply = QtWidgets.QMessageBox.question(
                self,
                "Fechar sistema",
                "O sistema está em modo leitura. Deseja realmente sair?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                QtWidgets.QMessageBox.No
            )
            if reply == QtWidgets.QMessageBox.No:
                event.ignore()
                return
        elif getattr(self, "session_dirty", False):
            reply = QtWidgets.QMessageBox.question(
                self,
                "Sair sem exportar?",
                "As alterações já estão salvas no sistema (SQLite),\n"
                "mas ainda não foram exportadas para Excel.\n"
                "Deseja sair mesmo assim?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                QtWidgets.QMessageBox.No
            )
            if reply == QtWidgets.QMessageBox.No:
                event.ignore()
                return
        else:
            reply = QtWidgets.QMessageBox.question(
                self,
                "Fechar sistema",
                "Deseja realmente sair do sistema?",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                QtWidgets.QMessageBox.No
            )
            if reply == QtWidgets.QMessageBox.No:
                event.ignore()
                return

        try:
            if not getattr(self, "read_only", False):
                self.remove_lock()
        except Exception:
            pass
        event.accept()

    def get_current_table(self):
        return self.future_table if self.tabs.currentIndex() == 0 else self.past_table

    def load_data(self):
        try:
            registros = self.db.buscar_ordenado(self.criterio_ordenacao)
            hoje = datetime.date.today()
            
            # Salvar estado das colunas antes de limpar
            future_widths = []
            past_widths = []
            
            if hasattr(self, 'future_table'):
                header = self.future_table.horizontalHeader()
                for i in range(self.future_table.columnCount()):
                    future_widths.append(header.sectionSize(i))
            
            if hasattr(self, 'past_table'):
                header = self.past_table.horizontalHeader()
                for i in range(self.past_table.columnCount()):
                    past_widths.append(header.sectionSize(i))
            
            # Limpar tabelas
            self.future_table.setRowCount(0)
            self.past_table.setRowCount(0)
            
            for registro in registros:
                try:
                    data_entrada = datetime.datetime.strptime(registro[4][:10], "%Y-%m-%d").date()
                    table = self.future_table if data_entrada >= hoje else self.past_table
                    
                    row = table.rowCount()
                    table.insertRow(row)
                    
                    # Mostrar apenas os primeiros 10 campos (visíveis na tabela)
                    for col, valor in enumerate(registro[:10]):
                        if col == 4 or col == 5:  # Datas
                            valor = formatar_data_display(valor)
                        
                        item = QtWidgets.QTableWidgetItem(str(valor) if valor else "")
                        item.setData(QtCore.Qt.UserRole, registro[0])  # Armazenar ID
                        
                        # Colorir por disponibilidade
                        if col == 8:  # Coluna Disponível
                            if valor == "Não":
                                item.setBackground(QtGui.QColor(255, 200, 200))
                                item.setForeground(QtGui.QColor(0, 0, 0))  # Texto preto
                            else:
                                item.setBackground(QtGui.QColor(200, 255, 200))
                                item.setForeground(QtGui.QColor(0, 0, 0))  # Texto preto
                        
                        table.setItem(row, col, item)
                        
                except (ValueError, IndexError):
                    continue
            
            # Restaurar larguras das colunas
            if future_widths and hasattr(self, 'future_table'):
                header = self.future_table.horizontalHeader()
                for i, width in enumerate(future_widths):
                    if i < self.future_table.columnCount():
                        header.resizeSection(i, width)
            
            if past_widths and hasattr(self, 'past_table'):
                header = self.past_table.horizontalHeader()
                for i, width in enumerate(past_widths):
                    if i < self.past_table.columnCount():
                        header.resizeSection(i, width)
            
            # Atualizar status
            self.statusBar().showMessage(f"Carregados {len(registros)} registros")
                    
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Erro ao carregar dados: {str(e)}")

    def adicionar(self):
        if self.read_only:
            QtWidgets.QMessageBox.warning(
                self,
                "Somente Leitura",
                "O sistema está em modo leitura. Não é possível alterar dados agora."
            )
            return

        dialog = EditDialog()
        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            dados = dialog.get_dados()
            if dados[0]:  # Cotista obrigatório
                try:
                    self.db.inserir(dados)
                    self.load_data()
                    self.statusBar().showMessage("Registro adicionado com sucesso!")
                    self.session_dirty = True
                except Exception as e:
                    QtWidgets.QMessageBox.critical(self, "Erro", f"Erro ao adicionar: {str(e)}")
            else:
                QtWidgets.QMessageBox.warning(self, "Aviso", "O campo Cotista é obrigatório!")

    def exportar_excel(self, automatico=False, sufixo=""):
        try:
            registros = self.db.buscar_ordenado(self.criterio_ordenacao)
            if not registros:
                if not automatico:
                    QtWidgets.QMessageBox.information(self, "Aviso", "Nenhum registro encontrado para exportar!")
                return
            
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo = f"multipool_export_{timestamp}{sufixo}"
            
            caminho = exportar_para_excel(registros, nome_arquivo)
            if not automatico:
                QtWidgets.QMessageBox.information(
                    self, "Exportação Concluída", 
                    f"Arquivo exportado com sucesso!\n\nLocal: {caminho}\nRegistros: {len(registros)}"
                )
                self.statusBar().showMessage(f"Exportados {len(registros)} registros para Excel")
        except Exception as e:
            if not automatico:
                QtWidgets.QMessageBox.critical(self, "Erro na Exportação", f"Erro ao exportar: {str(e)}")

    def importar_excel(self):
        # openpyxl não lê .xls — limitamos a .xlsx para evitar erros de importação
        arquivo, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Selecionar arquivo Excel", "", "Arquivos Excel (*.xlsx)"
        )
        
        if not arquivo:
            return
        
        try:
            wb = load_workbook(arquivo)
            ws = wb.active
            
            registros_importados = 0
            erros = []
            
            # Criar barra de progresso
            progress = QtWidgets.QProgressDialog("Importando registros...", "Cancelar", 0, ws.max_row - 1, self)
            progress.setWindowModality(QtCore.Qt.WindowModal)
            
            for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                progress.setValue(row_num - 2)
                
                if progress.wasCanceled():
                    break
                
                try:
                    if not any(row[:7]):  # Pular linhas vazias
                        continue
                    
                    cotista = str(row[0]).strip() if row[0] else ""
                    contato = str(row[1]).strip() if row[1] else ""
                    empreendimento = str(row[2]).strip() if row[2] else ""
                    entrada = normalizar_data(row[3])
                    saida = normalizar_data(row[4])
                    dormitorio = str(row[5]).strip() if row[5] else ""
                    valor = str(row[6]).strip() if row[6] else ""
                    
                    # Novas colunas
                    disponivel = str(row[7]).strip() if len(row) > 7 and row[7] else "Sim"
                    fonte = str(row[8]).strip() if len(row) > 8 and row[8] else "Cliente"
                    
                    # Campos internos (opcionais no Excel)
                    numero_cota = str(row[9]).strip() if len(row) > 9 and row[9] else ""
                    numero_apartamento = str(row[10]).strip() if len(row) > 10 and row[10] else ""
                    torre = str(row[11]).strip() if len(row) > 11 and row[11] else ""
                    letra_prioridade = str(row[12]).strip() if len(row) > 12 and row[12] else ""
                    
                    # Validar valores
                    if disponivel not in ["Sim", "Não"]:
                        disponivel = "Sim"
                    if fonte not in ["Cliente", "Lead Internet", "Terceiros"]:
                        fonte = "Cliente"
                    
                    if not cotista:
                        erros.append(f"Linha {row_num}: Campo Cotista é obrigatório")
                        continue
                    
                    dados = [cotista, contato, empreendimento, entrada, saida, dormitorio, valor, disponivel, fonte, numero_cota, numero_apartamento, torre, letra_prioridade]
                    
                    if not self.db.existe_duplicata(cotista, entrada, empreendimento):
                        self.db.inserir(dados)
                        registros_importados += 1
                    else:
                        erros.append(f"Linha {row_num}: Duplicata - {cotista} em {formatar_data_display(entrada)}")
                        
                except Exception as e:
                    erros.append(f"Linha {row_num}: {str(e)}")
            
            progress.close()
            self.load_data()
            
            # Relatório da importação
            mensagem = f"Importação concluída!\n\nRegistros importados: {registros_importados}"
            if erros:
                mensagem += f"\nErros encontrados: {len(erros)}"
                if len(erros) <= 15:
                    mensagem += "\n\nDetalhes dos erros:\n" + "\n".join(erros)
                else:
                    mensagem += f"\n\nPrimeiros 15 erros:\n" + "\n".join(erros[:15])
                    mensagem += f"\n... e mais {len(erros) - 15} erros."
            
            QtWidgets.QMessageBox.information(self, "Resultado da Importação", mensagem)
            self.statusBar().showMessage(f"Importados {registros_importados} registros")
            if registros_importados > 0:
                self.session_dirty = True
            
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Erro na Importação", f"Erro ao importar arquivo:\n{str(e)}")

    def carregar_logs(self):
        try:
            logs_content = ""
            
            if os.path.exists(LOG_DIR):
                log_files = sorted([f for f in os.listdir(LOG_DIR) if f.endswith('.txt')], reverse=True)
                
                if log_files:
                    for log_file in log_files[:10]:  # Últimos 10 arquivos
                        log_path = os.path.join(LOG_DIR, log_file)
                        with open(log_path, 'r', encoding='utf-8') as f:
                            logs_content += f"{'='*50}\n{log_file}\n{'='*50}\n"
                            logs_content += f.read()
                            logs_content += "\n\n"
                else:
                    logs_content = "Nenhum arquivo de log encontrado."
            else:
                logs_content = "Diretório de logs não existe ainda.\nOs logs serão criados quando você fizer alterações nos dados."
            
            self.logs_text.setPlainText(logs_content)
            self.statusBar().showMessage("Logs atualizados")
            
        except Exception as e:
            self.logs_text.setPlainText(f"Erro ao carregar logs:\n{str(e)}")

    def atualizar_graficos(self):
        try:
            registros = self.db.buscar_ordenado("ENTRADA")

            # Aplicar filtros escolhidos na aba
            start_date = self.filter_start.date().toPyDate()
            end_date = self.filter_end.date().toPyDate()
            empreendimento_filtro = self.filter_empreendimento.text().strip().lower()

            # Filtrar registros
            filtrados = []
            for r in registros:
                try:
                    data_entrada = datetime.datetime.strptime(r[4][:10], "%Y-%m-%d").date()
                    if not (start_date <= data_entrada <= end_date):
                        continue
                    if empreendimento_filtro and empreendimento_filtro not in (r[3] or "").lower():
                        continue
                    filtrados.append(r)
                except:
                    pass
            # Agora sim, depois do loop
            registros = filtrados

            # Remover resumo antigo (se existir)
            if hasattr(self, "resumo_label"):
                self.resumo_label.setParent(None)

            # Calcular valor total
            total_valor = 0
            for r in registros:
                try:
                    valor_str = str(r[7]).replace(',', '.').replace('R$', '').strip()
                    total_valor += float(valor_str) if valor_str else 0
                except:
                    pass
            
            # Criar label de resumo
            resumo_texto = f"Registros: {len(registros)}  |  Valor Total: R$ {total_valor:,.2f}"
            self.resumo_label = QtWidgets.QLabel(resumo_texto)
            self.resumo_label.setStyleSheet("color: white; font-weight: bold; margin: 5px;")
            self.accounting_tab.layout().insertWidget(1, self.resumo_label)

            if not registros:
                QtWidgets.QMessageBox.information(self, "Aviso", "Nenhum dado encontrado para gerar gráficos.")
                return
            
            # Preparar dados para os gráficos
            from collections import defaultdict
            dados_por_mes = defaultdict(int)
            dados_disponibilidade = {"Sim": 0, "Não": 0}
            dados_fonte = {"Cliente": 0, "Lead Internet": 0, "Terceiros": 0}
            valores_por_mes = defaultdict(float)

            for registro in registros:
                try:
                    data_entrada = datetime.datetime.strptime(registro[4][:10], "%Y-%m-%d").date()
                    mes_ano = data_entrada.strftime("%Y-%m")
                    dados_por_mes[mes_ano] += 1

                    # Disponibilidade
                    disponivel = registro[8] if len(registro) > 8 and registro[8] else "Sim"
                    if disponivel in dados_disponibilidade:
                        dados_disponibilidade[disponivel] += 1

                    # Fonte
                    fonte = registro[9] if len(registro) > 9 and registro[9] else "Cliente"
                    if fonte in dados_fonte:
                        dados_fonte[fonte] += 1

                    # Valor
                    valor_str = str(registro[7]).replace(',', '.').replace('R$', '').strip()
                    valor = float(valor_str) if valor_str else 0
                    valores_por_mes[mes_ano] += valor
                
                except (ValueError, IndexError):
                    continue

            # Limpar e configurar figura
            self.canvas.fig.clear()
            self.canvas.fig.patch.set_facecolor('#2d2d2d')

            if sum(dados_por_mes.values()) > 0:
                import matplotlib.pyplot as plt
                plt.rcParams.update({
                    "text.color": "white",
                    "axes.labelcolor": "white",
                    "axes.edgecolor": "white",
                    "xtick.color": "white",
                    "ytick.color": "white",
                    "axes.titleweight": "bold",
                    "axes.titlecolor": "white",
                    "font.size": 11
                })
                
            # ---------- Gráfico 1: Registros por mês ----------
            ax1 = self.canvas.fig.add_subplot(2, 2, 1, facecolor='#2d2d2d')
            meses = sorted(dados_por_mes.keys())[-12:]
            valores = [dados_por_mes[mes] for mes in meses]

            bars = ax1.bar(meses, valores, color='#4CAF50', alpha=0.85, edgecolor='white')
            ax1.set_title('Registros por Mês', fontsize=13, color='white', pad=12)
            ax1.set_ylabel('Quantidade', color='white')
            ax1.tick_params(axis='x', rotation=30)
            for spine in ["top", "right"]:
                ax1.spines[spine].set_visible(False)
            for bar, valor in zip(bars, valores):
                ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1,
                         str(valor), ha='center', va='bottom', color='white', fontsize=9)
            
            # ---------- Gráfico 2: Disponibilidade ----------
            ax2 = self.canvas.fig.add_subplot(2, 2, 2, facecolor='#2d2d2d')
            if sum(dados_disponibilidade.values()) > 0:
                wedges, texts, autotexts = ax2.pie(
                    dados_disponibilidade.values(),
                    autopct='%1.1f%%',
                    startangle=90,
                    colors=['#4CAF50', '#f44336'],
                    wedgeprops={'edgecolor': 'white'}
                )
                ax2.set_title('Disponibilidade', fontsize=13, color='white', pad=12)
                ax2.legend(dados_disponibilidade.keys(), loc="center left", bbox_to_anchor=(1, 0.5))
                for autotext in autotexts:
                    autotext.set_color("white")
                    autotext.set_fontweight("bold")
            
            # ---------- Gráfico 3: Fontes ----------
            ax3 = self.canvas.fig.add_subplot(2, 2, 3, facecolor='#2d2d2d')
            if sum(dados_fonte.values()) > 0:
                wedges, texts, autotexts = ax3.pie(
                    dados_fonte.values(),
                    autopct='%1.1f%%',
                    startangle=90,
                    colors=['#FF9800', '#2196F3', '#9C27B0'],
                    wedgeprops={'edgecolor': 'white'}
                )
                ax3.set_title('Fontes', fontsize=13, color='white', pad=12)
                ax3.legend(dados_fonte.keys(), loc="center left", bbox_to_anchor=(1, 0.5))
                for autotext in autotexts:
                    autotext.set_color("white")
                    autotext.set_fontweight("bold")
            
            # ---------- Gráfico 4: Valores por mês ----------
            ax4 = self.canvas.fig.add_subplot(2, 2, 4, facecolor='#2d2d2d')
            if valores_por_mes:
                valores_mes = [valores_por_mes[mes] for mes in meses]
                ax4.plot(meses, valores_mes, marker='o', color='#00BCD4', linewidth=2, markersize=6)
                ax4.set_title('Valores por Mês (R$)', fontsize=13, color='white', pad=12)
                ax4.set_ylabel('Valor (R$)', color='white')
                ax4.tick_params(axis='x', rotation=30)
                for spine in ["top", "right"]:
                    ax4.spines[spine].set_visible(False)
                ax4.grid(True, alpha=0.2, color='white')

            self.canvas.fig.tight_layout(pad=2.0)
            self.canvas.draw()
            self.statusBar().showMessage("Gráficos atualizados")

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Erro nos Gráficos", f"Erro ao gerar gráficos:\n{str(e)}")
    
    def mostrar_estatisticas(self):
        try:
            registros = self.db.buscar_ordenado("ENTRADA")
            hoje = datetime.date.today()
            
            # Contadores
            total_registros = len(registros)
            futuras = passadas = proximos_7_dias = 0
            disponivel_sim = disponivel_nao = 0
            fonte_cliente = fonte_lead = fonte_terceiros = 0
            valor_total = 0
            
            for registro in registros:
                try:
                    data_entrada = datetime.datetime.strptime(registro[4][:10], "%Y-%m-%d").date()
                    dias_diferenca = (data_entrada - hoje).days
                    
                    # Contagem temporal
                    if dias_diferenca >= 0:
                        futuras += 1
                        if 0 <= dias_diferenca <= 7:
                            proximos_7_dias += 1
                    else:
                        passadas += 1
                    
                    # Disponibilidade
                    disp = registro[8] if len(registro) > 8 and registro[8] else "Sim"
                    if disp == "Sim":
                        disponivel_sim += 1
                    else:
                        disponivel_nao += 1
                    
                    # Fonte
                    fonte = registro[9] if len(registro) > 9 and registro[9] else "Cliente"
                    if fonte == "Cliente":
                        fonte_cliente += 1
                    elif fonte == "Lead Internet":
                        fonte_lead += 1
                    else:
                        fonte_terceiros += 1
                    
                    # Valor
                    try:
                        valor_str = str(registro[7]).replace(',', '.').replace('R$', '').strip()
                        valor = float(valor_str) if valor_str else 0
                        valor_total += valor
                    except:
                        pass
                        
                except (ValueError, IndexError):
                    passadas += 1
            
            estatisticas = f"""
📊 ESTATÍSTICAS GERAIS

📈 Total de registros: {total_registros:,}
🔮 Datas futuras: {futuras:,}
📅 Datas passadas: {passadas:,}
⚠️ Próximos 7 dias: {proximos_7_dias:,}

💰 Valor total: R$ {valor_total:,.2f}
💰 Valor médio: R$ {valor_total/total_registros if total_registros > 0 else 0:,.2f}

✅ DISPONIBILIDADE
Disponível (Sim): {disponivel_sim:,} ({(disponivel_sim/total_registros*100) if total_registros > 0 else 0:.1f}%)
Não disponível (Não): {disponivel_nao:,} ({(disponivel_nao/total_registros*100) if total_registros > 0 else 0:.1f}%)

🎯 FONTES DE LEADS
Cliente: {fonte_cliente:,} ({(fonte_cliente/total_registros*100) if total_registros > 0 else 0:.1f}%)
Lead Internet: {fonte_lead:,} ({(fonte_lead/total_registros*100) if total_registros > 0 else 0:.1f}%)
Terceiros: {fonte_terceiros:,} ({(fonte_terceiros/total_registros*100) if total_registros > 0 else 0:.1f}%)

📊 PERCENTUAIS GERAIS
Datas futuras: {(futuras/total_registros*100) if total_registros > 0 else 0:.1f}%
Urgentes (7 dias): {(proximos_7_dias/futuras*100) if futuras > 0 else 0:.1f}% das futuras
"""
            
            # Criar dialog customizado
            dialog = QtWidgets.QDialog(self)
            dialog.setWindowTitle("📊 Estatísticas do Sistema")
            dialog.setMinimumSize(500, 600)
            dialog.setModal(True)
            
            layout = QtWidgets.QVBoxLayout()
            
            text_edit = QtWidgets.QTextEdit()
            text_edit.setReadOnly(True)
            text_edit.setPlainText(estatisticas)
            text_edit.setStyleSheet("""
                QTextEdit {
                    font-family: 'Consolas', 'Monaco', monospace;
                    font-size: 12px;
                    background-color: #1e1e1e;
                    color: #d4d4d4;
                    border: 1px solid #3d3d3d;
                    padding: 15px;
                }
            """)
            layout.addWidget(text_edit)
            
            btn = QtWidgets.QPushButton("Fechar")
            btn.clicked.connect(dialog.accept)
            btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; padding: 8px 20px; border: none; border-radius: 4px; }")
            layout.addWidget(btn)
            
            dialog.setLayout(layout)
            dialog.exec_()
            
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Erro ao calcular estatísticas:\n{str(e)}")

    def mostrar_alerta_proximos_7dias(self):
        try:
            registros = self.db.buscar_ordenado("ENTRADA")
            hoje = datetime.date.today()
            proximos = []
            
            for registro in registros:
                try:
                    data_entrada = datetime.datetime.strptime(registro[4][:10], "%Y-%m-%d").date()
                    dias_diferenca = (data_entrada - hoje).days
                    
                    if 0 <= dias_diferenca <= 7:
                        disponivel = registro[8] if len(registro) > 8 and registro[8] else "Sim"
                        fonte = registro[9] if len(registro) > 9 and registro[9] else "Cliente"
                        
                        status_icon = "🟢" if disponivel == "Sim" else "🔴"
                        fonte_icon = {"Cliente": "👤", "Lead Internet": "🌐", "Terceiros": "🤝"}.get(fonte, "❓")
                        
                        info = f"{status_icon} {formatar_data_display(registro[4])} - {registro[1]} ({registro[3]})"
                        info += f"\n   📞 {registro[2]} | 🏠 {registro[6]} | 💰 {registro[7]}"
                        info += f"\n   {fonte_icon} {fonte} | ⏰ Em {dias_diferenca} dia(s)"
                        
                        if dias_diferenca == 0:
                            info += " (HOJE!)"
                        elif dias_diferenca == 1:
                            info += " (AMANHÃ!)"
                        
                        proximos.append(info)
                        
                except (ValueError, IndexError):
                    continue
            
            if proximos:
                dialog = Proximos7DiasDialog(proximos)
                dialog.exec_()
            else:
                QtWidgets.QMessageBox.information(
                    self, "Próximos 7 Dias", 
                    "🎉 Nenhum registro encontrado para os próximos 7 dias!"
                )
                
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Erro ao verificar próximos 7 dias:\n{str(e)}")

    def ordenar_por_data(self):
        """Ordenar registros por data de entrada (mais próxima primeiro)"""
        try:
            self.criterio_ordenacao = "ENTRADA"
            self.load_data()
            self.atualizar_indicador_ordenacao()
            self.statusBar().showMessage("Ordenado por data mais próxima")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Erro ao ordenar por data:\n{str(e)}")

    def ordenar_alfabeticamente(self):
        """Ordenar registros alfabeticamente por cotista"""
        try:
            self.criterio_ordenacao = "COTISTA"
            self.load_data()
            self.atualizar_indicador_ordenacao()
            self.statusBar().showMessage("Ordenado alfabeticamente por cotista")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Erro ao ordenar alfabeticamente:\n{str(e)}")

    def atualizar_indicador_ordenacao(self):
        """Atualizar o indicador visual da ordenação atual"""
        try:
            if hasattr(self, 'label_ordenacao'):
                if self.criterio_ordenacao == "ENTRADA":
                    self.label_ordenacao.setText("📊 Ordenação: Data mais próxima")
                else:
                    self.label_ordenacao.setText("📊 Ordenação: Ordem alfabética")
        except:
            pass

    def focar_pesquisa(self):
        """Focar no campo de pesquisa"""
        try:
            if hasattr(self, 'search_input'):
                self.search_input.setFocus()
                self.search_input.selectAll()
        except:
            pass

    def salvar_configuracao(self):
        """Salvar configuração atual (aba e critério de ordenação)"""
        try:
            aba_atual = self.tabs.currentIndex()
            salvar_config(aba_atual, self.criterio_ordenacao)
            
            # Feedback visual
            QtWidgets.QMessageBox.information(
                self, "Configuração Salva", 
                f"Configuração salva com sucesso!\n\n"
                f"Aba atual: {self.tabs.tabText(aba_atual)}\n"
                f"Ordenação: {'Data mais próxima' if self.criterio_ordenacao == 'ENTRADA' else 'Ordem alfabética'}"
            )
            self.statusBar().showMessage("Configuração salva com sucesso!")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Erro ao salvar configuração:\n{str(e)}")

    def filtrar_dados(self):
        """Filtrar dados baseado no texto de pesquisa"""
        try:
            if not hasattr(self, 'search_input'):
                return
                
            texto_pesquisa = self.search_input.text().lower().strip()
            
            if not texto_pesquisa:
                self.load_data()  # Mostrar todos os dados
                return
            
            # Buscar todos os registros incluindo campos internos
            todos_registros = self.db.buscar_ordenado(self.criterio_ordenacao)
            registros_filtrados = []
            
            for registro in todos_registros:
                campos_busca = [
                    str(registro[1]) if registro[1] else "",  # cotista
                    str(registro[2]) if registro[2] else "",  # contato
                    str(registro[3]) if registro[3] else "",  # empreendimento
                    str(registro[4]) if registro[4] else "",  # entrada
                    str(registro[5]) if registro[5] else "",  # saida
                    str(registro[6]) if registro[6] else "",  # dormitorio
                    str(registro[7]) if registro[7] else "",  # valor
                    str(registro[8]) if registro[8] else "",  # disponivel
                    str(registro[9]) if registro[9] else "",  # fonte
                    str(registro[10]) if len(registro) > 10 and registro[10] else "",  # numero_cota
                    str(registro[11]) if len(registro) > 11 and registro[11] else "",  # numero_apartamento
                    str(registro[12]) if len(registro) > 12 and registro[12] else "",  # torre
                    str(registro[13]) if len(registro) > 13 and registro[13] else "",  # letra_prioridade
                ]
                
                # Verificar se o texto de pesquisa está em algum campo
                if any(texto_pesquisa in campo.lower() for campo in campos_busca):
                    registros_filtrados.append(registro)
            
            # Método mais simples: recarregar tudo com filtro
            self.carregar_dados_filtrados(registros_filtrados)
            
            # Atualizar status
            total_encontrados = len(registros_filtrados)
            if hasattr(self, 'statusBar') and self.statusBar():
                self.statusBar().showMessage(f"Pesquisa: '{texto_pesquisa}' - {total_encontrados} registro(s) encontrado(s)")
            
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Erro na pesquisa:\n{str(e)}")





    def editar(self, *args, **kwargs):
        if self.read_only:
            QtWidgets.QMessageBox.warning(
                self,
                "Somente Leitura",
                "O sistema está em modo leitura. Não é possível alterar dados agora."
            )
            return

        table = self.get_current_table()
        current_row = table.currentRow()

        if current_row < 0:
            QtWidgets.QMessageBox.information(self, "Aviso", "Selecione um registro para editar.")
            return

        id_item = table.item(current_row, 0)
        if not id_item:
            QtWidgets.QMessageBox.information(self, "Aviso", "Selecione um registro válido.")
            return

        try:
            id_registro = int(id_item.text())
            registro = self.db.buscar_por_id(id_registro)
            if not registro:
                QtWidgets.QMessageBox.warning(self, "Aviso", "Registro não encontrado.")
                return

            # Guardar valores antigos para possível sincronização
            cotista_antigo = (registro[1] or "").strip()
            contato_antigo = (registro[2] or "").strip()

            dialog = EditDialog(registro)
            if dialog.exec_() != QtWidgets.QDialog.Accepted:
                # Usuário cancelou; não faz nada
                return

            dados = dialog.get_dados()
            if not dados or not dados[0]:
                QtWidgets.QMessageBox.warning(self, "Aviso", "O campo Cotista é obrigatório!")
                return

            # Atualiza o registro selecionado
            self.db.atualizar(id_registro, dados)

            # Atualiza a interface SEMPRE (evita impressão de que não salvou)
            self.load_data()
            self.statusBar().showMessage("Registro atualizado com sucesso!")
            self.session_dirty = True

            # Sincronizar outros registros (apenas se cotista e contato não vazios)
            cotista_editado = (dados[0] or "").strip()
            contato_editado = (dados[1] or "").strip()
            if cotista_editado and contato_editado and (cotista_antigo or contato_antigo):
                resp = QtWidgets.QMessageBox.question(
                    self,
                    "Sincronizar?",
                    "Deseja atualizar COTISTA e CONTATO também em outros registros com os mesmos valores antigos?",
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                    QtWidgets.QMessageBox.No
                )
                if resp == QtWidgets.QMessageBox.Yes:
                    try:
                        with get_conn(DB_FILE) as conn:
                            conn.execute("BEGIN IMMEDIATE;")
                            conn.execute(
                                """
                                UPDATE registros
                                SET cotista=?, contato=?
                                WHERE id <> ? AND cotista=? AND contato=?
                                """,
                                (cotista_editado, contato_editado, id_registro, cotista_antigo, contato_antigo)
                            )
                            conn.commit()
                        self.statusBar().showMessage("Registro atualizado e sincronizado com outros iguais!")
                    except Exception as e:
                        QtWidgets.QMessageBox.warning(self, "Aviso", f"Falha ao sincronizar registros iguais:\n{str(e)}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Erro ao editar: {str(e)}")

    def excluir(self):
        if self.read_only:
            QtWidgets.QMessageBox.warning(
                self,
                "Somente Leitura",
                "O sistema está em modo leitura. Não é possível alterar dados agora."
            )
            return

        table = self.get_current_table()
        current_row = table.currentRow()

        if current_row >= 0:
            id_item = table.item(current_row, 0)
            cotista_item = table.item(current_row, 1)

            if id_item and cotista_item:
                resposta = QtWidgets.QMessageBox.question(
                    self, "Confirmar Exclusão", 
                    f"Tem certeza que deseja excluir o registro de:\n{cotista_item.text()}?",
                    QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                    QtWidgets.QMessageBox.No
                )

                if resposta == QtWidgets.QMessageBox.Yes:
                    try:
                        id_registro = int(id_item.text())
                        cotista_excluido = self.db.excluir(id_registro)
                        if cotista_excluido:
                            self.ultimo_excluido = id_registro
                            self.load_data()
                            self.statusBar().showMessage(f"Registro de {cotista_excluido} excluído!")
                            self.session_dirty = True
                    except Exception as e:
                        QtWidgets.QMessageBox.critical(self, "Erro", f"Erro ao excluir: {str(e)}")
        else:
            QtWidgets.QMessageBox.information(self, "Aviso", "Selecione um registro para excluir.")

    def carregar_dados_filtrados(self, registros_filtrados):
        """Carregar dados filtrados nas tabelas"""
        try:
            from datetime import datetime, date

            # Separar registros por categoria
            hoje = date.today()
            futuras = []
            passadas = []

            for registro in registros_filtrados:
                try:
                    data_entrada = datetime.strptime(registro[4][:10], "%Y-%m-%d").date()
                    if data_entrada >= hoje:
                        futuras.append(registro)
                    else:
                        passadas.append(registro)
                except:
                    futuras.append(registro)

            # Carregar dados nas tabelas
            self.carregar_tabela_com_dados(self.future_table, futuras)
            self.carregar_tabela_com_dados(self.past_table, passadas)

            # Atualizar status
            if hasattr(self, 'statusBar') and self.statusBar():
                self.statusBar().showMessage(f"Filtro aplicado - {len(registros_filtrados)} registro(s)")

        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Erro ao carregar dados filtrados:\n{str(e)}")

    def carregar_tabela_com_dados(self, table, registros):
        """Carregar dados em uma tabela específica"""
        try:
            # Salvar estado atual das colunas
            header = table.horizontalHeader()
            column_widths = []
            for i in range(table.columnCount()):
                column_widths.append(header.sectionSize(i))

            table.setRowCount(0)

            for row, registro in enumerate(registros):
                table.insertRow(row)

                # Mostrar apenas os primeiros 10 campos
                for col, valor in enumerate(registro[:10]):
                    if col == 4 or col == 5:  # Datas
                        valor = formatar_data_display(valor)

                    item = QtWidgets.QTableWidgetItem(str(valor) if valor else "")
                    item.setData(QtCore.Qt.UserRole, registro[0])  # Armazenar ID

                    # Colorir por disponibilidade
                    if col == 8:  # Coluna Disponível
                        if valor == "Não":
                            item.setBackground(QtGui.QColor(255, 200, 200))
                            item.setForeground(QtGui.QColor(0, 0, 0))  # Texto preto
                        else:
                            item.setBackground(QtGui.QColor(200, 255, 200))
                            item.setForeground(QtGui.QColor(0, 0, 0))  # Texto preto

                    table.setItem(row, col, item)

            # Restaurar larguras das colunas
            for i, width in enumerate(column_widths):
                if i < table.columnCount():
                    header.resizeSection(i, width)

        except Exception as e:
            print(f"Erro ao carregar tabela: {e}")

    def limpar_pesquisa(self):
        """Limpar campo de pesquisa e mostrar todos os dados"""
        try:
            if hasattr(self, 'search_input'):
                self.search_input.clear()
            self.load_data()
            if hasattr(self, 'statusBar') and self.statusBar():
                self.statusBar().showMessage("Pesquisa limpa - Mostrando todos os registros")
        except Exception as e:
            print(f"Erro ao limpar pesquisa: {e}")
            QtWidgets.QMessageBox.critical(self, "Erro", f"Erro ao limpar pesquisa:\n{str(e)}")

def main():
    # Configurações iniciais
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    QtWidgets.QApplication.setAttribute(QtCore.Qt.AA_UseHighDpiPixmaps)
    
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName("Multipool Olímpia")
    app.setApplicationVersion("2.0")
    app.setOrganizationName("Multipool")
    
    # Configurar estilo
    app.setStyle('Fusion')
    
    # Tema escuro personalizado
    palette = QtGui.QPalette()
    palette.setColor(QtGui.QPalette.Window, QtGui.QColor(45, 45, 45))
    palette.setColor(QtGui.QPalette.WindowText, QtGui.QColor(255, 255, 255))
    palette.setColor(QtGui.QPalette.Base, QtGui.QColor(30, 30, 30))
    palette.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor(60, 60, 60))
    palette.setColor(QtGui.QPalette.ToolTipBase, QtGui.QColor(0, 0, 0))
    palette.setColor(QtGui.QPalette.ToolTipText, QtGui.QColor(255, 255, 255))
    palette.setColor(QtGui.QPalette.Text, QtGui.QColor(255, 255, 255))
    palette.setColor(QtGui.QPalette.Button, QtGui.QColor(60, 60, 60))
    palette.setColor(QtGui.QPalette.ButtonText, QtGui.QColor(255, 255, 255))
    palette.setColor(QtGui.QPalette.BrightText, QtGui.QColor(255, 0, 0))
    palette.setColor(QtGui.QPalette.Link, QtGui.QColor(76, 175, 80))
    palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor(76, 175, 80))
    palette.setColor(QtGui.QPalette.HighlightedText, QtGui.QColor(0, 0, 0))
    app.setPalette(palette)
    
    # CSS global
    app.setStyleSheet("""
        QMainWindow {
            background-color: #2d2d2d;
        }
        QMenuBar {
            background-color: #3d3d3d;
            color: white;
            border: none;
        }
        QMenuBar::item {
            background-color: transparent;
            padding: 6px 12px;
        }
        QMenuBar::item:selected {
            background-color: #4CAF50;
        }
        QStatusBar {
            background-color: #3d3d3d;
            color: white;
            border-top: 1px solid #5d5d5d;
        }
        QMessageBox {
            background-color: #2d2d2d;
            color: white;
        }
        QMessageBox QPushButton {
            background-color: #4CAF50;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            min-width: 80px;
        }
        QMessageBox QPushButton:hover {
            background-color: #45a049;
        }
    """)
    
    try:
        # Criar e mostrar aplicação
        window = MultipoolOlimpiaApp()
        window.show()
        window.statusBar().showMessage("Sistema iniciado com sucesso!")
        
        # Executar aplicação
        sys.exit(app.exec_())
        
    except Exception as e:
        QtWidgets.QMessageBox.critical(None, "Erro Fatal", f"Erro ao iniciar aplicação:\n{str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
