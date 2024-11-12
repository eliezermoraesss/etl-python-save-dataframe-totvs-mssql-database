import os
import re
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox

import pandas as pd
import pyodbc


def setup_mssql():
    caminho_do_arquivo = (r"\\192.175.175.4\desenvolvimento\REPOSITORIOS\resources\application-properties"
                          r"\USER_PASSWORD_MSSQL_PROD.txt")
    try:
        with open(caminho_do_arquivo, 'r') as arquivo:
            string_lida = arquivo.read()
            username_txt, password_txt, database_txt, server_txt = string_lida.split(';')
            return username_txt, password_txt, database_txt, server_txt

    except FileNotFoundError:
        exibir_mensagem("EUREKA®", f"Erro ao ler credenciais de acesso ao banco de dados MSSQL.\n\nBase de "
                                   f"dados ERP TOTVS PROTHEUS.\n\nPor favor, informe ao desenvolvedor/TI "
                                   f"sobre o erro exibido.\n\nTenha um bom dia! ツ", 'error')
        sys.exit()

    except Exception as ex:
        exibir_mensagem("EUREKA®", f"Ocorreu um erro ao ler o arquivo: {ex}", 'error')
        sys.exit()


def get_env_var_windows(env_var):
    return os.getenv(env_var)


def get_excel_filepath(filename):
    base_path = os.environ.get('TEMP')
    return os.path.join(base_path, filename + '.xlsm')


def qp_validate(codigo_qp):
    padrao = r'^QP-E\d{4}$'
    return True if re.match(padrao, codigo_qp) else False


def delete_file(file_path):
    if os.path.exists(file_path):
        os.remove(file_path)


def exibir_mensagem(title, message, icon_type):
    window = tk.Tk()
    window.withdraw()
    window.lift()  # Garante que a janela esteja na frente
    window.title(title)
    window.attributes('-topmost', True)

    if icon_type == 'info':
        messagebox.showinfo(title, message)
    elif icon_type == 'warning':
        messagebox.showwarning(title, message)
    elif icon_type == 'error':
        messagebox.showerror(title, message)

    window.destroy()


def verify_if_baseline_exists(codigo_qp):
    try:
        with pyodbc.connect(
                f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};'
                f'PWD={password}') as conn:
            query = f"SELECT cod_qp FROM enaplic_management.dbo.tb_baseline WHERE cod_qp LIKE '{codigo_qp}%'"
            cursor = conn.cursor()
            result = cursor.execute(query).fetchone()
            return result if result is not None else None
    except Exception as ex:
        exibir_mensagem(f"Eureka® Falha de transação com banco de dados",
                        f"Não foi possível consultar a baseline {codigo_qp} na tabela tb_baseline.\n\n{str(ex)}"
                        f"\n\nInforme o administrador do sistema.", 'error')
        return None


def delete_if_baseline_exists(codigo_qp):
    try:
        with pyodbc.connect(
                f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};'
                f'PWD={password}') as conn:
            query = f"DELETE FROM enaplic_management.dbo.tb_baseline WHERE cod_qp LIKE '{codigo_qp}%'"
            cursor = conn.cursor()
            cursor.execute(query)
            return True
    except Exception as ex:
        exibir_mensagem(f"Eureka® Falha de transação com banco de dados",
                        f"Não foi possível excluir a baseline {codigo_qp} na tabela tb_baseline.\n\n{str(ex)}"
                        f"\n\nInforme o administrador do sistema.", 'error')
        return False


def insert_baseline(self, dataframe):
    try:
        with pyodbc.connect(
                f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};'
                f'PWD={password}') as conn:
            cursor = conn.cursor()
            total_rows = len(dataframe)

            for index, row in dataframe.iterrows():
                progress = 25 + (index / (total_rows - 1)) * 74
                self.update_progress(progress)
                insert_query = """
                INSERT INTO enaplic_management.dbo.tb_baseline (
                    cod_qp, equipamento, grupo, nivel, codigo, codigo_pai, descricao, tipo, qtde_bl, unid_medida,
                    especificacoes, qtde_proj, qtde_total, status, status_op
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""

                cursor.execute(insert_query, row['QP'], row['EQUIPAMENTO'], row['GRUPO'], row['NIVEL'],
                               row['CÓDIGO'], row['CÓDIGO PAI'], row['DESCRIÇÃO'], row['TIPO'], row['QTDE\nBL'],
                               row['UND'], row['ESPECIFICAÇÕES'], row['QTDE PROJ.'], row['QTDE\nTOTAL'],
                               row['STATUS'], row['STATUS_OP'])
            conn.commit()

    except Exception as ex:
        message = f"Erro ao inserir dados, transação revertida: {ex}\nindex: {index}\nrow: {row}"
        print(message)
        conn.rollback()
        raise Exception(message)


class ETLBaselineMSSQL:
    def __init__(self, window):
        window.title("Eureka® Monitor de progresso")
        self.start_time = time.time()

        self.qp_label = tk.Label(window, text="")
        self.qp_label.pack(pady=15)
        self.status_label = tk.Label(window, text="")
        self.status_label.pack(pady=20)

        self.progress = ttk.Progressbar(window, orient="horizontal", length="300", mode="determinate")
        self.progress.pack(pady=30)

    def start_etl(self):
        excel_filepath = None
        try:
            delay = 0.4
            codigo_qp = get_env_var_windows('QP_BASELINE')
            codigo_qp_formatado = codigo_qp.replace('QP-E', '').zfill(6)
            validar_codigo_qp = qp_validate(codigo_qp)
            self.qp_label.config(text=f"Salvar baseline da {codigo_qp} no TOTVS")
            self.status_label.config(text=f"Iniciando processo..."
                                          f"{codigo_qp_formatado.replace('0', '')} no TOTVS")
            time.sleep(delay)
            self.update_progress(5)

            if validar_codigo_qp:
                excel_filepath = get_excel_filepath(codigo_qp)
                self.status_label.config(text="Extraindo dados...")
                time.sleep(delay)
                self.update_progress(10)
                dataframe_original = pd.read_excel(excel_filepath, sheet_name="PROJETO", engine="openpyxl")

                self.status_label.config(text="Transformando dados...")
                time.sleep(delay)
                self.update_progress(15)
                dataframe_baseline = dataframe_original.copy()
                colunas_para_remover = ['ID', 'VISÃO GERAL', 'LINK', 'OBSERVAÇÕES', 'PEÇA\nREPOSIÇÃO', 'TOTVs', 'QTDE',
                                        '%', '0', '1', '2', '3', '4', '5', '6', '7', '8', '9']

                dataframe_baseline = dataframe_baseline.drop(columns=colunas_para_remover)
                dataframe_baseline = dataframe_baseline.fillna('')

                colunas_qtde = ['QTDE\nBL', 'QTDE PROJ.', 'QTDE\nTOTAL']
                dataframe_baseline[colunas_qtde] = dataframe_baseline[colunas_qtde].replace('', 0)

                dataframe_baseline.insert(0, 'QP', codigo_qp_formatado)

                baseline_exist = verify_if_baseline_exists(codigo_qp_formatado)
                self.update_progress(20)
                if baseline_exist:
                    baseline_deleted = delete_if_baseline_exists(codigo_qp_formatado)
                    self.update_progress(25)
                    if not baseline_deleted:
                        return

                self.status_label.config(text="Carregando dados...")
                time.sleep(delay)
                insert_baseline(self, dataframe_baseline)

                # delete_file(excel_filepath)
                end_time = time.time()
                elapsed = end_time - self.start_time
                self.status_label.config(text=f"✔️ Processo finalizado com sucesso!\n\n{elapsed:.3f} segundos"
                                              f"\n\n🦾🤖 EUREKA®")
                self.update_progress(100)
            else:
                message = 'Código da QP inválido! Por favor corrigir o código da QP no arquivo da baseline.'
                raise Exception(message)
        except Exception as ex:
            self.status_label.config(text='❌ Processo cancelado!')
            self.update_progress(100)
            exibir_mensagem('Eureka® Erro de processamento', {ex}, 'warning')
            return None
        finally:
            delete_file(excel_filepath)

    def start_task(self):
        thread = threading.Thread(target=self.start_etl)
        thread.start()

    def update_progress(self, value):
        self.progress['value'] = value
        root.update_idletasks()


if __name__ == '__main__':
    root = tk.Tk()
    username, password, database, server = setup_mssql()
    driver = '{SQL Server}'
    etlBaseline = ETLBaselineMSSQL(root)
    etlBaseline.start_task()

    root.attributes('-topmost', True)
    root.geometry("400x280")
    root.mainloop()
