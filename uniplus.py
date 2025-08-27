
# Autor: Matheus Jonathan Lima Simoes
# Projeto: Automação Uniplus - Impressao de Pedidos
# Data de criaçao: 23/08/2025
# Última atualizaçao: 27/08/2025
# Descriçao: Script para automação de impressão de pedidos no Uniplus.


import subprocess
import time
import requests
import os
import threading
import tkinter as tk
from tkinter import scrolledtext
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# CONFIGURAÇÕES 
PORTA_DEBUG = 9222
BAT_CAMINHO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "abrir_uniplus.bat")
SENHA = "07051995"
LOGIN = "47"
EMPRESA = "UNIAO DA CAIRO COMERCIO DE CEREAIS E TEMPEROS LTDA"
TEMPO_ESPERA = 10
INTERVALO_MINUTOS = 1


# VARIÁVEIS DE CONTROLE
executando = False
thread_execucao = None

# INTERFACE COM TKINTER 
def escrever_log(mensagem):
    terminal.configure(state='normal')
    terminal.insert(tk.END, mensagem + '\n')
    terminal.see(tk.END)
    terminal.configure(state='disabled')

def iniciar_chrome_com_bat():
    escrever_log("🚀 Abrindo Chrome via .bat")
    subprocess.Popen(BAT_CAMINHO, shell=True)

    escrever_log("⏳ Aguardando Chrome responder na porta 9222")
    for _ in range(30):
        try:
            r = requests.get(f"http://localhost:{PORTA_DEBUG}/json")
            if r.status_code == 200:
                escrever_log("✅ Chrome conectado com sucesso!")
                return
        except:
            pass
        time.sleep(1)
    raise Exception("❌ Falha ao conectar com o Chrome (porta 9222 não respondeu)")

def iniciar_driver():
    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", f"localhost:{PORTA_DEBUG}")
    escrever_log("🟡 Conectando ao Chrome já iniciado com perfil salvo")
    return webdriver.Chrome(options=chrome_options)

def tratar_erros_popup(driver):
    try:
        for el in driver.find_elements(By.CSS_SELECTOR, "a.toast-close-button"):
            driver.execute_script("arguments[0].click();", el)
            escrever_log("❌ Toast de erro fechado.")
        for el in driver.find_elements(By.ID, "fechar"):
            driver.execute_script("arguments[0].click();", el)
            escrever_log("❌ Modal principal de erro fechado.")
        for el in driver.find_elements(By.ID, "close_adm_cart"):
            driver.execute_script("arguments[0].click();", el)
            escrever_log("❌ Modal secundário (cart) fechado.")
    except:
        pass

def realizar_login(driver):
    escrever_log("🔐 Verificando status de login")
    try:
        senha = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "current-password")))
        try:
            usuario = driver.find_element(By.ID, "username")
            escrever_log("✏️ Usuário e senha detectados. Preenchendo os dois")
            usuario.clear()
            usuario.send_keys(LOGIN)
        except:
            escrever_log("🔐 Apenas senha requerida (usuário já salvo)")

        senha.clear()
        senha.send_keys(SENHA)
        time.sleep(1)
        driver.find_element(By.XPATH, "//button[contains(text(),'ENTRAR')]").click()
        escrever_log("✅ Login realizado com sucesso!")
        time.sleep(3)
    except:
        escrever_log("✅ Já está logado (nenhum campo de login detectado)")

def ir_para_pedidos(driver):
    escrever_log("📂 Acessando VENDAS → Pedidos de faturamento")
    try:
        WebDriverWait(driver, TEMPO_ESPERA).until(EC.element_to_be_clickable((By.XPATH, "//*[normalize-space(text())='Vendas']"))).click()
        time.sleep(2)
        WebDriverWait(driver, TEMPO_ESPERA).until(EC.element_to_be_clickable((By.XPATH, "//*[normalize-space(text())='Pedidos de faturamento']"))).click()
        time.sleep(3)
        escrever_log("✅ Tela de pedidos acessada com sucesso!")
    except Exception as e:
        escrever_log(f"❌ Erro ao acessar pedidos: {e}")
        tratar_erros_popup(driver)

def aplicar_filtro_pre_pedido(driver):
    escrever_log("🔎 Aplicando filtro de pré-pedido")
    try:
        driver.find_element(By.ID, "filtrarGrid").click()
        time.sleep(2)
        driver.find_element(By.XPATH, "//label[contains(text(),'Situação')]/..//button").click()
        driver.find_element(By.XPATH, "//div[contains(@class, 'wj-listbox-item') and text()='Pré-pedido']").click()
        campo_data = driver.find_element(By.ID, "FiltroFaturamentoWrapper_dataInicial")
        campo_data.click()
        campo_data.send_keys(Keys.DELETE)
        driver.find_element(By.ID, "aplicar").click()
        time.sleep(2)
        escrever_log("✅ Filtro aplicado!")
    except Exception as e:
        escrever_log(f"⚠️ Erro ao aplicar filtro: {e}")
        tratar_erros_popup(driver)

def imprimir_pedidos(driver):
    escrever_log("🖨️ Iniciando impressão!")
    try:
        linhas = driver.find_elements(By.XPATH, "//div[@role='row' and @aria-selected]")
        escrever_log(f"🔎 {len(linhas)} pedidos encontrados.")
        for index, linha in enumerate(linhas):
            colunas = linha.find_elements(By.XPATH, ".//div[@role='gridcell']")
            if len(colunas) < 10:
                continue
            if colunas[-1].text.strip() == "":
                escrever_log(f"📄 Pedido {index+1} sem data. Imprimindo...")
                driver.execute_script("arguments[0].click();", colunas[0])
                driver.find_element(By.ID, "_imprimir").click()
                time.sleep(2)

                # Troca para a janela de impressão
                driver.switch_to.window(driver.window_handles[-1])

                try:
                    # Aguarda até 10 segundos o botão "Imprimir" estar clicável
                    for tentativa in range(20):
                        try:
                            app = driver.find_element(By.TAG_NAME, "print-preview-app")
                            shadow1 = driver.execute_script("return arguments[0].shadowRoot", app)

                            button_strip = shadow1.find_element(By.CSS_SELECTOR, "print-preview-button-strip")
                            shadow2 = driver.execute_script("return arguments[0].shadowRoot", button_strip)

                            botao = shadow2.find_element(By.CSS_SELECTOR, "cr-button.action-button:not([disabled])")
                            botao.click()
                            escrever_log("✅ Impresso com sucesso.")
                            break
                        except Exception as e:
                            time.sleep(0.5)
                    else:
                        escrever_log("⚠️ Botão de impressão não encontrado ou ainda desativado.")
                except Exception as e:
                    escrever_log(f"⚠️ Erro ao acessar o botão de impressão: {e}")

                driver.switch_to.window(driver.window_handles[0])
    except Exception as e:
        escrever_log(f"⚠️ Erro ao imprimir: {e}")


def executar_fluxo():
    global executando
    try:
        iniciar_chrome_com_bat()
        driver = iniciar_driver()
        realizar_login(driver)
        time.sleep(5)
        tratar_erros_popup(driver)
        ir_para_pedidos(driver)

        while executando:
            aplicar_filtro_pre_pedido(driver)
            imprimir_pedidos(driver)

            escrever_log("🔄 Aguardando nova atualização (15 segundos)...\n")
            for _ in range(15):
                if not executando:
                    escrever_log("⏹️ Execução pausada.")
                    return
                time.sleep(1)

            
            time.sleep(3)  # Espera após o refresh

    except Exception as e:
        escrever_log(f"❌ Erro no fluxo principal: {e}")


def iniciar():
    global executando, thread_execucao
    if not executando:
        executando = True
        thread_execucao = threading.Thread(target=executar_fluxo)
        thread_execucao.start()
        escrever_log("▶️ Execução iniciada.")

def parar():
    global executando
    executando = False
    escrever_log("⏸️ Execução pausada")

def fechar():
    global executando
    executando = False
    root.destroy()

# GUI
root = tk.Tk()
root.title("Automação Uniplus")
root.geometry("700x500")

frame_botoes = tk.Frame(root)
frame_botoes.pack(pady=10)

btn_iniciar = tk.Button(frame_botoes, text="▶️ Iniciar", width=15, command=iniciar)
btn_iniciar.grid(row=0, column=0, padx=5)

btn_parar = tk.Button(frame_botoes, text="⏸️ Pausar", width=15, command=parar)
btn_parar.grid(row=0, column=1, padx=5)

btn_fechar = tk.Button(frame_botoes, text="❌ Fechar", width=15, command=fechar)
btn_fechar.grid(row=0, column=2, padx=5)

terminal = scrolledtext.ScrolledText(root, height=25, state='disabled', bg="black", fg="lime", font=("Consolas", 10))
terminal.pack(fill='both', expand=True, padx=10, pady=10)

root.mainloop()
