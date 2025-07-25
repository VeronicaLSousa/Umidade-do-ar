import tkinter as tk
from tkinter import messagebox
import requests
from datetime import datetime
from planilha_excel import salvar_excel

def captar_dados():
    try:
        url = "https://api.open-meteo.com/v1/forecast?latitude=-23.55&longitude=-46.63&current_weather=true&hourly=relativehumidity_2m"
        resposta = requests.get(url)
        dados = resposta.json()

        temperatura = dados["current_weather"]["temperature"]
        horario_api = dados["current_weather"]["time"]  # Ex: "2025-06-18T13:00"

        # Arredonda o horário para o mesmo padrão de "hourly"
        horario_dt = datetime.strptime(horario_api, "%Y-%m-%dT%H:%M")
        horario_formatado = horario_dt.strftime("%Y-%m-%dT%H:00")

        horarios = dados["hourly"]["time"]
        umidades = dados["hourly"]["relativehumidity_2m"]

        if horario_formatado in horarios:
            indice = horarios.index(horario_formatado)
            umidade = umidades[indice]
        else:
            umidade = "N/A"

        data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        salvar_excel(data_hora, f"{temperatura} °C", f"{umidade} %")

        return temperatura, umidade

    except Exception as e:
        print("Erro:", e)
        return None, None

def buscar_previsao():
    temperatura, umidade = captar_dados()
    if temperatura is not None:
        messagebox.showinfo("Sucesso", f"Dados captados:\nTemperatura: {temperatura} °C\nUmidade: {umidade} %")
    else:
        messagebox.showerror("Erro", "Não foi possível obter os dados.")

janela = tk.Tk()
janela.title("Captador de Temperatura - São Paulo")
janela.geometry("360x200")

titulo = tk.Label(janela, text="Clique para captar dados do clima em SP", font=("Arial", 12))
titulo.pack(pady=10)

botao = tk.Button(janela, text="Buscar Previsão", command=buscar_previsao)
botao.pack(pady=20)

janela.mainloop()


