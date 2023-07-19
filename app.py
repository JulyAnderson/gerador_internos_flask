import sys
sys.setrecursionlimit(1500)

from flask import Flask, render_template, request, url_for
from import_df import DataHandler
from interno import InternoGenerator
from cargas_suplementar import CargaGenerator
import os
import webbrowser

# Obtenha o diretório atual do script
base_dir = os.path.dirname(os.path.abspath(__file__))

# Define o caminho absoluto para os diretórios de templates e arquivos estáticos
template_dir = os.path.join(base_dir, 'templates')
static_dir = os.path.join(base_dir, 'static')

# Configura o Flask para usar os caminhos definidos
app = Flask(__name__, template_folder=template_dir, static_folder=static_dir)


# Define o caminho para a pasta temporária usada pelo Flask- Modificações para o pyinstaller
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.config['EXPLAIN_TEMPLATE_LOADING'] = True
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0

dict_mes = {
        1: 'janeiro',
        2: 'fevereiro',
        3: 'março',
        4: 'abril',
        5: 'maio',
        6: 'junho',
        7: 'julho',
        8: 'agosto',
        9: 'setembro',
        10: 'outubro',
        11: 'novembro',
        12: 'dezembro'
    }

@app.route('/')
def index():
    return render_template('index.html', dict_mes=dict_mes, mensagem_sucesso='')

@app.route('/gerar_interno', methods=['POST'])
def gerar_internos():
    mes_pagamento = int(request.form['mes_var'])
    numero_interno = int(request.form['numero_var'])
    dia_envio = int(request.form['dia_var'])
    responsavel_assinatura = request.form['responsavel_var']

    data_handler = DataHandler()
    data_handler.import_data()
    interno_generator = InternoGenerator(data_handler, numero_interno, dia_envio, mes_pagamento, responsavel_assinatura)

    # Verificar se a pasta 'internos' existe, caso contrário, criar
    if not os.path.exists('interno'):
        os.makedirs('interno')

    # Obter as matrículas filtradas
    data_handler.adiciona_filtros_internos()
    matriculas_filtradas = data_handler.get_filtered_matriculas()

    # Gerar documentos internos para cada matrícula filtrada
    for matricula in matriculas_filtradas:
        interno_generator.gerar_interno(matricula)

 
    mensagem = "Obrigado por usar o Gerador de Pagamento. Seus Internos foram Gerados com Sucesso!"

    return render_template('index.html', mensagem_sucesso = mensagem, dict_mes=dict_mes)

@app.route('/gerar_carga', methods=['POST'])
def gerar_cargas():
    mes_pagamento = int(request.form['mes_var'])     

    data_handler = DataHandler()
    data_handler.import_data()  

    carga_generator = CargaGenerator(data_handler, mes_pagamento)

    # Verificar se a pasta 'carga' existe, caso contrário, criar
    if not os.path.exists('carga'):
        os.makedirs('carga')

    carga = data_handler.carga_suplementar(mes_pagamento)   

    carga_generator.criar_carga(carga)
 
    mensagem = "Obrigado por usar o Gerador de Pagamento. As Cargas Suplementares foram Geradas com Sucesso!"

    return render_template('index.html', mensagem_sucesso = mensagem, dict_mes=dict_mes)


@app.route('/processar', methods=['POST'])
def processar_formulario():
   
    if request.form['acao'] == 'gerar_internos':
        gerar_internos()
        mensagem_sucesso = "Internos gerados com sucesso!"
    elif request.form['acao'] == 'gerar_cargas':
        gerar_cargas()
        mensagem_sucesso = "Carga suplementar gerada com sucesso!"
    
    return render_template('index.html', dict_mes=dict_mes, mensagem_sucesso=mensagem_sucesso)
  

#webbrowser.open('http://127.0.0.1:5000')

# if __name__ == '__main__':
#     app.run(host='127.0.0.1', port=5000)

app.run(debug=True)