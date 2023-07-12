import pandas as pd
import sys

class DataHandler:
    def __init__(self, alteracao_file, homologado_file):
        self.alteracao_file = alteracao_file
        self.homologado_file = homologado_file
        self.alteracao = None
        self.homologado = None

    def import_data(self):
        self.alteracao = pd.read_csv(self.alteracao_file, usecols=["MATRICULA",'Escolha uma opção','DIA DA ALTERAÇÃO',
                                                                    'ENTRADA','SAÍDA PARA ALMOÇO',
                                                                    'ENTRADA APÓS O ALMOÇO','SAÍDA',
                                                                    'ENTRADA NOTURNA','SAÍDA NOTURNA'], sep=',')
        self.alteracao.columns = ["Mat", "OPCAO", "Dia da Alteração", "Entrada 01", "Saída Intervalo 01",
                                  "Entrada Intervalo 01", "Saída 01", "Entrada 02", "Saída 02"]
        self.alteracao.fillna('--', inplace=True)

        self.homologado = pd.read_csv(self.homologado_file, usecols=["Mat", "Nome", "Sala 2ª à 6ª feira",
                                                                      "PPM 2ª feira", "HTPC 3ª feira", "AP", "HL", "HLE",
                                                                      "Horas Compensadas", "C.H. Semanal"], sep=",")
        self.homologado.fillna('--', inplace=True)

    def get_filtered_matriculas(self):
        return self.alteracao.loc[(self.alteracao['OPCAO'] == 'ALTERAÇÃO DE HORÁRIO') &
                                  (self.alteracao['Mat'].isin(self.homologado['Mat'])), 'Mat'].unique()

    def get_alteracoes_filtradas(self, matricula, mes_pagamento):
        alteracoes_filtradas = self.alteracao[self.alteracao['Mat'] == matricula]
        alteracoes_filtradas = alteracoes_filtradas[alteracoes_filtradas['OPCAO'] == 'ALTERAÇÃO DE HORÁRIO']
        colunas_desejadas = [
            'Dia da Alteração',
            'Entrada 01',
            'Saída Intervalo 01',
            'Entrada Intervalo 01',
            'Saída 01',
            'Entrada 02',
            'Saída 02'
        ]
        alteracoes_filtradas['Dia da Alteração'] = pd.to_datetime(alteracoes_filtradas['Dia da Alteração'], format='%d/%m/%Y')
        alteracoes_filtradas = alteracoes_filtradas[alteracoes_filtradas['Dia da Alteração'].dt.month == mes_pagamento]
        alteracoes_filtradas = alteracoes_filtradas.loc[:, colunas_desejadas]
        return alteracoes_filtradas

    def carga_suplementar(self, mes):
        self.carga = pd.read_csv(self.alteracao_file, usecols=['MATRICULA','Escolha uma opção',
                                                                    'DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR',
                                                                    'QUANTIDADE DE HORAS:'], sep=',')
        self.nomes = pd.read_csv(self.homologado_file, usecols= ["Mat", "Nome", sep=","] )
               
        self.carga = self.carga[self.carga['Escolha uma opção'] == 'CARGA SUPLEMENTAR']

        self.carga.drop('Escolha uma opção', axis = 1, inplace=True)

        sys.setrecursionlimit(10**6)

        self.carga['MATRICULA'] = pd.to_numeric(self.carga['MATRICULA'], errors='ignore', downcast='integer')
        self.carga['QUANTIDADE DE HORAS:'] = self.carga['QUANTIDADE DE HORAS:'].str.replace(',', '.')
        self.carga['QUANTIDADE DE HORAS:'] = pd.to_numeric(self.carga['QUANTIDADE DE HORAS:'] , downcast='float') 


        self.carga['DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR'] = pd.to_datetime(self.carga['DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR'], format='%d/%m/%Y', errors ='coerce')
        self.carga = self.carga.dropna(subset=['DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR'])

        self.carga = self.carga.sort_values('DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR')

        self.carga = self.carga[(self.carga['DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR'].dt.month == mes)]

        soma_horas = self.carga.groupby('MATRICULA')['QUANTIDADE DE HORAS:'].sum().reset_index().set_index('MATRICULA')

        soma_horas['DIA'] = self.carga.groupby('MATRICULA')['DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR'].agg(lambda x: x.dt.day.tolist())

        self.nomes.set_index('Matricula', inplace= True)

        carga_mes = soma_horas.join(self.nomes)

        carga_mes = carga_mes.rename(columns={'Nome':'Nome','DIA':'Dias de Carga','QUANTIDADE DE HORAS:':'Carga Total'})[['Nome','Dias de Carga','Carga Total']]

        return carga_mes , mes
