import pandas as pd
import sys
from quickstart import BASE_ALTERACAO, BASE_HOMOLOGADO


class DataHandler:
    """Essa classe recebe duas planilhas do google e retornam dois dataframes bases"""
    def import_data(self):
        # Lista das colunas necessárias para a tabela ALTERACAO
        colunas_alteracao = [
            'MATRICULA', 'Escolha uma opção', 'DIA DA ALTERAÇÃO', 'ENTRADA', 'SAÍDA PARA ALMOÇO',
            'ENTRADA APÓS O ALMOÇO', 'SAÍDA', 'ENTRADA NOTURNA', 'SAÍDA NOTURNA',
            'DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR', 'QUANTIDADE DE HORAS:'
        ]

        # Criar DataFrame diretamente com as colunas especificadas
        self.df = pd.DataFrame(BASE_ALTERACAO, columns=BASE_ALTERACAO[0]).drop(0, axis=0)
        self.df = self.df[colunas_alteracao]

        # Lista das colunas necessárias para a tabela HOMOLOGADO
        colunas_homologado = [
            "Mat", "Nome", "Sala 2ª à 6ª feira", "PPM 2ª feira", "HTPC 3ª feira",
            "AP", "HL", "HLE", "Horas Compensadas", "C.H. Semanal"
        ]

        # Criar DataFrame diretamente com as colunas especificadas
        self.df2 = pd.DataFrame(BASE_HOMOLOGADO, columns=BASE_HOMOLOGADO[0]).drop(0, axis=0)
        self.df2 = self.df2[colunas_homologado]

        return self.df, self.df2

    def adiciona_filtros_internos(self):
        # Filtrar colunas relevantes da tabela ALTERACAO
        self.alteracao = self.df[
            ["MATRICULA", 'Escolha uma opção', 'DIA DA ALTERAÇÃO', 'ENTRADA', 'SAÍDA PARA ALMOÇO',
             'ENTRADA APÓS O ALMOÇO', 'SAÍDA', 'ENTRADA NOTURNA', 'SAÍDA NOTURNA']
        ]
        
        # Renomear as colunas para nomes mais claros
        self.alteracao.columns = [
            "Mat", "OPCAO", "Dia da Alteração", "Entrada 01", "Saída Intervalo 01",
            "Entrada Intervalo 01", "Saída 01", "Entrada 02", "Saída 02"
        ]
        # Preencher valores ausentes com '--'
        self.alteracao=self.alteracao.fillna('--')

        # Filtrar colunas relevantes da tabela HOMOLOGADO
        self.homologado = self.df2[
            ["Mat", "Nome", "Sala 2ª à 6ª feira", "PPM 2ª feira", "HTPC 3ª feira",
             "AP", "HL", "HLE", "Horas Compensadas", "C.H. Semanal"]
        ]
        # Preencher valores ausentes com '--'
        self.homologado = self.homologado.fillna('--')

        return self.alteracao, self.homologado

    def get_filtered_matriculas(self):
        # Filtrar matrículas que têm OPÇÃO como 'ALTERAÇÃO DE HORÁRIO' e estão presentes na tabela HOMOLOGADO
        return self.alteracao.loc[
            (self.alteracao['OPCAO'] == 'ALTERAÇÃO DE HORÁRIO') &
            (self.alteracao['Mat'].isin(self.homologado['Mat'])), 'Mat'
        ].unique()

    def get_alteracoes_filtradas(self, matricula, mes_pagamento):
        # Filtrar as alterações para uma matrícula específica e um mês de pagamento específico
        alteracoes_filtradas = self.alteracao[self.alteracao['Mat'] == matricula]
        alteracoes_filtradas = alteracoes_filtradas[alteracoes_filtradas['OPCAO'] == 'ALTERAÇÃO DE HORÁRIO']
        colunas_desejadas = [
            'Dia da Alteração', 'Entrada 01', 'Saída Intervalo 01',
            'Entrada Intervalo 01', 'Saída 01', 'Entrada 02', 'Saída 02'
        ]
        alteracoes_filtradas['Dia da Alteração'] = pd.to_datetime(alteracoes_filtradas['Dia da Alteração'], format='%d/%m/%Y')
        alteracoes_filtradas = alteracoes_filtradas[alteracoes_filtradas['Dia da Alteração'].dt.month == mes_pagamento]
        alteracoes_filtradas = alteracoes_filtradas.loc[:, colunas_desejadas]
        return alteracoes_filtradas

    def carga_suplementar(self, mes_pagamento):
        # Aumentar o limite de recursão para lidar com grandes conjuntos de dados
        sys.setrecursionlimit(10 ** 6)

        # Filtrar colunas relevantes da tabela de CARGA SUPLEMENTAR
        self.carga = self.df[['MATRICULA', 'Escolha uma opção', 'DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR', 'QUANTIDADE DE HORAS:']]
        self.carga = self.carga.rename(columns={'MATRICULA':'Mat',
                                                'Escolha uma opção':'Escolha uma opção', 
                                                'DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR':'DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR',
                                                'QUANTIDADE DE HORAS:':'QUANTIDADE DE HORAS:'})

        # Filtrar registros com 'Escolha uma opção' igual a 'CARGA SUPLEMENTAR'
        self.carga = self.carga[self.carga['Escolha uma opção'] == 'CARGA SUPLEMENTAR']

        # Remover coluna 'Escolha uma opção'
        self.carga= self.carga.drop('Escolha uma opção', axis=1)

        # Converter MATRICULA para tipo numérico e QUANTIDADE DE HORAS: para tipo float
        self.carga['Mat'] = pd.to_numeric(self.carga['Mat'], errors='ignore', downcast='integer')
        self.carga['QUANTIDADE DE HORAS:'] = self.carga['QUANTIDADE DE HORAS:'].str.replace(',', '.')
        self.carga['QUANTIDADE DE HORAS:'] = pd.to_numeric(self.carga['QUANTIDADE DE HORAS:'], downcast='float')

        # Converter 'DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR' para formato de data
        self.carga['DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR'] = pd.to_datetime(
            self.carga['DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR'],
            format='%d/%m/%Y',
            errors='coerce'
        )

        # Remover registros com valores nulos em 'DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR'
        self.carga = self.carga.dropna(subset=['DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR'])

        # Ordenar os registros por data
        self.carga = self.carga.sort_values('DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR')

        # Filtrar registros para um mês de pagamento específico
        self.carga = self.carga[(self.carga['DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR'].dt.month == mes_pagamento)]

        # Calcular a soma das horas para cada MATRICULA e obter os dias de realização da carga
        self.soma_horas = self.carga.groupby('Mat')['QUANTIDADE DE HORAS:'].sum().reset_index()
        self.soma_horas['DIA'] = self.carga.groupby('Mat')['DIA DE REALIZAÇÃO DA CARGA SUPLEMENTAR']\
        .agg(lambda x: x.dt.day.tolist()).tolist()

        # Filtrar colunas relevantes da tabela HOMOLOGADO
        self.nomes = self.df2[["Mat", "Nome"]]
        
        # Converte 'Mat' series em self.nomes para tipo numérico (int ou float)
        self.nomes['Mat'] = self.nomes['Mat'].astype(int)
                
        # Merge das tabelas usando a coluna 'Mat' como chave
        self.carga_mes = self.soma_horas.merge(self.nomes, on='Mat')

        # Converter a coluna 'MATRICULA' para tipo inteiro
        self.carga_mes['Mat'] = self.carga_mes['Mat'].astype(int)

        # Renomear as colunas
        self.carga_mes = self.carga_mes.rename(
            columns={'DIA': 'Dias de Carga', 'QUANTIDADE DE HORAS:': 'Carga Total'}
            )[['Mat', 'Nome', 'Dias de Carga', 'Carga Total']]
        
        return self.carga_mes
    

