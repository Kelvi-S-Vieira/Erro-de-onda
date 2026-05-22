import pandas as pd
import numpy as np
from datetime import datetime, timedelta

editorLocal = pd.read_csv("C:\\Users\\884684\\Painel_PCO\\Editor_Local_Separacao\\Editor_Local_Separacao.csv", low_memory=False)
enderecoMarcado = pd.read_csv("C:\\Users\\884684\\Painel_PCO\\Endereco_Marcado\\Endereco_Marcado_Contagem.csv", low_memory=False)
estoqueEndereco = pd.read_csv("C:\\Users\\884684\\Painel_PCO\\Estoque_por_Endereco\\Estoque_por_Endereco.csv", low_memory=False)
frequenciaAbastecimento = pd.read_excel("C:\\Users\\884684\\Painel_PCO\\Frequência\\Simulador-de-Frequência-Abastecimento-Semana.xlsx", sheet_name="Tabela Datas (Query)")
mapaVolumes = pd.read_csv("C:\\Users\\884684\\Painel_PCO\\Mapa_Volumes\\Mapa_de_Volumes.csv", low_memory=False)

map_classe_setor = {
    "CLSETK": "K",
    "CLSETEV": "EV",
    "DEFAULT": "INCORRETO",
    "CLSETA": "A",
    "CLSETG": "G",
    "CLSETJ": "J",
    "CLSETAJ": "AJ",
    "CLSETC": "C",
    "CLSETY": "Y",
    "CLSETEL": "EL",
    "CLSETEE": "EE",
    "CLSETS": "S",
    "CLSETTB": "AJ",
    "CLSETUN": "XX",
    "CLSETE": "E",
    "CLSETU": "U",
    "CLSETSS": "SS",
    "CLSETAG": "AJ",
    "CLSETV": "V",
    "CLSETL": "L",
    "CLSETEF": "EF",
    "CLSETEB": "G",
    "CLSETYY": "YY"
}

map_erro_responsavel = {
    "ITEM SEM ESTOQUE DISPONIVEL": "C.E",
    "MARCADO MANUALMENTE": "C.E",
    "ENDERECO BLOQUEADO": "C.E",
    "CASEPACK": "C.E",
    "ITEM SEM ESTOQUE PARCIAL": "C.E",
    "ENDERECO FORA DE SERVICO   ": "C.E",
    "IN-TRANSIT - CE": "C.E",
    "CLASSE INCORRETA": "C.E",
    "FLOW THROUNG": "FLOW THROUNG",
    "RTV": "O.P",
    "IN-TRANSIT": "O.P",
    "UNLOCATEDLOC": "O.P",
    "ARMAZENAGEM": "O.P",
    "DEVOL-ESTQ": "O.P",
    "ESTOQUE INELEGIVEL": "O.P",
    "ENDEREÇAMENTO INCORRETO": "PCP",
    "ITEM SEM LOCAL DE SEPARACAO": "PCP",
    "RECEBIMENTO": "REC",
    "ERRO SISTEMICO": "T.I",
    "RESOLVIDO": "OK",
    "ENDEREÇADO": "OK",
    "REARMAZENAR - CLASSE INCORRETA": "O.P"
}

def projecaoFillRate(editorLocal, enderecoMarcado, estoqueEndereco, frequenciaAbastecimento, mapaVolumes):

    print("Iniciando projeção de fill rate...")

    data_alvo = (datetime.now() + timedelta(days=2)).date()

    frequenciaAbastecimento['EXPEDIÇÃO'] = pd.to_datetime(
        frequenciaAbastecimento['EXPEDIÇÃO'],
        errors='coerce'
    ).dt.date

    lojas_na_data = frequenciaAbastecimento.loc[
        frequenciaAbastecimento['EXPEDIÇÃO'] == data_alvo,
        'COD'
    ].unique()

    mapaVolumes['EXPEDE_D2'] = mapaVolumes['FILIAL'].isin(lojas_na_data)

    mapaVolumesFiltrado = mapaVolumes.loc[mapaVolumes['EXPEDE_D2']].copy()

    # -------------------------------------------------
    # CRIANDO CHAVE ÚNICA
    # -------------------------------------------------

    mapaVolumesFiltrado['CHAVE'] = (
        mapaVolumesFiltrado['STOCK_ORDER'].astype(str) + "_" +
        mapaVolumesFiltrado['FILIAL'].astype(str) + "_" +
        mapaVolumesFiltrado['ARTIGO'].astype(str)
    )

    # -------------------------------------------------
    # SOMASES ESTOQUE (TOTAL POR ITEM)
    # -------------------------------------------------

    filtro_endereco = ~estoqueEndereco['TIP_END'].isin(
    ['SDR','MISC','PT BINS SEP1','PT BINS 1 ARM','INTRANSIT']
    )

    filtro_intransit = (
        (estoqueEndereco['TIP_END'] == 'INTRANSIT') &
        (estoqueEndereco['DATA_RECEB'].notna())
    )

    filtro_total = filtro_endereco | filtro_intransit

    soma_artigos = (
        estoqueEndereco[filtro_total]
        .groupby('ITEM_ID')['VOLUMES']
        .sum()
    )

    mapaVolumesFiltrado['POSSUI_ESTOQUE'] = (
        mapaVolumesFiltrado['ARTIGO']
        .map(soma_artigos)
        .fillna(0)
        .apply(lambda x: 'Sim' if x > 0 else 'Sem Estoque')
    )

    rtv_unlocated = (
    estoqueEndereco[
        estoqueEndereco['ENDERECO'].isin(['RTV', 'UNLOCATEDLOC'])
    ]
    .groupby('ITEM_ID')['ENDERECO']
    .first()
    )

    mapaVolumesFiltrado['LOCAL_ESTOQUE_STATUS'] = (
    mapaVolumesFiltrado['ARTIGO']
    .map(rtv_unlocated)
    .fillna('Nao')
    )

    devol_arm_sep = (
        estoqueEndereco[                        
                        estoqueEndereco['ENDERECO'].isin(['DEVOL-ARM', 'DEVOL-SEP', 'DEVOL-ARMK', 'DEVOLUCAOK'])]
    .groupby('ITEM_ID')['ENDERECO']
    .first()
    )

    mapaVolumesFiltrado['DEVOLUCAO_STATUS'] = (
    mapaVolumesFiltrado['ARTIGO']
    .map(devol_arm_sep)
    .fillna('Nao')
    )

    estoque_real = (
        estoqueEndereco[
            estoqueEndereco['TIP_END'] != 'RDR'
        ]
        .groupby('ITEM_ID')['VOLUMES']
        .sum()
    )

    mapaVolumesFiltrado['ESTOQUE_REAL'] = (
        mapaVolumesFiltrado['ARTIGO']
        .map(estoque_real)
        .fillna(0)
    )

    mapaVolumesFiltrado['ARMAZENAR_ITEM'] = (
        np.where(
            mapaVolumesFiltrado['ARTIGO'].isin(
                estoqueEndereco.loc[estoqueEndereco['TIP_END'] == 'RDR', 'ITEM_ID']
            )
            & (mapaVolumesFiltrado['VOLUME'] > mapaVolumesFiltrado['ESTOQUE_REAL']),
         'Sim', 'Nao'
        )
    )

    filtro_armazenar = estoqueEndereco[
    estoqueEndereco['TIP_END'].isin([
        'PT PALETE 3 ARM',
        'PT PALETE 1 ARM',
        'PT PALETE 2 ARM'
    ])
    ].copy()

    filtro_armazenar['SETOR_CORRETO'] = (
    filtro_armazenar['CLASSE']
    .map(map_classe_setor)
    .fillna('INCORRETO')
    )

    filtro_armazenar['REARMAZENAR'] = (
    filtro_armazenar['SETOR_PLAN_ESTOQUE'] != filtro_armazenar['SETOR_CORRETO']
    )

    rearmazenar_item = (
    filtro_armazenar
    .groupby('ITEM_ID')['REARMAZENAR']
    .any()
    )

    mapaVolumesFiltrado['REARMAZENAR'] = (
    mapaVolumesFiltrado['ARTIGO']
    .map(rearmazenar_item)
    .fillna(False)
    )

    # -------------------------------------------------
    # LOOKUP LOCAL
    # -------------------------------------------------

    editorLocal_lookup = (
        editorLocal[['ITEM','LOCAL']]
        .drop_duplicates('ITEM')
    )

    mapaVolumesFiltrado = mapaVolumesFiltrado.merge(
        editorLocal_lookup,
        how='left',
        left_on='ARTIGO',
        right_on='ITEM'
    ).drop(columns=['ITEM']).fillna({'LOCAL': 'Enderecar'})

    mapaVolumesFiltrado['LOCAL'] = np.where(
        mapaVolumesFiltrado['SETOR'] == 'K',
        'Encabidados',
        mapaVolumesFiltrado['LOCAL']
    )

    # -------------------------------------------------
    # CLASSE DO SETOR (AGREGADA)
    # -------------------------------------------------

    estoqueEndereco['CLASSE_SETOR'] = estoqueEndereco['CLASSE'].map(map_classe_setor).fillna('INCORRETO')

    classe_lookup = (
        estoqueEndereco[['ITEM_ID','CLASSE_SETOR']]
        .drop_duplicates('ITEM_ID')
    )

    mapaVolumesFiltrado = mapaVolumesFiltrado.merge(
        classe_lookup,
        how='left',
        left_on='ARTIGO',
        right_on='ITEM_ID'
    ).drop(columns=['ITEM_ID']).fillna({'CLASSE_SETOR': 'INCORRETO'})

    mapaVolumesFiltrado['CLASSE_CORRETA'] = np.where(
        mapaVolumesFiltrado['CLASSE_SETOR'] == mapaVolumesFiltrado['SETOR'],
        'Sim',
        'Nao'
    )



    # -------------------------------------------------
    # STATUS LOCAL
    # -------------------------------------------------

    status_lookup = (
        editorLocal[['ITEM','STATUS_LOCAL']]
        .drop_duplicates('ITEM')
    )

    mapaVolumesFiltrado = mapaVolumesFiltrado.merge(
        status_lookup,
        how='left',
        left_on='ARTIGO',
        right_on='ITEM'
    ).drop(columns=['ITEM'])

    mapaVolumesFiltrado['ARTIGO'] = mapaVolumesFiltrado['ARTIGO'].astype(str)
    enderecoMarcado['ITEM_ID'] = enderecoMarcado['ITEM_ID'].astype(str)

    mapaVolumesFiltrado = mapaVolumesFiltrado.merge(
        enderecoMarcado[['ITEM_ID','P_TIPO']],
        how='left',
        left_on='ARTIGO',
        right_on='ITEM_ID'
    ).drop(columns=['ITEM_ID'])

    mapaVolumesFiltrado['MARCADO PARA CONTAGEM'] = np.where(
        mapaVolumesFiltrado['P_TIPO'] == 'MM',
        'Sim',
        'Nao'
    )

        # -------------------------------------------------
    # DEFINIÇÃO DO ERRO REAL
    # -------------------------------------------------

    mapaVolumesFiltrado['ERRO_REAL'] = np.select(

        [
            # 🔴 Sem estoque
            mapaVolumesFiltrado['POSSUI_ESTOQUE'] == 'Sem Estoque',

            # 🔴 Marcado manualmente
            mapaVolumesFiltrado['MARCADO PARA CONTAGEM'] == 'Sim',

            # 🔴 Endereço bloqueado
            mapaVolumesFiltrado['STATUS_LOCAL'] == 'OUT-SERVICE',

            # 🟡 Classe incorreta (rearmazenar)
            mapaVolumesFiltrado['REARMAZENAR'] == True,

            # 🟡 Item precisa armazenar
            mapaVolumesFiltrado['ARMAZENAR_ITEM'] == 'Sim',

            # 🔵 RTV / UNLOCATED
            mapaVolumesFiltrado['LOCAL_ESTOQUE_STATUS'] == 'RTV',
            mapaVolumesFiltrado['LOCAL_ESTOQUE_STATUS'] == 'UNLOCATEDLOC',

            # 🔵 Devolução
            mapaVolumesFiltrado['DEVOLUCAO_STATUS'] != 'Nao',

            # 🔵 Sem local
            mapaVolumesFiltrado['LOCAL'] == 'Enderecar'
        ],

        [
            'ITEM SEM ESTOQUE DISPONIVEL',
            'MARCADO MANUALMENTE',
            'ENDERECO FORA DE SERVICO',
            'REARMAZENAR - CLASSE INCORRETA',
            'ARMAZENAR',
            'RTV',
            'UNLOCATEDLOC',
            'DEVOL-ESTQ',
            'ITEM SEM LOCAL DE SEPARACAO'
        ],

        default='Ok'
    )

    mapaVolumesFiltrado['AREA_RESPONSAVEL'] = (
    mapaVolumesFiltrado['ERRO_REAL']
    .map(map_erro_responsavel)
    .fillna('NAO CLASSIFICADO')
    )

    resumo_erros = (
    mapaVolumesFiltrado
    .groupby(['ERRO_REAL','AREA_RESPONSAVEL'])
    .size()
    .reset_index(name='UNIDADES')
    .sort_values('UNIDADES', ascending=False)
    )
    # -------------------------------------------------
    # GARANTIR 1 LINHA POR CHAVE
    # -------------------------------------------------

    mapaVolumesFiltrado = mapaVolumesFiltrado.drop_duplicates('CHAVE')

    mapaVolumesFiltrado.to_csv(
        r"C:\Users\884684\Desktop\mapaVolumesFiltrado.csv",
        index=False,
        sep=';'
    )

    return mapaVolumesFiltrado


if __name__ == "__main__":
    projecaoFillRate(
        editorLocal,
        enderecoMarcado,
        estoqueEndereco,
        frequenciaAbastecimento,
        mapaVolumes
    )