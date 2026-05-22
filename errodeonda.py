import pandas as pd
import numpy as np
from datetime import datetime, timedelta


editorLocal = pd.read_csv("C:\\Users\\melho\\Downloads\\Editor_Local_Separacao (17).csv", low_memory=False)
enderecoMarcado = pd.read_csv("C:\\Users\\melho\\Downloads\\Endereco_Marcado_Contagem (19).csv", low_memory=False)
estoqueEndereco = pd.read_csv("C:\\Users\\melho\\Downloads\\Estoque_por_Endereco (73).csv", low_memory=False)
indicador_operacao = pd.read_excel("C:\\Users\\melho\\Downloads\\Indicador de Operação 22-05.xlsb", sheet_name= "1º - Ind. Operação")
pendencia_embarque = pd.read_csv("C:\\Users\\melho\\Downloads\\Indicador - Pendencia de Embarque (41).csv", low_memory=False)
base_erro_onda = pd.read_excel("C:\\Users\\melho\\Downloads\\BASE FILL RATE 22.05.xlsx", sheet_name= "IND. PROG.")

OUTPUT_PATH = r"C:\Users\melho\Erro_de_onda.xlsx"

map_classe_setor = {
    "CLSETK":  "K",  "CLSETEV": "EV", "DEFAULT":  "INCORRETO",
    "CLSETA":  "A",  "CLSETG":  "G",  "CLSETJ":   "J",
    "CLSETAJ": "AJ", "CLSETC":  "C",  "CLSETY":   "Y",
    "CLSETEL": "EL", "CLSETEE": "EE", "CLSETS":   "S",
    "CLSETTB": "AJ", "CLSETUN": "XX", "CLSETE":   "E",
    "CLSETU":  "U",  "CLSETSS": "SS", "CLSETAG":  "AJ",
    "CLSETV":  "V",  "CLSETL":  "L",  "CLSETEF":  "EF",
    "CLSETEB": "G",  "CLSETYY": "YY",
}

map_erro_responsavel = {
    "ITEM SEM ESTOQUE DISPONIVEL":    "C.E",
    "MARCADO MANUALMENTE":            "C.E",
    "ENDERECO BLOQUEADO":             "C.E",
    "CASEPACK":                       "C.E",
    "ITEM SEM ESTOQUE PARCIAL":       "C.E",
    "ENDERECO FORA DE SERVICO":       "C.E",
    "IN-TRANSIT - CE":                "C.E",
    "CLASSE INCORRETA":               "C.E",
    "FLOW THROUNG":                   "FLOW THROUNG",
    "RTV":                            "O.P",
    "IN-TRANSIT":                     "O.P",
    "UNLOCATEDLOC":                   "O.P",
    "ARMAZENAGEM":                    "O.P",
    "ARMAZENAR":                      "O.P",
    "DEVOL-ESTQ":                     "O.P",
    "ESTOQUE INELEGIVEL":             "O.P",
    "ENDEREÇAMENTO INCORRETO":        "PCP",
    "ITEM SEM LOCAL DE SEPARACAO":    "PCP",
    "REARMAZENAR - CLASSE INCORRETA": "O.P",
    "RECEBIMENTO":                    "REC",
    "ERRO SISTEMICO":                 "T.I",
    "RESOLVIDO":                      "OK",
    "ENDEREÇADO":                     "OK",
    "Ok":                             "OK",
}

filtro_endereco   = ~estoqueEndereco['TIP_END'].isin(['RDR','SDR','MISC','PT BINS SEP1','PT BINS 1 ARM','STAGE'])
filtro_intransit  = (estoqueEndereco['TIP_END'] == 'INTRANSIT') & (estoqueEndereco['DATA_RECEB'].notna())
soma_artigos = estoqueEndereco[filtro_endereco | filtro_intransit].groupby('ITEM_ID')['VOLUMES'].sum()

indicador_operacao['Possui_Estoque'] = (
        indicador_operacao['ITEM_ID'].map(soma_artigos).fillna(0)
        .apply(lambda x: 'Sim' if x > 0 else 'Sem Estoque')
)

indicador_operacao['Chave'] = (indicador_operacao['DISTRO_NBR'].astype(str)
                                + "|" + indicador_operacao['LOJA'].astype(str)
                                  + "|" + indicador_operacao['ITEM_ID'].astype(str))

pendencia_embarque['Chave'] = (pendencia_embarque['DISTRO'].astype(str)
                                + "|" + pendencia_embarque['LOJA'].astype(str)
                                  + "|" + pendencia_embarque['ITEM'].astype(str))

base_erro_onda['Chave'] = (base_erro_onda['DISTRO_NBR'].fillna(0).astype(int).astype(str)
                                + "|" + base_erro_onda['LOJA'].fillna(0).astype(int).astype(str)
                                  + "|" + base_erro_onda['ITEM_ID'].fillna(0).astype(int).astype(str))

indicador_operacao['Emitida'] = (indicador_operacao['Chave'].isin(pendencia_embarque['Chave'])
                                 .map({True: 'Sim', False: 'Nao'}))
indicador_operacao['Erro de onda'] = (indicador_operacao['Chave'].map(base_erro_onda.set_index('Chave')['ERRO DE ONDA']).fillna('Nao'))

tem_endereco = indicador_operacao['ITEM_ID'].isin(editorLocal['ITEM'])
indicador_operacao['Enderecar ?'] = (
    (~indicador_operacao['SETOR'].eq('K')) &  # não é K
    (~tem_endereco)                           # não tem endereço
).map({True: 'Sim', False: 'Nao'})

indicador_operacao['Endereco Marcado'] = (
    indicador_operacao['ITEM_ID']
    .isin(enderecoMarcado['ITEM_ID'])
    .map({True: 'Sim', False: 'Nao'})
)

rtv_unlocated = estoqueEndereco[
        estoqueEndereco['ENDERECO'].isin(['RTV','UNLOCATEDLOC'])
    ].groupby('ITEM_ID')['ENDERECO'].first()
indicador_operacao['Local_Estoque_Status'] = indicador_operacao['ITEM_ID'].map(rtv_unlocated).fillna('Nao')

devol_arm_sep = estoqueEndereco[
        estoqueEndereco['ENDERECO'].isin(['DEVOL-ARM','DEVOL-SEP','DEVOL-ARMK','DEVOLUCAOK'])
    ].groupby('ITEM_ID')['ENDERECO'].first()
indicador_operacao['Devolucao_Status'] = indicador_operacao['ITEM_ID'].map(devol_arm_sep).fillna('Nao')

estoque_filtrado = estoqueEndereco[filtro_endereco]

estoque_real = (
    estoque_filtrado
    .groupby('ITEM_ID')['VOLUMES']
    .sum()
)

indicador_operacao['Estoque_Real'] = (
    indicador_operacao['ITEM_ID']
    .map(estoque_real)
    .fillna(0)
    .astype(int)
)

indicador_operacao['Armazenar_Item'] = np.where(
        indicador_operacao['ITEM_ID'].isin(
            estoqueEndereco.loc[estoqueEndereco['TIP_END'] == 'RDR', 'ITEM_ID']
        ),
        'Sim', 'Nao'
)

filtro_armazenar = estoqueEndereco[
        estoqueEndereco['TIP_END'].isin(['PT PALETE 3 ARM','PT PALETE 1 ARM','PT PALETE 2 ARM','PT PALETE 4 ARM'])
    ].copy()
filtro_armazenar['Setor_Correto'] = filtro_armazenar['CLASSE'].map(map_classe_setor).fillna('INCORRETO')
filtro_armazenar['Rearmazenar']   = filtro_armazenar['SETOR_PLAN_ESTOQUE'] != filtro_armazenar['Setor_Correto']
rearmazenar_item = filtro_armazenar.groupby('ITEM_ID')['Rearmazenar'].any()
indicador_operacao['Rearmazenar'] = indicador_operacao['ITEM_ID'].map(rearmazenar_item).fillna(False).map({True: "Sim", False: "Nao"})

status_lookup = editorLocal.drop_duplicates('ITEM').set_index('ITEM')['STATUS_LOCAL']

indicador_operacao['Status_Local'] = (
    indicador_operacao['ITEM_ID']
    .map(status_lookup)
    .fillna('Nao')
)

indicador_operacao['ERRO_REAL'] = np.select(
    [
        indicador_operacao['Possui_Estoque'] == 'Sem Estoque',
        indicador_operacao['Erro de onda'] != 'Nao',
        indicador_operacao['Status_Local'] == 'OUT-SERVICE',
        indicador_operacao['Endereco Marcado'] == 'Sim',
        indicador_operacao['Local_Estoque_Status'] == 'RTV',
        indicador_operacao['Local_Estoque_Status'] == 'UNLOCATEDLOC',
        indicador_operacao['Devolucao_Status'] != 'Nao',
        indicador_operacao['Enderecar ?'] == 'Sim',  # corrigido
        indicador_operacao['Rearmazenar'] == "Sim",
        indicador_operacao['Armazenar_Item'] == 'Sim',
        indicador_operacao['Emitida'] == 'Sim',
    ],
    [
        'ITEM SEM ESTOQUE DISPONIVEL',
        'ERRO DE ONDA',
        'ENDERECO FORA DE SERVICO',
        'MARCADO MANUALMENTE',
        'RTV',
        'UNLOCATEDLOC',
        'DEVOL-ESTQ',
        'ITEM SEM LOCAL DE SEPARACAO',
        'REARMAZENAR - CLASSE INCORRETA',
        'ARMAZENAR',
        'EMITIDA',
    ],
    default='Ok'
)

indicador_operacao['AREA_RESPONSAVEL'] = (
        indicador_operacao['ERRO_REAL'].map(map_erro_responsavel).fillna('NAO CLASSIFICADO')
    )

indicador_operacao['Erro solucionado'] = np.select(
    [

        ((indicador_operacao['Erro de onda'] == 'SEM ESTOQUE')) &
        (indicador_operacao['Possui_Estoque'] == 'Sim'),

        ((indicador_operacao['Erro de onda'] == 'SEM ESTOQUE')) &
        (indicador_operacao['Possui_Estoque'] == 'Sem Estoque'),

        ((indicador_operacao['Erro de onda'] == 'FORA DE SERVIÇO')) &
        (indicador_operacao['Status_Local'] != 'OUT-SERVICE'),

        ((indicador_operacao['Erro de onda'] == 'FORA DE SERVIÇO')) &
        (indicador_operacao['Status_Local'] == 'OUT-SERVICE'),

        ((indicador_operacao['Erro de onda'] == 'MARCADO MANUALMENTE')) &
        (indicador_operacao['Endereco Marcado'] == 'Nao'),

        ((indicador_operacao['Erro de onda'] == 'MARCADO MANUALMENTE')) &
        (indicador_operacao['Endereco Marcado'] == 'Sim'),

        ((indicador_operacao['Erro de onda'] == 'RTV')) &
        (indicador_operacao['Local_Estoque_Status'] != 'RTV'),

        ((indicador_operacao['Erro de onda'] == 'RTV')) &
        (indicador_operacao['Local_Estoque_Status'] == 'RTV'),

        ((indicador_operacao['Erro de onda'] == 'UNLOCATEDLOC')) &
        (indicador_operacao['Local_Estoque_Status'] != 'UNLOCATEDLOC'),

        ((indicador_operacao['Erro de onda'] == 'UNLOCATEDLOC')) &
        (indicador_operacao['Local_Estoque_Status'] == 'UNLOCATEDLOC'),

        ((indicador_operacao['Erro de onda'] == 'DEVOL-ARM')) &
        (indicador_operacao['Devolucao_Status'] == 'Nao'),

        ((indicador_operacao['Erro de onda'] == 'DEVOL-ARM')) &
        (indicador_operacao['Devolucao_Status'] == 'Sim'),

        ((indicador_operacao['Erro de onda'] == 'ARMAZENAR')) &
        (indicador_operacao['Armazenar_Item'] == 'Nao'),

        ((indicador_operacao['Erro de onda'] == 'ARMAZENAR')) &
        (indicador_operacao['Armazenar_Item'] == 'Sim'),

        ((indicador_operacao['Erro de onda'] == 'REARMAZENAR - CLASSE INCORRETA')) &
        (indicador_operacao['Rearmazenar'] == "Sim"),

        ((indicador_operacao['Erro de onda'] == 'REARMAZENAR - CLASSE INCORRETA')) &
        (indicador_operacao['Rearmazenar'] == "Nao"),

        ((indicador_operacao['Erro de onda'] == 'Nao')) &
        (indicador_operacao['Emitida'] == 'Sim'),

        ((indicador_operacao['Erro de onda'] == 'Nao')) &
        (indicador_operacao['Emitida'] == 'Nao'),
    ],

['Resolvido', 'Não resolvido', 'Resolvido', 'Não resolvido', 'Resolvido', 'Não resolvido',
 'Resolvido', 'Não resolvido', 'Resolvido', 'Não resolvido', 'Resolvido', 'Não resolvido',
 'Resolvido', 'Não resolvido','Resolvido', 'Não resolvido', 'Resolvido', 'Não resolvido', ],
default='Nao'
)  

indicador_operacao.to_excel(OUTPUT_PATH, index=False)

print(f"✅ Excel exportado para: {OUTPUT_PATH}")
print(f"   Total linhas  : {len(indicador_operacao):,}")
total = len(indicador_operacao)
erros = (indicador_operacao['ERRO_REAL'] != 'Ok').sum()

print(indicador_operacao[['Chave', 'Emitida', 'Erro de onda', 'Enderecar ?', 'Endereco Marcado', 'Local_Estoque_Status',
                           'Devolucao_Status', 'Estoque_Real', 'Armazenar_Item', 'Rearmazenar']])
