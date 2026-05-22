import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# ── Caminhos de entrada — ajuste conforme necessário ──────────────
editorLocal           = pd.read_csv(r"C:\Users\melho\Desktop\Pasta relatorios\EditorLocal\Editor_Local_Separacao.csv", low_memory=False)
enderecoMarcado       = pd.read_csv(r"C:\Users\melho\Desktop\Pasta relatorios\EnderecoMarcado\Endereco_Marcado_Contagem.csv", low_memory=False)
estoqueEndereco       = pd.read_csv(r"C:\Users\melho\Desktop\Pasta relatorios\Estoque\Estoque_por_Endereco.csv", low_memory=False)
frequenciaAbastecimento = pd.read_excel(r"C:\Users\melho\Desktop\Pasta relatorios\Frequência\Simulador-de-Frequência-Abastecimento-Semana.xlsx", sheet_name="Tabela Datas (Query)")
mapaVolumes           = pd.read_csv(r"C:\Users\melho\Desktop\Pasta relatorios\Mapa\Mapa_de_Volumes.csv", low_memory=False)

# ── Caminho de saída — CSV consumido pelo dashboard HTML ──────────
OUTPUT_PATH = r"C:\Users\melho\fillrate_dashboard.csv"

# ─────────────────────────────────────────────────────────────────
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

# ─────────────────────────────────────────────────────────────────
def projecaoFillRate(editorLocal, enderecoMarcado, estoqueEndereco, frequenciaAbastecimento, mapaVolumes):

    print("Iniciando projeção de fill rate...")

    # ── Ajuste de tipos (CRÍTICO) ───────────────────────────────
    mapaVolumes['FILIAL'] = mapaVolumes['FILIAL'].astype(str)
    frequenciaAbastecimento['COD'] = frequenciaAbastecimento['COD'].astype(str)

    # ── Ajuste de datas ────────────────────────────────────────
    frequenciaAbastecimento['EXPEDIÇÃO'] = pd.to_datetime(
        frequenciaAbastecimento['EXPEDIÇÃO'], errors='coerce'
    ).dt.date

    data_inicio = datetime.now().date()
    data_fim    = data_inicio + timedelta(days=2)

    # ── Buscar lojas no período (mais robusto) ─────────────────
    lojas_na_data = frequenciaAbastecimento.loc[
        frequenciaAbastecimento['EXPEDIÇÃO'].between(data_inicio, data_fim),
        'COD'
    ].unique()

    # ── Fallback (EVITA BASE VAZIA) ────────────────────────────
    if len(lojas_na_data) == 0:
        print("⚠️ Nenhuma loja encontrada no período — usando todas")
        lojas_na_data = frequenciaAbastecimento['COD'].unique()

    # ── Criar flag (SEM FILTRAR BASE) ──────────────────────────
    mapaVolumes['EXPEDE_D2'] = mapaVolumes['FILIAL'].isin(lojas_na_data)

    # ── BASE COMPLETA (igual Excel) ────────────────────────────
    mapaVolumesFiltrado = mapaVolumes.copy()

    mapaVolumesFiltrado['CHAVE'] = (
        mapaVolumesFiltrado['STOCK_ORDER'].astype(str) + "_" +
        mapaVolumesFiltrado['FILIAL'].astype(str) + "_" +
        mapaVolumesFiltrado['ARTIGO'].astype(str)
    )

    # ── Estoque total por item ────────────────────────────────────
    filtro_endereco   = ~estoqueEndereco['TIP_END'].isin(['SDR','MISC','PT BINS SEP1','PT BINS 1 ARM','INTRANSIT'])
    filtro_intransit  = (estoqueEndereco['TIP_END'] == 'INTRANSIT') & (estoqueEndereco['DATA_RECEB'].notna())
    soma_artigos = estoqueEndereco[filtro_endereco | filtro_intransit].groupby('ITEM_ID')['VOLUMES'].sum()

    mapaVolumesFiltrado['POSSUI_ESTOQUE'] = (
        mapaVolumesFiltrado['ARTIGO'].map(soma_artigos).fillna(0)
        .apply(lambda x: 'Sim' if x > 0 else 'Sem Estoque')
    )

    rtv_unlocated = estoqueEndereco[
        estoqueEndereco['ENDERECO'].isin(['RTV','UNLOCATEDLOC'])
    ].groupby('ITEM_ID')['ENDERECO'].first()
    mapaVolumesFiltrado['LOCAL_ESTOQUE_STATUS'] = mapaVolumesFiltrado['ARTIGO'].map(rtv_unlocated).fillna('Nao')

    devol_arm_sep = estoqueEndereco[
        estoqueEndereco['ENDERECO'].isin(['DEVOL-ARM','DEVOL-SEP','DEVOL-ARMK','DEVOLUCAOK'])
    ].groupby('ITEM_ID')['ENDERECO'].first()
    mapaVolumesFiltrado['DEVOLUCAO_STATUS'] = mapaVolumesFiltrado['ARTIGO'].map(devol_arm_sep).fillna('Nao')

    estoque_real = estoqueEndereco[estoqueEndereco['TIP_END'] != 'RDR'].groupby('ITEM_ID')['VOLUMES'].sum()
    mapaVolumesFiltrado['ESTOQUE_REAL'] = mapaVolumesFiltrado['ARTIGO'].map(estoque_real).fillna(0)

    mapaVolumesFiltrado['ARMAZENAR_ITEM'] = np.where(
        mapaVolumesFiltrado['ARTIGO'].isin(
            estoqueEndereco.loc[estoqueEndereco['TIP_END'] == 'RDR', 'ITEM_ID']
        ) & (mapaVolumesFiltrado['VOLUME'] > mapaVolumesFiltrado['ESTOQUE_REAL']),
        'Sim', 'Nao'
    )

    filtro_armazenar = estoqueEndereco[
        estoqueEndereco['TIP_END'].isin(['PT PALETE 3 ARM','PT PALETE 1 ARM','PT PALETE 2 ARM'])
    ].copy()
    filtro_armazenar['SETOR_CORRETO'] = filtro_armazenar['CLASSE'].map(map_classe_setor).fillna('INCORRETO')
    filtro_armazenar['REARMAZENAR']   = filtro_armazenar['SETOR_PLAN_ESTOQUE'] != filtro_armazenar['SETOR_CORRETO']
    rearmazenar_item = filtro_armazenar.groupby('ITEM_ID')['REARMAZENAR'].any()
    mapaVolumesFiltrado['REARMAZENAR'] = mapaVolumesFiltrado['ARTIGO'].map(rearmazenar_item).fillna(False)

    # ── Lookup LOCAL ──────────────────────────────────────────────
    editorLocal_lookup = editorLocal[['ITEM','LOCAL']].drop_duplicates('ITEM')
    mapaVolumesFiltrado = mapaVolumesFiltrado.merge(
        editorLocal_lookup, how='left', left_on='ARTIGO', right_on='ITEM'
    ).drop(columns=['ITEM']).fillna({'LOCAL': 'Enderecar'})

    mapaVolumesFiltrado['LOCAL'] = np.where(
        mapaVolumesFiltrado['SETOR'] == 'K', 'Encabidados', mapaVolumesFiltrado['LOCAL']
    )

    # ── Classe do setor ───────────────────────────────────────────
    estoqueEndereco['CLASSE_SETOR'] = estoqueEndereco['CLASSE'].map(map_classe_setor).fillna('INCORRETO')
    classe_lookup = estoqueEndereco[['ITEM_ID','CLASSE_SETOR']].drop_duplicates('ITEM_ID')
    mapaVolumesFiltrado = mapaVolumesFiltrado.merge(
        classe_lookup, how='left', left_on='ARTIGO', right_on='ITEM_ID'
    ).drop(columns=['ITEM_ID']).fillna({'CLASSE_SETOR': 'INCORRETO'})

    mapaVolumesFiltrado['CLASSE_CORRETA'] = np.where(
        mapaVolumesFiltrado['CLASSE_SETOR'] == mapaVolumesFiltrado['SETOR'], 'Sim', 'Nao'
    )

    # ── Status local ──────────────────────────────────────────────
    status_lookup = editorLocal[['ITEM','STATUS_LOCAL']].drop_duplicates('ITEM')
    mapaVolumesFiltrado = mapaVolumesFiltrado.merge(
        status_lookup, how='left', left_on='ARTIGO', right_on='ITEM'
    ).drop(columns=['ITEM'])

    mapaVolumesFiltrado['ARTIGO']            = mapaVolumesFiltrado['ARTIGO'].astype(str)
    enderecoMarcado['ITEM_ID']               = enderecoMarcado['ITEM_ID'].astype(str)
    mapaVolumesFiltrado = mapaVolumesFiltrado.merge(
        enderecoMarcado[['ITEM_ID','P_TIPO']], how='left', left_on='ARTIGO', right_on='ITEM_ID'
    ).drop(columns=['ITEM_ID'])

    mapaVolumesFiltrado['MARCADO PARA CONTAGEM'] = np.where(
        mapaVolumesFiltrado['P_TIPO'] == 'MM', 'Sim', 'Nao'
    )

    # ── Definição do erro real ────────────────────────────────────
    mapaVolumesFiltrado['ERRO_REAL'] = np.select(
        [
            mapaVolumesFiltrado['POSSUI_ESTOQUE']          == 'Sem Estoque',
            mapaVolumesFiltrado['MARCADO PARA CONTAGEM']   == 'Sim',
            mapaVolumesFiltrado['STATUS_LOCAL']            == 'OUT-SERVICE',
            mapaVolumesFiltrado['REARMAZENAR']             == True,
            mapaVolumesFiltrado['ARMAZENAR_ITEM']          == 'Sim',
            mapaVolumesFiltrado['LOCAL_ESTOQUE_STATUS']    == 'RTV',
            mapaVolumesFiltrado['LOCAL_ESTOQUE_STATUS']    == 'UNLOCATEDLOC',
            mapaVolumesFiltrado['DEVOLUCAO_STATUS']        != 'Nao',
            mapaVolumesFiltrado['LOCAL']                   == 'Enderecar',
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
            'ITEM SEM LOCAL DE SEPARACAO',
        ],
        default='Ok'
    )

    mapaVolumesFiltrado['AREA_RESPONSAVEL'] = (
        mapaVolumesFiltrado['ERRO_REAL'].map(map_erro_responsavel).fillna('NAO CLASSIFICADO')
    )

    # ── Deduplicar ────────────────────────────────────────────────
    mapaVolumesFiltrado = mapaVolumesFiltrado.drop_duplicates('CHAVE')

    # ── Exportar CSV para o dashboard ────────────────────────────
    # Colunas essenciais para o front-end:
    #   ERRO_REAL, AREA_RESPONSAVEL, LOCAL
    # Exportamos o dataframe completo (o front ignora colunas extras)
    mapaVolumesFiltrado.to_csv(OUTPUT_PATH, index=False, sep=';', encoding='utf-8-sig')

    print(f"✅ CSV exportado para: {OUTPUT_PATH}")
    print(f"   Total linhas  : {len(mapaVolumesFiltrado):,}")
    total = len(mapaVolumesFiltrado)
    erros = (mapaVolumesFiltrado['ERRO_REAL'] != 'Ok').sum()

    fr = (1 - erros / total) * 100 if total > 0 else 0
    print(f"   Total erros   : {erros:,}")
    print(f"   Fill Rate     : {fr:.2f}%")

    return mapaVolumesFiltrado


if __name__ == "__main__":
    projecaoFillRate(
        editorLocal,
        enderecoMarcado,
        estoqueEndereco,
        frequenciaAbastecimento,
        mapaVolumes
    )
