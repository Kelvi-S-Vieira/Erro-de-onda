"""
Análise de Erro de Onda
Execução diária — lê os relatórios operacionais e gera:
  • Excel : Aba ANALISE (dados completos) + Aba RESUMO (consolidado)
  • HTML  : Dashboard separado — hospedar no GitHub Pages e importar o Excel
"""

import sys
import json
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# CONFIGURAÇÃO DE PASTA — automática
# ─────────────────────────────────────────────

PADROES_ARQUIVOS = {
    "editor_local":       "Editor_Local_Separacao*.csv",
    "endereco_marcado":   "Endereco_Marcado_Contagem*.csv",
    "estoque_endereco":   "Estoque_por_Endereco*.csv",
    "indicador_operacao": "Indicador de Operação*.xlsb",
    "pendencia_embarque": "Indicador - Pendencia de Embarque*.csv",
    "base_erro_onda":     "BASE FILL RATE*.xlsx",
}

# ─────────────────────────────────────────────
# SELEÇÃO DE PASTA — salva config.json ao lado do .exe
# ─────────────────────────────────────────────
def _exe_dir() -> Path:
    """Retorna a pasta onde o executável (ou script) está."""
    if getattr(sys, 'frozen', False):          # rodando como .exe (PyInstaller)
        return Path(sys.executable).parent
    return Path(__file__).parent               # rodando como .py


def _config_path() -> Path:
    return _exe_dir() / "config.json"


def _carregar_pasta_config() -> Path | None:
    cfg = _config_path()
    if cfg.exists():
        try:
            data = json.loads(cfg.read_text(encoding='utf-8'))
            p = Path(data.get('pasta', ''))
            if p.exists():
                return p
        except Exception:
            pass
    return None


def _salvar_pasta_config(pasta: Path):
    _config_path().write_text(
        json.dumps({'pasta': str(pasta)}, ensure_ascii=False, indent=2),
        encoding='utf-8'
    )


def _selecionar_pasta(pasta_atual: Path | None) -> Path:
    """Abre janela Tkinter para o usuário escolher a pasta."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    inicial = str(pasta_atual) if pasta_atual else str(Path.home())

    pasta = filedialog.askdirectory(
        title="Selecione a pasta com os relatórios",
        initialdir=inicial,
    )
    root.destroy()

    if not pasta:
        messagebox.showerror("Cancelado", "Nenhuma pasta selecionada. O programa será encerrado.")
        sys.exit(0)

    return Path(pasta)


def resolver_pasta(forcar_novo: bool = False) -> Path:
    """
    Retorna a pasta de trabalho:
    - Se config.json existe e pasta é válida → usa direto (sem janela)
    - Caso contrário → abre seletor, salva e usa
    - forcar_novo=True → sempre abre o seletor (útil para trocar pasta)
    """
    pasta_salva = _carregar_pasta_config() if not forcar_novo else None

    if pasta_salva:
        print(f"📁 Pasta configurada: {pasta_salva}")
        return pasta_salva

    print("📂 Nenhuma pasta configurada. Abrindo seletor...")
    pasta = _selecionar_pasta(pasta_salva)
    _salvar_pasta_config(pasta)
    print(f"✅ Pasta salva: {pasta}")
    return pasta

# ─────────────────────────────────────────────
# MAPEAMENTOS
# ─────────────────────────────────────────────
MAP_CLASSE_SETOR = {
    "CLSETK":  "K",  "CLSETEV": "EV", "DEFAULT":  "INCORRETO",
    "CLSETA":  "A",  "CLSETG":  "G",  "CLSETJ":   "J",
    "CLSETAJ": "AJ", "CLSETC":  "C",  "CLSETY":   "Y",
    "CLSETEL": "EL", "CLSETEE": "EE", "CLSETS":   "S",
    "CLSETTB": "AJ", "CLSETUN": "XX", "CLSETE":   "E",
    "CLSETU":  "U",  "CLSETSS": "SS", "CLSETAG":  "AJ",
    "CLSETV":  "V",  "CLSETL":  "L",  "CLSETEF":  "EF",
    "CLSETEB": "G",  "CLSETYY": "YY",
}

MAP_ERRO_RESPONSAVEL = {
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

REGRAS_RESOLUCAO = {
    "SEM ESTOQUE":                    ("Possui_Estoque",        "Sim"),
    "FORA DE SERVIÇO":                ("Status_Local",          "OUT-SERVICE"),  # invertido
    "MARCADO MANUALMENTE":            ("Endereco Marcado",      "Nao"),
    "RTV":                            ("Local_Estoque_Status",  "Ok"),
    "UNLOCATEDLOC":                   ("Local_Estoque_Status",  "Ok"),
    "DEVOL-ARM":                      ("Devolucao_Status",      "Nao"),
    "ARMAZENAR":                      ("Armazenar_Item",        "Nao"),
    "REARMAZENAR - CLASSE INCORRETA": ("Rearmazenar",           "Nao"),
}

ENDERECO_BLOQUEADOS = ['RDR', 'SDR', 'MISC', 'PT BINS SEP1', 'PT BINS 1 ARM', 'STAGE']


# ─────────────────────────────────────────────
# 1. LEITURA DE ARQUIVOS
# ─────────────────────────────────────────────
def encontrar_arquivo(pasta: Path, padrao: str) -> Path:
    arquivos = sorted(pasta.glob(padrao), key=lambda p: p.stat().st_mtime, reverse=True)
    if not arquivos:
        raise FileNotFoundError(
            f"Nenhum arquivo encontrado com padrão '{padrao}' em '{pasta}'.\n"
            "Verifique se o arquivo foi baixado e a pasta está correta."
        )
    if len(arquivos) > 1:
        print(f"  ⚠  Múltiplos arquivos para '{padrao}'. Usando o mais recente: {arquivos[0].name}")
    return arquivos[0]


def carregar_dados(pasta: Path, padroes: dict) -> dict:
    print("📂 Carregando arquivos...")
    dfs = {}
    for chave, padrao in padroes.items():
        caminho = encontrar_arquivo(pasta, padrao)
        print(f"  ✓  {chave}: {caminho.name}")
        ext = caminho.suffix.lower()
        if ext == ".csv":
            dfs[chave] = pd.read_csv(caminho, low_memory=False)
        elif ext in (".xlsx", ".xlsb"):
            sheet = "1º - Ind. Operação" if chave == "indicador_operacao" else "IND. PROG."
            dfs[chave] = pd.read_excel(caminho, sheet_name=sheet)
    return dfs


# ─────────────────────────────────────────────
# 2. CÁLCULOS
# ─────────────────────────────────────────────
def calcular_estoque(estoque_df: pd.DataFrame) -> tuple:
    filtro_normal    = ~estoque_df['TIP_END'].isin(ENDERECO_BLOQUEADOS)
    filtro_intransit = (estoque_df['TIP_END'] == 'INTRANSIT') & estoque_df['DATA_RECEB'].notna()

    soma_artigos = (
        estoque_df[filtro_normal | filtro_intransit]
        .groupby('ITEM_ID')['VOLUMES'].sum()
    )
    rtv_unlocated = (
        estoque_df[estoque_df['ENDERECO'].isin(['RTV', 'UNLOCATEDLOC'])]
        .groupby('ITEM_ID')['ENDERECO'].first()
    )
    devol_status = (
        estoque_df[estoque_df['ENDERECO'].isin(['DEVOL-ARM', 'DEVOL-SEP', 'DEVOL-ARMK', 'DEVOLUCAOK'])]
        .groupby('ITEM_ID')['ENDERECO'].first()
    )
    estoque_real = estoque_df[filtro_normal].groupby('ITEM_ID')['VOLUMES'].sum()
    armazenar_item = estoque_df.loc[estoque_df['TIP_END'] == 'RDR', 'ITEM_ID']

    filtro_arm = estoque_df[
        estoque_df['TIP_END'].isin(['PT PALETE 3 ARM', 'PT PALETE 1 ARM', 'PT PALETE 2 ARM', 'PT PALETE 4 ARM'])
    ].copy()
    filtro_arm['Setor_Correto'] = filtro_arm['CLASSE'].map(MAP_CLASSE_SETOR).fillna('INCORRETO')
    filtro_arm['Rearmazenar']   = filtro_arm['SETOR_PLAN_ESTOQUE'] != filtro_arm['Setor_Correto']
    rearmazenar_item = filtro_arm.groupby('ITEM_ID')['Rearmazenar'].any()

    return soma_artigos, rtv_unlocated, devol_status, estoque_real, armazenar_item, rearmazenar_item


def montar_chaves(dfs: dict) -> dict:
    dfs['indicador_operacao']['Chave'] = (
        dfs['indicador_operacao']['DISTRO_NBR'].astype(str) + "|" +
        dfs['indicador_operacao']['LOJA'].astype(str)       + "|" +
        dfs['indicador_operacao']['ITEM_ID'].astype(str)
    )
    dfs['pendencia_embarque']['Chave'] = (
        dfs['pendencia_embarque']['DISTRO'].astype(str) + "|" +
        dfs['pendencia_embarque']['LOJA'].astype(str)   + "|" +
        dfs['pendencia_embarque']['ITEM'].astype(str)
    )
    dfs['base_erro_onda']['Chave'] = (
        dfs['base_erro_onda']['DISTRO_NBR'].fillna(0).astype(int).astype(str) + "|" +
        dfs['base_erro_onda']['LOJA'].fillna(0).astype(int).astype(str)       + "|" +
        dfs['base_erro_onda']['ITEM_ID'].fillna(0).astype(int).astype(str)
    )
    return dfs


def enriquecer_indicador(df, dfs, soma_artigos, rtv_unlocated, devol_status,
                          estoque_real, armazenar_item, rearmazenar_item):
    chaves_pendencia  = set(dfs['pendencia_embarque']['Chave'])
    lookup_erro_onda  = dfs['base_erro_onda'].set_index('Chave')['ERRO DE ONDA']
    status_local_look = dfs['editor_local'].drop_duplicates('ITEM').set_index('ITEM')['STATUS_LOCAL']
    itens_editor      = set(dfs['editor_local']['ITEM'])
    itens_marc        = set(dfs['endereco_marcado']['ITEM_ID'])
    itens_armazenar   = set(armazenar_item)

    df['Possui_Estoque']       = df['ITEM_ID'].map(soma_artigos).fillna(0).gt(0).map({True: 'Sim', False: 'Sem Estoque'})
    df['Emitida']              = df['Chave'].isin(chaves_pendencia).map({True: 'Sim', False: 'Nao'})
    df['Erro de onda']         = df['Chave'].map(lookup_erro_onda).fillna('Sem erro')
    df['Enderecar ?']          = (~df['SETOR'].eq('K') & ~df['ITEM_ID'].isin(itens_editor)).map({True: 'Realizar Endereçamento', False: 'Nao'})
    df['Endereco Marcado']     = df['ITEM_ID'].isin(itens_marc).map({True: 'Sim', False: 'Nao'})
    df['Local_Estoque_Status'] = df['ITEM_ID'].map(rtv_unlocated).fillna('Ok')
    df['Devolucao_Status']     = df['ITEM_ID'].map(devol_status).fillna('Nao')
    df['Estoque_Real']         = df['ITEM_ID'].map(estoque_real).fillna(0).astype(int)
    df['Armazenar_Item']       = df['ITEM_ID'].isin(itens_armazenar).map({True: 'Sim', False: 'Nao'})
    df['Rearmazenar']          = df['ITEM_ID'].map(rearmazenar_item).fillna(False).map({True: 'Sim', False: 'Nao'})
    df['Status_Local']         = df['ITEM_ID'].map(status_local_look).fillna('Sem Endereço')
    return df


def classificar_erro_real(df):
    condicoes = [
        df['Possui_Estoque']       == 'Sem Estoque',
        df['Erro de onda']         != 'Sem erro',
        df['Status_Local']         == 'OUT-SERVICE',
        df['Endereco Marcado']     == 'Sim',
        df['Local_Estoque_Status'] == 'RTV',
        df['Local_Estoque_Status'] == 'UNLOCATEDLOC',
        df['Devolucao_Status']     != 'Nao',
        df['Enderecar ?']          == 'Realizar Endereçamento',
        df['Rearmazenar']          == 'Sim',
        df['Armazenar_Item']       == 'Sim',
        df['Emitida']              == 'Sim',
    ]
    valores = [
        'ITEM SEM ESTOQUE DISPONIVEL', 'ERRO DE ONDA', 'ENDERECO FORA DE SERVICO',
        'MARCADO MANUALMENTE', 'RTV', 'UNLOCATEDLOC', 'DEVOL-ESTQ',
        'ITEM SEM LOCAL DE SEPARACAO', 'REARMAZENAR - CLASSE INCORRETA', 'ARMAZENAR', 'EMITIDA',
    ]
    df['ERRO_REAL']        = np.select(condicoes, valores, default='Ok')
    df['AREA_RESPONSAVEL'] = df['ERRO_REAL'].map(MAP_ERRO_RESPONSAVEL).fillna('NAO CLASSIFICADO')
    return df


def calcular_resolucao(df):
    tem_erro = df['Erro de onda'] != 'Sem erro'

    # Padrão: sem erro = sem pendência | com erro = não resolvido (até prova em contrário)
    resultado = pd.Series('Sem pendência', index=df.index)
    resultado[tem_erro] = 'Não resolvido'

    for erro, (coluna, valor_resolvido) in REGRAS_RESOLUCAO.items():
        mascara = tem_erro & (df['Erro de onda'] == erro)
        if erro == "FORA DE SERVIÇO":
            resultado[mascara & (df[coluna] != 'OUT-SERVICE')] = 'Resolvido'
            resultado[mascara & (df[coluna] == 'OUT-SERVICE')] = 'Não resolvido'
        else:
            resultado[mascara & (df[coluna] == valor_resolvido)] = 'Resolvido'
            resultado[mascara & (df[coluna] != valor_resolvido)] = 'Não resolvido'

    # Sem erro na base: verifica emissão
    sem_erro = ~tem_erro
    resultado[sem_erro & (df['Emitida'] == 'Sim')] = 'Resolvido'

    df['Erro solucionado'] = resultado
    return df


# ─────────────────────────────────────────────
# 3. EXCEL — ANALISE + RESUMO
# ─────────────────────────────────────────────
COR_HEADER      = "1F3864"
COR_HEADER_FONT = "FFFFFF"
COR_VERDE       = "C6EFCE"
COR_VERDE_FONT  = "276221"
COR_VERMELHO    = "FFC7CE"
COR_VERM_FONT   = "9C0006"
COR_AMARELO     = "FFEB9C"
COR_AMAR_FONT   = "9C6500"
COR_CINZA       = "F2F2F2"
COR_ACCENT      = "2E75B6"


def estilo_header(ws, linha, n_colunas):
    fill = PatternFill("solid", start_color=COR_HEADER, end_color=COR_HEADER)
    font = Font(bold=True, color=COR_HEADER_FONT, name="Arial", size=10)
    alin = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col in range(1, n_colunas + 1):
        c = ws.cell(row=linha, column=col)
        c.fill, c.font, c.alignment = fill, font, alin


def aplicar_cor_celula(ws, col_letra, linha_ini, linha_fim, mapa):
    for row in range(linha_ini, linha_fim + 1):
        cell = ws[f"{col_letra}{row}"]
        val  = str(cell.value or "")
        for chave, (bg, fg) in mapa.items():
            if chave.lower() in val.lower():
                cell.fill = PatternFill("solid", start_color=bg, end_color=bg)
                cell.font = Font(color=fg, name="Arial", size=9, bold=True)
                break


def escrever_aba_analise(wb, df):
    ws = wb.create_sheet("ANALISE")
    ws.freeze_panes = "A2"
    headers = list(df.columns)
    ws.append(headers)
    for row in df.itertuples(index=False):
        ws.append(list(row))
    n_linhas = len(df)
    estilo_header(ws, 1, len(headers))
    for i, col in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(str(col)) + 2)
    col_idx = {h: i for i, h in enumerate(headers, 1)}
    mapa_res = {
        "Resolvido":     (COR_VERDE,    COR_VERDE_FONT),
        "Não resolvido": (COR_VERMELHO, COR_VERM_FONT),
    }
    mapa_err = {"Ok": (COR_VERDE, COR_VERDE_FONT)}
    if "Erro solucionado" in col_idx:
        aplicar_cor_celula(ws, get_column_letter(col_idx["Erro solucionado"]), 2, n_linhas + 1, mapa_res)
    if "ERRO_REAL" in col_idx:
        aplicar_cor_celula(ws, get_column_letter(col_idx["ERRO_REAL"]), 2, n_linhas + 1, mapa_err)
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"


def escrever_aba_resumo(wb, df, data_exec):
    ws = wb.create_sheet("RESUMO")
    ws.column_dimensions['A'].width = 34
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 18

    def titulo(texto, row, span=5):
        c = ws.cell(row=row, column=1, value=texto)
        c.font      = Font(bold=True, color="FFFFFF", name="Arial", size=11)
        c.fill      = PatternFill("solid", start_color=COR_ACCENT, end_color=COR_ACCENT)
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[row].height = 22
        for col in range(2, span + 1):
            ws.cell(row=row, column=col).fill = PatternFill("solid", start_color=COR_ACCENT, end_color=COR_ACCENT)

    def sub_header(row, cols):
        fill = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
        font = Font(bold=True, name="Arial", size=9, color="1F3864")
        for i, txt in enumerate(cols, 1):
            c = ws.cell(row=row, column=i, value=txt)
            c.fill, c.font = fill, font
            c.alignment = Alignment(horizontal="center")

    def linha(row, valores, dest=None):
        for i, v in enumerate(valores, 1):
            c = ws.cell(row=row, column=i, value=v)
            c.font      = Font(name="Arial", size=9)
            c.alignment = Alignment(horizontal="center" if i > 1 else "left")
            if i % 2 == 0:
                c.fill = PatternFill("solid", start_color=COR_CINZA, end_color=COR_CINZA)
        if dest:
            idx, bg, fg = dest
            c = ws.cell(row=row, column=idx)
            c.fill = PatternFill("solid", start_color=bg, end_color=bg)
            c.font = Font(bold=True, color=fg, name="Arial", size=9)

    row = 1
    ws.cell(row=row, column=1, value=f"Relatório de Erro de Onda — {data_exec}").font = Font(
        bold=True, size=14, name="Arial", color=COR_ACCENT)
    ws.row_dimensions[row].height = 28
    row += 2

    total    = len(df)
    com_erro = (df['Erro de onda'] != 'Sem erro').sum()
    sem_erro = total - com_erro
    emitidas = (df['Emitida'] == 'Sim').sum()

    titulo("📊 VISÃO GERAL", row); row += 1
    sub_header(row, ["Métrica", "Qtd", "% do Total"]); row += 1
    for label, qtd in [("Total analisado", total), ("Com erro", com_erro),
                        ("Sem erro (Ok)", sem_erro), ("Emitidos", emitidas)]:
        linha(row, [label, qtd, f"{qtd/total:.1%}" if total else "—"]); row += 1
    row += 1

    titulo("🔴 ERROS POR TIPO", row); row += 1
    sub_header(row, ["Tipo de Erro", "Qtd", "% s/ Erros", "Responsável"]); row += 1
    dist_erro = (df[df['Erro de onda'] != 'Sem erro']
                 .groupby(['ERRO_REAL', 'AREA_RESPONSAVEL']).size()
                 .reset_index(name='Qtd').sort_values('Qtd', ascending=False))
    for _, r in dist_erro.iterrows():
        pct = f"{r['Qtd']/com_erro:.1%}" if com_erro else "—"
        linha(row, [r['ERRO_REAL'], r['Qtd'], pct, r['AREA_RESPONSAVEL']]); row += 1
    row += 1

    titulo("✅ STATUS DE RESOLUÇÃO", row); row += 1
    sub_header(row, ["Status", "Qtd", "% c/ Erro Base"]); row += 1
    com_base   = df[df['Erro de onda'] != 'Sem erro']
    total_base = len(com_base)
    for status, qtd in com_base['Erro solucionado'].value_counts().items():
        pct = f"{qtd/total_base:.1%}" if total_base else "—"
        bg  = COR_VERDE   if status == "Resolvido"     else COR_VERMELHO if status == "Não resolvido" else COR_AMARELO
        fg  = COR_VERDE_FONT if status == "Resolvido"  else COR_VERM_FONT  if status == "Não resolvido" else COR_AMAR_FONT
        linha(row, [status, qtd, pct], dest=(1, bg, fg)); row += 1
    row += 1

    titulo("📋 RESOLUÇÃO POR TIPO DE ERRO", row); row += 1
    sub_header(row, ["Erro de Onda", "Total", "Resolvidos", "Não resolvidos", "% Resolução"]); row += 1
    pivot = (com_base.groupby(['Erro de onda', 'Erro solucionado'])
             .size().unstack(fill_value=0).reset_index())
    for _, r in pivot.iterrows():
        res  = r.get('Resolvido', 0)
        nres = r.get('Não resolvido', 0)
        tot  = res + nres
        pct  = f"{res/tot:.1%}" if tot else "—"
        bg = COR_VERDE if res >= nres else COR_VERMELHO
        fg = COR_VERDE_FONT if res >= nres else COR_VERM_FONT
        linha(row, [r['Erro de onda'], tot, res, nres, pct], dest=(5, bg, fg)); row += 1
    row += 1

    titulo("🏢 ERROS POR ÁREA RESPONSÁVEL", row); row += 1
    sub_header(row, ["Área", "Qtd", "% do Total"]); row += 1
    for area, qtd in df[df['Erro de onda'] != 'Sem erro']['AREA_RESPONSAVEL'].value_counts().items():
        pct = f"{qtd/com_erro:.1%}" if com_erro else "—"
        linha(row, [area, qtd, pct]); row += 1


def exportar_excel(df, caminho):
    data_exec = datetime.now().strftime("%d/%m/%Y %H:%M")
    print("\n📊 Gerando Excel...")
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="ANALISE", index=False)
    wb = load_workbook(caminho)
    escrever_aba_resumo(wb, df, data_exec)
    # Reordena: RESUMO primeiro
    wb.move_sheet("RESUMO", offset=-wb.sheetnames.index("RESUMO"))
    wb.save(caminho)
    print(f"  ✓  Excel salvo em: {caminho}")


# ─────────────────────────────────────────────
# 4. SUMÁRIO NO CONSOLE
# ─────────────────────────────────────────────
def imprimir_sumario(df):
    total   = len(df)
    erros   = (df['Erro de onda'] != 'Sem erro').sum()
    resolv  = (df['Erro solucionado'] == 'Resolvido').sum()
    nresolv = (df['Erro solucionado'] == 'Não resolvido').sum()

    print("\n" + "═" * 55)
    print(f"  SUMÁRIO — {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("═" * 55)
    print(f"  Total analisado   : {total:>6,}")
    print(f"  Com erro          : {erros:>6,}  ({erros/total:.1%})")
    print(f"  Resolvidos        : {resolv:>6,}")
    print(f"  Não resolvidos    : {nresolv:>6,}")
    print()
    print("  Erros por tipo:")
    for tipo, qtd in df[df['Erro de onda'] != 'Sem erro']['ERRO_REAL'].value_counts().items():
        print(f"    {tipo:<40} {qtd:>5,}")
    print()
    print("  Resolução por tipo de erro:")
    com_base = df[df['Erro de onda'] != 'Sem erro']
    for erro in com_base['Erro de onda'].unique():
        sub = com_base[com_base['Erro de onda'] == erro]
        res = (sub['Erro solucionado'] == 'Resolvido').sum()
        tot = len(sub)
        print(f"    {erro:<40} {res:>4}/{tot:<4}  ({res/tot:.0%})" if tot else f"    {erro:<40} —")
    print("═" * 55)


# ─────────────────────────────────────────────
# 5. MAIN
# ─────────────────────────────────────────────
def main():
    # Passa --nova-pasta como argumento para forçar novo seletor
    forcar = '--nova-pasta' in sys.argv
    PASTA = resolver_pasta(forcar_novo=forcar)
    OUTPUT_EXCEL = PASTA / "Erro_de_onda.xlsx"

    try:
        dfs = carregar_dados(PASTA, PADROES_ARQUIVOS)
    except FileNotFoundError as e:
        print(f"\n❌ ERRO AO CARREGAR ARQUIVO:\n{e}")
        sys.exit(1)

    print("\n⚙️  Processando dados...")
    dfs = montar_chaves(dfs)
    (soma_artigos, rtv_unlocated, devol_status,
     estoque_real, armazenar_item, rearmazenar_item) = calcular_estoque(dfs['estoque_endereco'])

    df = dfs['indicador_operacao'].copy()
    df = enriquecer_indicador(df, dfs, soma_artigos, rtv_unlocated, devol_status,
                               estoque_real, armazenar_item, rearmazenar_item)
    df = classificar_erro_real(df)
    df = calcular_resolucao(df)

    exportar_excel(df, OUTPUT_EXCEL)
    imprimir_sumario(df)
    print(f"\n✅ Concluído!")


if __name__ == "__main__":
    main()