import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

path_file = "crud.xlsx"
hoja = "Hoja1"
df = pd.read_excel(path_file, sheet_name=hoja)
pd.options.mode.chained_assignment = None  # default='warn'

# PARAMETROS GAVETAS
alto = 220
ancho = 400
largo = 600
peso = 30
a1 = largo * ancho * alto
a1_capacidad = 1
a2_ancho_gav = 390
a2_largo_gav = 293
a2_vol_gav = a2_ancho_gav * a2_largo_gav * alto
a2_peso_gav = peso / 2
a2_capacidad = 5
a3_ancho_gav = 390
a3_largo_gav = 145
a3_vol_gav = a3_ancho_gav * a3_largo_gav * alto
a3_peso_gav = peso / 3
a3_capacidad = 30

# TRATA COLUMNAS
df.columns = ["INFO", "COD", "EQUIV", "QTY", "TICK", "CLNT",
              "8020", "MES", "PESO", "ANCHO", "LARGO", "ALTO", "CD"]
df['VOL'] = df['ALTO'] * df['ANCHO'] * df['LARGO']
df['CANTxVOL'] = a1 / df['VOL']
df['CANTxPESO'] = peso / df['PESO']
df['LIMITANTE'] = np.where(df['CANTxPESO'] > df['CANTxVOL'], 'VOLUMEN', 'PESO')
df['CONTxMES'] = np.where(df['LIMITANTE'] == 'VOLUMEN', (df['QTY'] /
                                                         df['CANTxVOL']) / df['MES'], (df['QTY'] / df['CANTxPESO']) / df['MES'])
skus_in = len(df)


# FILTROS
meses_vendidos = 1  # MOD
df_desc_meses = df[df["MES"] <= meses_vendidos]
indexNames = df_desc_meses.index
df.drop(indexNames, inplace=True)

# Cantidad menor a 5
cantidad = 5
df_desc_QTY = df[df["QTY"] < cantidad]
indexNames = df_desc_QTY.index
df.drop(indexNames, inplace=True)

# Sin datos dimensionales
df_desc_PESO = df.loc[(df["PESO"] <= 0) | (pd.isnull(df["PESO"]))]
indexNames = df_desc_PESO.index
df.drop(indexNames, inplace=True)

df_desc_ANCHO = df.loc[(df["ANCHO"] <= 0) | (pd.isnull(df["ANCHO"]))]
indexNames = df_desc_ANCHO.index
df.drop(indexNames, inplace=True)

df_desc_LARGO = df.loc[(df["LARGO"] <= 0) | (pd.isnull(df["LARGO"]))]
indexNames = df_desc_LARGO.index
df.drop(indexNames, inplace=True)

df_desc_ALTO = df.loc[(df["ALTO"] <= 0) | (pd.isnull(df["ALTO"]))]
indexNames = df_desc_ALTO.index
df.drop(indexNames, inplace=True)

frames = [df_desc_PESO, df_desc_ANCHO, df_desc_LARGO, df_desc_ALTO]
df_desc_sindatos = pd.concat(frames)


# Cantidad por volumen <1
df_desc_CxV = df[df["CANTxVOL"] < 1]
indexNames = df_desc_CxV.index
df.drop(indexNames, inplace=True)
# Cantidad por peso <1
df_desc_CxP = df[df["CANTxPESO"] < 1]
indexNames = df_desc_CxP.index
df.drop(indexNames, inplace=True)
frames = [df_desc_CxV, df_desc_CxP]
df_desc_limitante = pd.concat(frames)

# CALCULO A3
df_a3 = df.loc[((df['ANCHO'] < a3_ancho_gav) & (df['LARGO'] < a3_largo_gav))]
indexNames = df_a3.index
df.drop(indexNames, inplace=True)
df_a3.drop(["CANTxVOL", "CANTxPESO", "LIMITANTE", "CONTxMES"], axis=1)
df_a3["CANTxVOL"] = a3_vol_gav / df_a3["VOL"]
df_a3["CANTxPESO"] = a3_peso_gav / df_a3["PESO"]
df_a3["LIMITANTE"] = np.where(
    df_a3["CANTxPESO"] > df_a3["CANTxVOL"], "VOLUMEN", "PESO")
df_a3["CONTxMES"] = np.where(df_a3["LIMITANTE"] == "VOLUMEN", (df_a3["QTY"] /
                                                               df_a3["CANTxVOL"]) / df_a3["MES"], (df_a3["QTY"] / df_a3["CANTxPESO"]) / df_a3["MES"])

# CALCULO A2
df_a2 = df.loc[((df['ANCHO'] < a2_ancho_gav) & (df['LARGO'] < a2_largo_gav))]
indexNames = df_a2.index
df.drop(indexNames, inplace=True)
df_a2.drop(["CANTxVOL", "CANTxPESO", "LIMITANTE", "CONTxMES"], axis=1)
df_a2["CANTxVOL"] = a2_vol_gav / df_a2["VOL"]
df_a2["CANTxPESO"] = a2_peso_gav / df_a2["PESO"]
df_a2["LIMITANTE"] = np.where(
    df_a2["CANTxPESO"] > df_a2["CANTxVOL"], "VOLUMEN", "PESO")
df_a2["CONTxMES"] = np.where(df_a2["LIMITANTE"] == "VOLUMEN", (df_a2["QTY"] /
                                                               df_a2["CANTxVOL"]) / df_a2["MES"], (df_a2["QTY"] / df_a2["CANTxPESO"]) / df_a2["MES"])

# CALCULO A1
df_a1 = df.loc[((df['ANCHO'] < ancho) & (df['LARGO'] < largo))]
indexNames = df_a1.index
df.drop(indexNames, inplace=True)
df_a1.drop(["CANTxVOL", "CANTxPESO", "LIMITANTE", "CONTxMES"], axis=1)
df_a1["CANTxVOL"] = a1 / df_a1["VOL"]
df_a1["CANTxPESO"] = peso / df_a1["PESO"]
df_a1["LIMITANTE"] = np.where(
    df_a1["CANTxPESO"] > df_a1["CANTxVOL"], "VOLUMEN", "PESO")
df_a1["CONTxMES"] = np.where(df_a1["LIMITANTE"] == "VOLUMEN", (df_a1["QTY"] /
                                                               df_a1["CANTxVOL"]) / df_a1["MES"], (df_a1["QTY"] / df_a1["CANTxPESO"]) / df_a1["MES"])




# VARIABLES SUMA UNIDADES Y SKUS
mes_SKU = len(df_desc_meses)
mes_QTY = df_desc_meses['QTY'].sum()  # 6532
qty_SKU = len(df_desc_QTY)
qty_QTY = df_desc_QTY['QTY'].sum()
sd_SKU = len(df_desc_sindatos)
sd_QTY = df_desc_sindatos['QTY'].sum()
limit_SKU = len(df_desc_limitante)
limit_QTY = df_desc_limitante['QTY'].sum()
a1_SKU = len(df_a1)
a1_QTY = df_a1['QTY'].sum()
a2_SKU = len(df_a2)
a2_QTY = df_a2['QTY'].sum()
a3_SKU = len(df_a3)
a3_QTY = df_a3['QTY'].sum()
df_SKU = len(df)
df_QTY = df['QTY'].sum()

data = [['descarte-mes', mes_SKU, mes_QTY], ['descarte-cantidad', qty_SKU, qty_QTY], ['descarte-sindatos', sd_SKU, sd_QTY],
        ['descarte-limitefisico', limit_SKU, limit_QTY], ['Sobra', df_SKU, df_QTY], ['A1', a1_SKU, a1_QTY], ['A2', a2_SKU, a2_QTY], ['A3', a3_SKU, a3_QTY]]

df_var = pd.DataFrame(data, columns = ['Name', 'SKU', 'QTY']) 

df.to_excel("df.xlsx", sheet_name="Sobra")
with pd.ExcelWriter('df.xlsx', mode='a') as writer:
    df_var.to_excel(writer, sheet_name="Vars")
    df_a1.to_excel(writer, sheet_name="A1")
    df_a2.to_excel(writer, sheet_name="A2")
    df_a3.to_excel(writer, sheet_name="A3")
    df_desc_meses.to_excel(writer, sheet_name="out-meses")
    df_desc_QTY.to_excel(writer, sheet_name="out-cantidad")
    df_desc_sindatos.to_excel(writer, sheet_name="out-sindatos")
    df_desc_limitante.to_excel(writer, sheet_name="out-limitante")



print(a1_SKU, a2_SKU, a3_SKU, mes_SKU, qty_SKU, sd_SKU, limit_SKU, df_SKU)
