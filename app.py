import pandas as pd
import os
import sys
import shutil

ruta = input("Ruta excel (sin comillas): ")

url_des = input("DIR destino (sin comillas): ")
url_root_gas = r'\\arfile01\buegadministracion\proveedores\FACTURAS DIGITALIZADAS'
url_root_gen = r'\\arfile01\buegadministracion\Proveedores\FACTURAS DIGITALIZADAS PROVEEDORES'

if not os.path.exists(url_des):
    sys.exit("ERROR: El directorio {} no existe".format(url_des))


df = pd.read_excel(ruta, dtype={"Clave de referencia":str, "Acreedor":str})
codigos = pd.read_excel(r'C:\Users\ypajarino\Projects\2021.02.03 BUSCADOR PDF\codigos enargas.xlsx',header=None, index_col=0,dtype={0:str,1:str})
codigos = codigos.to_dict()[1]
tipo_docs = ["KR", "KE", "KJ", "KH", "KC", "K2", "K3", "K5"]
dropin = df.loc[(df["Doc.compensación"] > 1800000000) &\
                (df["Doc.compensación"] < 1900000000) |\
                (~df["Clase de documento"].isin(tipo_docs)),:].index
df.drop(index = dropin, inplace = True)

df["Prod_gas"] = [1 if int(value) > 2000000 or int(value) in [1645, 7914] else 0 for index, value in df["Acreedor"].items()]
df["URL_año"] = df["Fe.contabilización"].dt.year.astype(str)
df["URL_soc"] = df["Sociedad"].replace({70:"CGP", 80:"CGS", 45:"CEN"})

# Proveedores generales
gen = df.loc[df["Prod_gas"] == 0,["URL_soc","Clave de referencia","Acreedor","URL_año"]]
gen["URL_doc"] = [value[:-8] if value[0] != "5" else value[:-4] for index, value in gen["Clave de referencia"].items()]
gen["URL"] = url_root_gen + '\\' + gen['URL_soc'] + '\\' + gen["Acreedor"].astype(str) + "-" + gen["URL_doc"] + "-" + gen["URL_año"] + ".pdf"

# Productores de gas, Transporte y GLP
gas = df.loc[df["Prod_gas"] == 1, ["Referencia","Fe.contabilización", "Sociedad", "Acreedor","URL_soc","URL_año"]]
gas["URL_mes"] = [value.zfill(2) for index, value in gas["Fe.contabilización"].dt.month.astype(str).items()]
gas["soc_enargas"] = gas["Sociedad"].replace({70:"20003", 80:"20004", 45:"00000"})
gas["prov_enargas"] = [codigos[value] for index, value in gas["Acreedor"].astype(int).items()]
gas["URL_acreedor"] = [value.zfill(10) for index, value in gas["Acreedor"].astype(str).items()]
gas["URL"] = url_root_gas + '\\' + gas["URL_soc"] + '\\' + gas["URL_año"] + '\\' + gas["URL_mes"] + '\\' + gas["URL_acreedor"] + '\\' + gas["soc_enargas"] + '_' + gas["prov_enargas"] + '_' + gas["Referencia"] + '_' + gas["URL_año"] + '-' + gas["URL_mes"] + '.pdf'

dirs = pd.concat([gen["URL"],gas["URL"]], ignore_index=True)
sin_pdf = []
n = dirs.shape[0]
for index, value in dirs.items():
    print("Copiando {} de {}".format(index + 1, n), end='')
    if not os.path.exists(value):
        print("{}".format(value), end="")
        print(" - No tiene PDF asociado")
        sin_pdf.append(value)
    else:
        print()
        shutil.copy(value,url_des)

print(pd.Series(sin_pdf).to_excel(os.path.join(url_des,"sin_pdf.xlsx")))