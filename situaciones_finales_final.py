
import tabula 
import pandas as pd
from PyPDF2 import PdfFileReader

ruta2=str
nom_archivo=str
ruta_fin=str

data=pd.DataFrame()
pdf=input("Ingrese la ruta del archivo que desea convertir (incluyendo su extensión .pdf):\n")
print ("\n")
pdf2=PdfFileReader(open(pdf, "rb"))
pages=pdf2.numPages
print ("El numero de paginas es: ",pages)
print ("\n")
sf1=pd.DataFrame()
i=int()
# Recorre las paginas una a una y extrae los 3 primeros cuadros que se presentan en cada una
for i in range (1,pages+1,i+1): #La i destino debe ponerse +1 unidad cuando el rango comienza desde 1
  sf1 = tabula.read_pdf(pdf, lattice=True, pandas_options=({"header": None}), pages=i) #Multiple_tables=False
  #print (type(sf1))
  #contenido=len(sf1)
  #print (contenido)
  try:
    df1= pd.DataFrame(sf1[0])
    data=data.append(df1)
  except IndexError:
    pass

  try:
    df2= pd.DataFrame(sf1[1])
    data=data.append(df2)
  except IndexError:
    pass

  try:
    df3= pd.DataFrame(sf1[2])
    data=data.append(df3)
  except IndexError:
    pass

# print (data)
ruta2=input("Ingrese la ruta de la carpeta salida del nuevo archivo:\n")
print ("\n")
nom_archivo=input("ingrese el nombre del archivo de salida (sin extensión .xlsx):\n")
print ("\n")
ruta_fin = ruta2 + '\\' + nom_archivo + ".xlsx"
print (ruta_fin)
data.to_excel(ruta_fin)
print ("El archivo ha sido generado en la siguiente ubicación:\n",ruta_fin)