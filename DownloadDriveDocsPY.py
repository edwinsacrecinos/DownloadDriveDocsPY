import gdown
import aspose.words as convertPDF
import os as ManipulationFiles
from asyncio.log import logger
import pandas as pd

file_loc = "docsModificadoLinks.xlsx"

Id = pd.read_excel(file_loc, usecols="A")
doc1 = pd.read_excel(file_loc, usecols="R")



# 1 = image  2 = pdf 3= doc -> typeDoc 
def isNaN (value,path,nameFile,typeFile,typeDoc):
    if value != value:
        print('Sin archivo -> '+path+nameFile+typeFile)
    else:
        if typeDoc==1:
           
            
            doc = convertPDF.Document()
            builder = convertPDF.DocumentBuilder(doc)
            url = 'https://drive.google.com/uc?id='+value
            output = path+nameFile+typeFile
            gdown.download(url, output, quiet=False)
            builder.insert_image(path+nameFile+typeFile)
            doc.save(path+nameFile+".pdf")
            ManipulationFiles.remove(path+nameFile+typeFile)
            print("Descarga y conversión completada de -> "+path+nameFile+typeFile)
        elif typeDoc==2:
            
            url = 'https://drive.google.com/uc?id='+value
            output = path+nameFile+typeFile
            gdown.download(url, output, quiet=False)
            print("Descarga completada de -> "+path+nameFile+typeFile)
            

nameDocs = [
    'DOC' #doc1
   
]

index = 0
# len(Id)
print(":::::::::::::::::::::::::::::::::::::::::::::::::::::")
print("::::::::::::::::: INICIANDO SCRIPT ::::::::::::::::::")
print("::::::::::::::::::::::::::::::::::::::::::::::::::::: \n")
while index < 10:  
    try:
        # Crear el directorio exist_ok=True sobrescribe los directorios sin borrar nada y no da error
        paths= 'Descarga/'+str(Id['PROYECTO_ID'][index])+'/Documentos/'
        ManipulationFiles.makedirs(paths, exist_ok=True)
        print(":::::::::::::::::::: FILA "+str(index+1)+' ::::::::::::::::::::::::::')
        #value (linkDocsDrive),path,nameFile,typeFile,typeDoc,
        
        isNaN(doc1['NOMBRE COLUMNA'][index],paths,nameDocs[0],'.pdf',2)
       
        

    except Exception as e:
        logger.error('Error al descargar el documento en el directorio -> '+str(Id['PROYECTO_ID'][index])+ ' Error tipo: '+str(e))
    index = index + 1

    
print("Descargas finalizadas :::: Script Finish ::::")

