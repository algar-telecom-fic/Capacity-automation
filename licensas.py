# -*- coding: utf-8 -*-
# -*- coding: cp1252 -*-
import pandas as pd
import re
import numpy as np

def getData(dataFile):
    fo = open(dataFile)
    lines = fo.readlines()
    lengthFile = len(lines)
    body = []
    headers = []
    splitBody = []
    
    for i in range(lengthFile):
        #cria a variavel onde o programa deve começar a ler os dados (start)
        if(lines[i].startswith('License Usage')):
            start = i+2
        
        #cria a variavel onde o programa deve parar de ler os dados (end)
        elif(lines[i].startswith('(Number of')):
            h = (re.split(r'\s{2,}', lines[start]))       #splita o cabeçalho
            end = i - 1
            
            #Faz o split do corpo do arquivo
            for t in range(start+2, end+1):
                splitLine = (re.split(r'\s{2,}', lines[t][1:]))
                splitBody.append(splitLine)   #vetor de linhas splitadas
        
            headers.append(h)     #vetor de cabeçalhos (nesse caso temos apenas um cabeçalho)
            body.extend(splitBody)  #vetor final que guarda todo o arquivo splitado
            
            #trocando os espaços da coluna License Item por Underscore "_"
            # for t in range(len(body)):
            #     body[t][1] = body[t][1].replace(" ", "_")
            break;

    # Tratando a string vazia qua aparece no final (posição 7 do vetor interno)
    for t in body:
        del t[7]
        del t[6]

    for t in range(len(body)):
        for ti in range(len(body[t])):
            if(ti > 2 and (body[t][ti] != '' and body[t][ti] != '-')):
                body[t][ti] = float(body[t][ti])
			    

    return body


def createCsv(File):
    tuples = getData(File)
    #df = pd.DataFrame(tuples, columns = [' License ID', 'License Item', 'Type', 'Authorization-values', 'Real-values', 'Usage-percent(%)'])
    #df = df.sort_values(by=['Real-values'])
    
    df = pd.DataFrame(tuples, columns = ['License ID', 'License Item', 'Type', 'maximum_tuple_number', 'used_number', 'Usage-percent(%)'])
    return df


def renameColumns(df, columns): 
    newColumns = {}
    
    for c in columns:
        newColumns[c] = c.lower().replace(" ", "_")
    
    df = df.rename(columns=newColumns)

    return df

def generateCompleteReport(fileBase, createdFile, countDays, name_csv):
    created = createdFile
    created = renameColumns(created, list(created.columns))
    
    base = fileBase
    base = renameColumns(base, list(base.columns))

    completeReport = []
    
    for index,row in created.iterrows():
        rowBase = base.loc[base['license_id'] == row.license_id]
        if(row.maximum_tuple_number == 0):
            usage = 0
        else:
            usage = str(int(round(row.used_number/row.maximum_tuple_number, 2)*100)) + '%'
            
        # monthlyGrowth = int(round(((row.used_number - rowBase.used_number.values[0])/countDays)))
        realGrowth = int(round(((row.used_number - rowBase.used_number.values[0])/countDays)*30))
        
        available = round(row.maximum_tuple_number - row.used_number)
        if realGrowth != 0:
            forecastNumber = round(available/realGrowth)  
        else:
            forecastNumber = 0
        
        if forecastNumber == 0: forecast = 'Estavel'
        elif forecastNumber >= 1 and forecastNumber <= 24: forecast = forecastNumber
        elif forecastNumber > 24: forecast = 'Maior que 2 Anos'
        elif forecastNumber < 0: forecast = 'Decrescimento'
            
        newRow = [row.license_id, row.license_item, row.type, int(row.maximum_tuple_number), int(row.used_number), usage, realGrowth, forecast]
        completeReport.append(newRow)
    
    report = pd.DataFrame(completeReport, columns=['ID', 'DESCRIPTION', 'TYPE', 'CAPACIDADE', 'UTILIZADO', 'UTILIZADO %','CRESCIMENTO MENSAL','PREVISAO DE ESGOTAMENTO / MES'])

    report.to_csv(name_csv, index=False)
    
    return report



def run(current_file_ULA, previous_file_ULA, name_csv_ULA, days_ULA, current_file_SPO, previous_file_SPO, name_csv_SPO, days_SPO, current_file_FAC, previous_file_FAC, name_csv_FAC, days_FAC):
    current_ULA = createCsv(current_file_ULA)
    previous_ULA = createCsv(previous_file_ULA)
    generateCompleteReport(previous_ULA, current_ULA, days_ULA, name_csv_ULA)
    
    current_SPO = createCsv(current_file_SPO)
    previous_SPO = createCsv(previous_file_SPO)
    generateCompleteReport(previous_SPO, current_SPO, days_SPO, name_csv_SPO)

    current_FAC = createCsv(current_file_FAC)
    previous_FAC = createCsv(previous_file_FAC)
    generateCompleteReport(previous_FAC, current_FAC, days_FAC, name_csv_FAC)
	

if __name__ == '__main__':

    archive = open('arguments_licenses.txt','r')
    list_path = []
    for lines in archive:
        lines = lines.strip()
        count = 0
        for character in lines:
        	count += 1
        	if character == "=":
        		list_path.append(lines[count+1:])
    archive.close()
	
    run(list_path[0], list_path[1], list_path[2] , int(list_path[3]), list_path[4], list_path[5], list_path[6] , int(list_path[7]), list_path[8], list_path[9], list_path[10] , int(list_path[11]))
