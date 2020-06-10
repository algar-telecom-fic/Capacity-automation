import re
import pandas as pd
import numpy as np

def getTuples(fileName):
    f = open(fileName)
    lines = f.readlines()
    i = 0 
    start = end = -1
    maxLines = len(lines) - 1
    tuples = []
    headers = []
    
    # Separa as tuplas 
    while i <= maxLines: 
        if lines[i].startswith('Maximum number'):
            start = i + 2

        elif lines[i].startswith('(Number of '):
            h = (re.split(r'\s{2,}', lines[start]))
            end = i - 1
            tup = [t.split() for t in lines[start+2:end+1]]
            
            # Verifica se possui campo "Module number"
            if len(h) == 4:
                k = 0
                
                for t in tup:
                    t.insert(2,None)
                    
            tuples.extend(tup)
            headers.append(h)

        i = i + 1
    
    i = 0 
    maxTuples = len(tuples) - 1
    
    # Converte de string pra float quando for digito
    while i <= maxTuples:            
        j = 0
        
        while j <= len(tuples[i]) - 1:
            if j == 1:
                tuples[i][j] = int(tuples[i][j])
                
            elif tuples[i][j] is not None and tuples[i][j].isdigit():
                tuples[i][j] = float(tuples[i][j])
            j = j + 1
                
        i = i + 1
        
    return tuples

def defineTableType(fileName, premise):
    tables = {}
    if premise == 'SPO':
        df = pd.read_excel(fileName, sheet_name='TABELAS SPO')
    elif premise == 'ULA':  
        df = pd.read_excel(fileName, sheet_name='TABELAS ULA')
    elif premise == 'FAC':
        df = pd.read_excel(fileName, sheet_name='TABELAS FAC')
    
    # Separa as tabelas marcadas como únicas (não consideram intervalo de módulo)
    unique = df.OBSERVAÇÃO.str.contains('TABELA ÚNICA')
    result = df[unique]
    
    tableIDS = list(result['Table ID'])
    for tableID in tableIDS: tables[tableID] = None
    
    # Separa as tabelas marcadas como duplicada sem módulo (não considera intervalo de módulo)
    notUnique = ~df.OBSERVAÇÃO.str.contains('TABELA ÚNICA')
    noModule = df.OBSERVAÇÃO.str.contains('NÃO POSSUI MÓDULO')
    result = df[notUnique]
    result = result[noModule]
    
    tableIDS = list(result['Table ID'])
    for tableID in tableIDS: tables[tableID] = None
        
    # Separa as tabelas marcadas como duplicada e com intervalo de módulo a considerar
    notUnique = ~df.OBSERVAÇÃO.str.contains('TABELA ÚNICA')
    withModule = df.OBSERVAÇÃO.str.contains('COM MÓDULO')
    result = df[notUnique]
    result = result[withModule]
    result = result[['Table ID', 'OBSERVAÇÃO']]
    
    for index, row in result.iterrows():
        startIntervalo = row['OBSERVAÇÃO'].find('DO MÓDULO')
        endIntervalo = row['OBSERVAÇÃO'].find('.')
        intervalo = row['OBSERVAÇÃO']
        intervalo = intervalo[startIntervalo+len('DO MÓDULO'):endIntervalo]

        startIntervalo = intervalo.find('AO')
        start = int(intervalo[:startIntervalo])
        end = int(intervalo[startIntervalo+3:])
        intervalo = (start,end)

        tables[row['Table ID']] = intervalo
        
    return tables

def createFinalCSV(tuplesFile, tablesFile, premise):
    tuples = getTuples(tuplesFile)
    df = pd.DataFrame(tuples, 
                      columns=['Table_name', 'Table_ID', 'Module_number', 'Maximum_tuple_number', 'Used_Number'])
    tables = defineTableType(tablesFile, premise)
    notFound = {}
    treated = {}
    data = []
    
    for index, row in df.iterrows():
        try:
            tableType = tables[row['Table_ID']]
            
            if row['Table_ID'] in treated: continue
                
            elif tableType is None:
                treated[row['Table_ID']] = True
                data.append(row.tolist())
                
            else:
                # Realiza tratativa nas tabelas que consideram intervalo de módulo
                treated[row['Table_ID']] = True
                start = tableType[0]
                end = tableType[1]
                
                result = df[(df.Table_ID == row['Table_ID']) & ((df.Module_number >= start ) & (df.Module_number <= end))]
                maxTupleNumber = result['Maximum_tuple_number'].sum()
                usedNumber = result['Used_Number'].sum()
                d = [row['Table_name'], row['Table_ID'], start, maxTupleNumber, usedNumber]
                data.append(d)
            
        except KeyError:
            # Salva os IDs não tratados
            notFound[row['Table_ID']] = True
    
    treatedCSV = pd.DataFrame(data, 
                              columns=['Table_name', 'Table_ID', 'Module_number', 'Maximum_tuple_number', 'Used_Number'])
    treatedCSV = treatedCSV.sort_values(by=['Table_ID'])
    treatedCSV = treatedCSV[['Table_ID','Table_name', 'Module_number', 'Maximum_tuple_number', 'Used_Number']]
    
    return treatedCSV

def renameColumns(df, columns): 
    newColumns = {}
    
    for c in columns:
        newColumns[c] = c.lower().replace(" ", "_")
    
    df = df.rename(columns=newColumns)

    return df

def generateCompleteReport(fileBase, fileActual, countDays, name):
    actual = fileActual
    base = fileBase
    base = renameColumns(base, list(base.columns))
    actual = renameColumns(actual, list(actual.columns))
    completeReport = []
    
    for index,row in actual.iterrows():
        rowBase = base.loc[base['table_id'] == row.table_id]
        usage = str(int(round(row.used_number/row.maximum_tuple_number, 2)*100)) + '%'
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
            
        newRow = [row.table_id, row.table_name, int(row.maximum_tuple_number), int(row.used_number), usage, realGrowth, forecast]
        completeReport.append(newRow)
    
    report = pd.DataFrame(completeReport, columns=['TABLE ID','TABLE NAME','CAPACIDADE','UTILIZADO','UTILIZADO %','CRESCIMENTO MENSAL','PREVISAO DE ESGOTAMENTO / MES'])

    report.to_csv(name, index=False)
    return report

def run(current_file_ULA, previous_file_ULA, name_csv_ULA, ULA_softx, days_ULA, current_file_SPO, previous_file_SPO, name_csv_SPO, SPO_softx, days_SPO, current_file_FAC, previous_file_FAC, name_csv_FAC, FAC_softx, days_FAC):
    current_ULA = createFinalCSV(current_file_ULA, 'Premissas Tabelas.xlsx', ULA_softx)
    previous_ULA = createFinalCSV(previous_file_ULA, 'Premissas Tabelas.xlsx', ULA_softx)
    generateCompleteReport(previous_ULA, current_ULA, days_ULA, name_csv_ULA)
    
    current_SPO = createFinalCSV(current_file_SPO, 'Premissas Tabelas.xlsx', SPO_softx)
    previous_SPO = createFinalCSV(previous_file_SPO, 'Premissas Tabelas.xlsx', SPO_softx)
    generateCompleteReport(previous_SPO, current_SPO, days_SPO, name_csv_SPO)

    current_FAC = createFinalCSV(current_file_FAC, 'Premissas Tabelas.xlsx', FAC_softx)
    previous_FAC = createFinalCSV(previous_file_FAC, 'Premissas Tabelas.xlsx', FAC_softx)
    generateCompleteReport(previous_FAC, current_FAC, days_FAC, name_csv_FAC)
    

if __name__ == '__main__':
    archive = open('arguments_tables.txt','r')
    line = []
    for lines in archive:
        lines = lines.strip()
        line.append(lines)
    archive.close()

    list_path = []
    for pos in line:
        count = 0
        for character in pos:
            count += 1
            if character == "=":
                list_path.append(pos[count+1:])

    run(list_path[0], list_path[1], list_path[2] ,'ULA', int(list_path[3]), list_path[4] , list_path[5], list_path[6], 'SPO', int(list_path[7]), list_path[8], list_path[9], list_path[10], 'FAC', int(list_path[11]))