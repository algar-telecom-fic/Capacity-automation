#Relatório de VLANs de BRAS MX Juniper (Expansão+Concessão) + NE40 Hauwei (Expansão)
#Script criado por Vitor Machado
import pkg_resources.py2_warn
from netmiko import Netmiko
from netmiko.ssh_autodetect import SSHDetect
from netmiko.ssh_exception import NetMikoAuthenticationException, NetMikoTimeoutException
import json
import pandas as pd
import openpyxl
import concurrent.futures
import time
import re

def Conectar_Equipamento(ip, device_type='-', comando='\n'):

    try:
        dados_conexao = {
            "host": ip,
            "username": usuario,
            "password": senha,
            "device_type": device_type,
            "banner_timeout": 100,
            "timeout": 100
        }

        if dados_conexao['device_type'] == '-':
            dados_conexao['device_type'] = 'autodetect'
            guesser = SSHDetect(**dados_conexao)
            best_match = guesser.autodetect()
            print('Equipamento {} é do tipo: {}'.format(dados_conexao['host'], best_match))
            dados_conexao["device_type"] = best_match
        net_connect = Netmiko(**dados_conexao)
        prompt = net_connect.find_prompt()

        # if device_type=='huawei':
        #     print(net_connect.disable_paging(command='screen-length 0 temporary'))

        info_equip = net_connect.send_command(comando).splitlines()

        return info_equip

    except (NetMikoAuthenticationException) as e:
        print('Falha ao conectar ao equipamento {}. {} '.format(ip, e))
        # output = ['', '-']

    except (NetMikoTimeoutException, Exception) as e:
        print('Falha ao conectar ao equipamento {}. {} '.format(ip, e))
        # output = ['', '-']
        #
    return None
    
start = time.time()

print('Digite seu usuário e senha de rede no arquivo .json que se encontra no mesmo diretório do script.')
input('Após, digite enter para iniciar: ')

with open('credenciais.json', 'r') as f:
    data_store = json.load(f)

usuario = data_store['usuario']
senha = data_store['senha']

diretorio = data_store['diretorio']

df_bras = pd.read_excel(diretorio+'bras-vendors.xlsx')
print(df_bras)

df_bras_info = pd.DataFrame(columns=['IP', 'VLAN', 'INTERFACE', 'SUBSCRIBERS'])
df_temp = df_bras_info.copy()

with concurrent.futures.ThreadPoolExecutor(max_workers=None) as executor:
    for info_equip, ip, device_type in zip(executor.map(Conectar_Equipamento, df_bras['IP'], df_bras['Device_Type'], df_bras['Comando']), df_bras['IP'], df_bras['Device_Type']):
        if device_type == 'juniper_junos':
            info_equip = ['Device: {}, SVLAN: {}'.format(info_equip[i][info_equip[i].find(': ')+2:], info_equip[i+1][info_equip[i+1].find('00.')+3:]) for info in range(0, len(info_equip), 2)]
        elif device_type == 'huawei':
            info_equip = [info for info in info_equip if ('GE0' in info or 'PPPoE' in info)]
            info_equip = ['Device: {}, SVLAN: {}'.format(info_equip[i][34:43], info_equip[i+1][10:14]) for i in range(0, len(info_equip), 2)]
        info_by_vlan = list((info, info_equip.count(info)) for info in set(info_equip))
        print(ip, dict(info_by_vlan))

        df_bras_info = df_bras_info.append({'IP': ip}, ignore_index=True)
        for vlan in info_by_vlan:
            df_temp = df_temp.append({'IP': ip, 'VLAN': vlan[0][-4:], 'INTERFACE': vlan[0][8:vlan[0].find(',')],'SUBSCRIBERS': vlan[1]}, ignore_index=True)
        df_temp.sort_values(["VLAN"], axis=0, ascending=True, inplace=True)
        df_bras_info = df_bras_info.append(df_temp)
        df_bras_info = df_bras_info.append({'IP': '', 'VLAN': '', 'INTERFACE': '', 'SUBSCRIBERS': ''}, ignore_index=True)
        df_temp.drop(df_temp.index, inplace=True)

print(df_bras_info)

df_bras_info.to_excel(diretorio+'VLANs_BRAS_todos.xlsx', index=None, header=True)

end = time.time()
print('Tempo de execução: {} segundos'.format((end - start)))
print('Busca Finalizada')
input('Digite enter para sair:')