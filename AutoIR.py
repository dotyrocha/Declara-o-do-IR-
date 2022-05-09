# Importar os Modulos 
from numpy import singlecomplex
from pandas.core.indexing import convert_to_index_sliceable
import pyautogui
import pandas as pd
from time import sleep


pyautogui.alert('O PROGRAMA VAI COMEÇAR NÃO MEXA EM NADA !')
# Abrir o programa
pyautogui.press('winleft')
sleep(0.5)
pyautogui.write('IRPF')
sleep(0.5)
pyautogui.press('enter')
sleep(10)
pyautogui.press('esc')
sleep(0.5)
pyautogui.hotkey('ctrl', 'n')
pyautogui.click(1497,619)
pyautogui.press('tab')
sleep(0.5)
pyautogui.write('45669262816')
sleep(0.5)
pyautogui.press('tab')
sleep(0.5)
pyautogui.write('Nathalia Eulita Oliveira Silva Rocha')
pyautogui.press('tab')
sleep(1)

# Informação do Contibuinte 
if '' == '': 
    pyautogui.click(1727,727)
    pyautogui.press('enter')
    sleep(2)
    pyautogui.click(491,275)
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write('14061997')
    pyautogui.press('tab')
    pyautogui.write('1234567890')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write('estrada')
    pyautogui.press('tab')
    pyautogui.write('Estrada do Aladino')
    pyautogui.press('tab')
    pyautogui.write('18')
    pyautogui.press('tab')
    pyautogui.write('A')
    pyautogui.press('tab')
    pyautogui.write('Fincos')
    pyautogui.press('tab')
    pyautogui.write('SP')
    pyautogui.press('tab')
    pyautogui.write('Sao Bernardo do Campo')
    pyautogui.press('tab')
    pyautogui.write('09831525')
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write('11')
    pyautogui.press('tab')
    pyautogui.write('948326478')
    pyautogui.press('tab')
    pyautogui.write('11')
    pyautogui.press('tab')
    pyautogui.write('985618719')
    pyautogui.press('tab')
    pyautogui.write('hadassaeulita@hotmail.com')
    pyautogui.press('tab')
    pyautogui.write('01')
    pyautogui.press('tab')
    pyautogui.write('420')
    pyautogui.press('enter')
sleep(1)


# BENS E DIREITOS
pyautogui.click(63,597)


# PANDAS(ATIVOS)
planilia = 'D:\Documentos\Financas\IRPF\Declaração IR.xlsx'
df_1 = pd.read_excel(planilia, sheet_name = 'ativos')


# Variaveis ATIVOS
codigo_1 = df_1.codigo
corretora_1 = df_1.corretora
ativo_1 = df_1.ativo
mercado_1 = df_1.mercado
qtd_1 = df_1.qtd
cnpj_1 = df_1.cnpj
medio_1 = df_1.medio
atual_1 = df_1.atual
anterior_1 = df_1.anterior


# Adicionando ATIVOS
for codigo, corretora, ativo, mercado, qtd, cnpj, medio, atual, anterior, in zip(codigo_1, corretora_1, ativo_1, mercado_1, qtd_1, cnpj_1, medio_1, atual_1, anterior_1,):
    
    pyautogui.click(2291,974)
    sleep(0.5)
    pyautogui.write(str(codigo))

    
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(str(cnpj))
    pyautogui.press('tab')
    
    # Descriminação 
    if codigo == 45:
        pyautogui.write(f'Ativo {ativo} custodiado no {corretora}.')
        pyautogui.press('tab')
        pyautogui.write(str(anterior).replace('.',','))
        pyautogui.press('tab')
        pyautogui.write(str(atual).replace('.',','))
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('enter')
        sleep(1)

    else:
        pyautogui.write(f'Ativo {ativo} custodiado na corretora {corretora}, ao valor medio de R$ {(str(medio))}.')
        pyautogui.press('tab')
        pyautogui.write(str(anterior).replace('.',','))
        pyautogui.press('tab')
        pyautogui.write(str(atual).replace('.',','))
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('enter')
        sleep(1)

# PANDAS(BENS)
planilia = 'D:\Documentos\Financas\IRPF\Declaração IR.xlsx'
df_2 = pd.read_excel(planilia, sheet_name = 'bens')

# Variaveis BENS 
codigo_2 = df_2.codigo_b
tipo_2 = df_2.tipo_b
renavam_2 = df_2.renavam_b
medio_2 = df_2.medio_b
atual_2 = df_2.atual_b
anterior_2 = df_2.anterior_b


# Adicionando BENS
for codigo_b, tipo_b, renavam_b, medio_b, atual_b, anterior_b in zip(codigo_2, tipo_2, renavam_2, medio_2, atual_2, anterior_2):

    pyautogui.click(2291,974)
    sleep(0.5)
    pyautogui.write(str(codigo_b))

    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.write(str(renavam_b))
    pyautogui.press('tab')
    
    # Descriminação 
    pyautogui.write(f'{tipo_b} comprado ao valor de R$ {medio_b},00.')
    
    pyautogui.press('tab')
    pyautogui.write(str(anterior_b).replace('.',','))
    pyautogui.press('tab')
    pyautogui.write(str(atual_b).replace('.',','))
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('enter')
    sleep(1)

pyautogui.click(319,57)
pyautogui.alert('Corriga as pendecias para continuar ')