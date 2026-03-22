
import requests
import pandas as pd
import time

ARQUIVO_ENTRADA = r"C:\Users\Bruno Melo\Downloads\clientes (1).xlsx"
ARQUIVO_SAIDA   = "enderecos_preenchidos.xlsx"
COLUNA_CEP      = "CEP"   

def buscar_cep(cep):
    cep_limpo = str(cep).replace("-", "").replace(".", "").replace(" ", "").zfill(8)

    if len(cep_limpo) != 8 or not cep_limpo.isdigit():
        return {'erro':'CEP Inválido'}
    
    try:
        url= f'https://viacep.com.br/ws/{cep_limpo}/json/'
        resposta = requests.get(url, timeout=5)
        dados = resposta.json()
        return dados
    except Exception as e:
        return{'erro':str(e)}
    

df = pd.read_excel(ARQUIVO_ENTRADA)
df[COLUNA_CEP] = df[COLUNA_CEP].astype(str)

total = len(df)
print(f'Planilha carregada: {total} linhas\n')

logradouros, bairros, cidades, estados, ceps_fmt, status_list = [], [], [], [], [], []

for i, cep in enumerate(df[COLUNA_CEP], start=1):
    print(f"   [{i}/{total}] CEP: {cep}", end=" → ")

    dados = buscar_cep(cep)

    if "erro" in dados or dados.get("erro") == True:
        print("❌ não encontrado")
        logradouros.append(None)
        bairros.append(None)
        cidades.append(None)
        estados.append(None)
        ceps_fmt.append(cep)
        status_list.append("Não encontrado")
    else:
        cep_fmt = f"{dados['cep']}"
        print(f"✅ {dados.get('localidade')} / {dados.get('uf')}")
        logradouros.append(dados.get("logradouro"))
        bairros.append(dados.get("bairro"))
        cidades.append(dados.get("localidade"))
        estados.append(dados.get("uf"))
        ceps_fmt.append(cep_fmt)
        status_list.append("OK")

    time.sleep(0.1)   # pausa leve para não sobrecarregar a API

#df[COLUNA_CEP] = cep_fmt
df['Logradouro'] = logradouros
df["Bairro"]     = bairros
df["Cidade"]     = cidades
df["Estado"]     = estados
df["Status CEP"] = status_list

df.to_excel(ARQUIVO_SAIDA, index=False)

ok = status_list.count("OK")
erros = status_list.count("Não encontrado")

print(f"\n✅ Arquivo salvo: {ARQUIVO_SAIDA}")
print(f"   Endereços preenchidos : {ok}")
print(f"   CEPs não encontrados  : {erros}")