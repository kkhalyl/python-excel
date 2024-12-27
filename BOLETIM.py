import pandas as pd
import math

boletim1 = pd.read_excel("Boletim.xlsx")
boletim2 = pd.read_excel("Boletim.xlsx", sheet_name="Boletim2")
boletim3 = pd.read_excel("Boletim.xlsx", sheet_name="Boletim3")

def negrito(text):
  return "*" + text + "*"

def formatar(x):
    if (x>1000000):
        x=x/1000000
        x=round(x, 2)
        x=str(x)+" mi"
    elif (x>1000):
        x=x/1000
        x=round(x, 2)
        x=str(x)+" mil"
    return str(x).replace('.',',');

#Entradas (data, icmsO, icmsH, diferença, porcentagem)

data=input('Insira a data de hoje. Exemplo: 01/01/2024 \n')

mes=(boletim1.iloc[1,1])

pICMSO=int(math.ceil(boletim1.iloc[2,1]))

pICMSH=int(math.ceil(boletim1.iloc[3,1]))

diferenca=int(boletim1.iloc[4,1])

difPorc=((diferenca/1000000)/pICMSO)*100

#Ranking Maiores Altas

maioresAltasL=[]

qtdA = 0

while (len(maioresAltasL) < 6):
    if pd.isna(boletim1.iloc[qtdA + 5, 1]) or str(boletim1.iloc[qtdA + 5, 1]) == "":
      break
    maioresAltasL.append(boletim1.iloc[qtdA+5,1])
    maioresAltasL.append(boletim1.iloc[qtdA+6,1])
    qtdA+=2;

#Ranking Maiores Quedas

maioresQuedasL=[]

qtdQ = 0

while ((len(maioresQuedasL) < 6)):
    if pd.isna(boletim1.iloc[qtdQ + 11, 1]) or str(boletim1.iloc[qtdQ + 11, 1]) == "":
      break
    maioresQuedasL.append(boletim1.iloc[qtdQ+11,1])
    maioresQuedasL.append(boletim1.iloc[qtdQ+12,1])
    qtdQ+=2;

#Crescimento Nominal

crescPorc=int(math.ceil(boletim1.iloc[17,1]))

crescDif=boletim1.iloc[18,1]

#Arrecadado ate o DIA

icmsP=boletim1.iloc[19,1]
icmsA=boletim1.iloc[20,1]
ipvaP=boletim1.iloc[21,1]
itcmdP=boletim1.iloc[22,1]
taxas=boletim1.iloc[23,1]
irrf=boletim1.iloc[24,1]

#Por Setor

industria=boletim3.iloc[0,1]

comercio=boletim3.iloc[1,1]

servico=boletim3.iloc[2,1]

valorT=(industria+comercio+servico)
industriaP=(industria/valorT)*100
comercioP=(comercio/valorT)*100
servicoP=(servico/valorT)*100

#Inadimplencia
inadICMS=boletim2.iloc[0,1]
inadICMSP=boletim2.iloc[1,1]
inadIPVA=boletim2.iloc[2,1]
inadIPVAP=boletim2.iloc[3,1]
inadITCMD=boletim2.iloc[4,1]
inadITCMDP=boletim2.iloc[5,1]

#Expeculação

pibT=int(input("Insira o PIB total: "))
cambio=int(input("Insira o câmbio: "))
selic=int(input("Insira a taxa Selic: "))
ipca=int(input("Insira o IPCA: "))
dataF=negrito(input("Insira a data do relatório FOCUS: "))

#Print

print("📊", negrito("Boletim GANS/DEARC"), "\n\n")
print("Destaques de hoje:", data ,"\n")
print("🔭", negrito("PREVISÃO ICMS:"), "🔭\n")

print("  Ontem              Hoje  \n")
print("R$ ", pICMSO/1000, " bi      R$ ", pICMSH/1000, "bi\n")


if (diferenca>0):
    print(f"🟢  *{difPorc:,.2f}*% ", "(R$" ,str(abs(round(diferenca/1000000,2))).replace('.',','), "mi)")
elif (diferenca<0):
    print(f"🔴  *{difPorc:,.2f}*% ", "(R$" ,str(abs(round(diferenca/1000000,2))).replace('.',','), "mi)")
else:
    print(f"🟰  *{difPorc:,.2f}*% ", "(R$" ,str(abs(round(diferenca/1000000,2))).replace('.',','), "mi)")

#Print Maiores Altas
if (qtdA==2):
    print("\n📈",  negrito("Maior Alta:"), "📈\n")
else:
    print("\n📈",  negrito("Maiores Altas:"), "📈\n")

while (qtdA>0):
    if (qtdA==6):
        print("🥇 ", maioresAltasL[0], ": R$", formatar(maioresAltasL[1]))
        print("\n🥈 ", maioresAltasL[2], ": R$", formatar(maioresAltasL[3]))
        print("\n🥉 ", maioresAltasL[4], ": R$", formatar(maioresAltasL[5]))
        qtdA-=6
    if (qtdA==4):
        print("🥇 ", maioresAltasL[0], ": R$", formatar(maioresAltasL[1]))
        print("\n🥈 ", maioresAltasL[2], ": R$", formatar(maioresAltasL[3]))
        qtdA-=4
    if (qtdA==2):
        print("🥇 ", maioresAltasL[0], ": R$", formatar(maioresAltasL[1]))
        qtdA-=2

#Print Maiores Quedas
if (qtdQ==1):
    print("\n\n📉", negrito("Maior Queda:"), "📉\n")
else:
    print("\n\n📉", negrito("Maiores Quedas:"), "📉\n")

while (qtdQ>0):
    if (qtdQ==6):
        print("🥇 ", maioresQuedasL[0], ": R$", formatar(maioresQuedasL[1]))
        print("\n🥈 ", maioresQuedasL[2], ": R$", formatar(maioresQuedasL[3]))
        print("\n🥉 ", maioresQuedasL[4], ": R$", formatar(maioresQuedasL[5]))
        qtdQ-=6
    if (qtdQ==4):
        print("🥇 ", maioresQuedasL[0], ": R$", formatar(maioresQuedasL[1]))
        print("\n🥈 ", maioresQuedasL[2], ": R$", formatar(maioresQuedasL[3]))
        qtdQ-=4
    if (qtdQ==2):
        print("🥇 ", maioresQuedasL[0], ": R$", formatar(maioresQuedasL[1]))
        qtdQ-=2

print("\n========================")

print("\n",negrito("💰Previsão"),mes,negrito("/24 x"), mes, negrito("/23 - ICMS PRINCIPAL"), "\n")

if (crescPorc>0):
    print ("\nCrescimento nominal de", str(crescPorc/100).replace('.',','),"% (R$", formatar(crescDif),")")
if (crescPorc<0):
    print ("\nDecréscimo nominal de", str(crescPorc/100).replace('.',','),"% (R$", formatar(crescDif),")")

print("\n========================")

print("\n", negrito("💰ARRECADADO ATÉ O DIA💰"))

print("\n🛍️ ICMS Principal:    R$ ",formatar( icmsP) )

print("\n🛒 Adicional ICMS:     R$ ", formatar( icmsA))

print("\n🚙 IPVA Principal:     R$ ", formatar( ipvaP))

print("\n⚰️ ITCMD Principal:    R$ ", formatar( itcmdP))

print("\n🪙 TAXAS:  R$ ", formatar( taxas))

print("\n🦁 IRRF:   R$ ", formatar( irrf))

print("\n========================")

print("\n", negrito("🧩 ARRECADAÇÃO POR SETOR: 🧩"))

print("\n🏭 Indústria:  R$", formatar( industria), str(f"({industriaP:.2f}%)").replace('.',','))

print("\n🏪 Comércio:  R$", formatar( comercio), str(f"({comercioP:.2f}%)").replace('.',','))

print("\n⚒️ Serviço:  R$", formatar(servico), str(f"({servicoP:.2f}%)").replace('.',','))

print("\n========================")

print("\n", negrito("💢INADIMPLÊNCIAS💢"))

print("\n⛔ ICMS:    R$", formatar( inadICMS), str(f"({inadICMSP/100:.2f}%)").replace('.',','))

print("\n⛔ IPVA:   R$", formatar( inadIPVA), str(f"({inadIPVAP/100:.2f}%)").replace('.',','))

print("\n⛔ ITCMD:    R$", formatar( inadITCMD), str(f"({inadITCMDP/100:.2f}%)").replace('.',','))

print("\n========================")

print("\n", negrito("🏹 EXPEC. REL. FOCUS 2024 🏹"))

print("\n🎯 PIB Total:  ", f"{pibT/100:.2f}%")

print("\n🎯 Câmbio: R$", f"{cambio/100:.2f}")

print("\n🎯 Selic:  ", f"{selic/100:.2f}%")

print("\n🎯 IPCA:   ", f"{ipca/100:.2f}%")

print("\n", negrito("🔖 Fonte: FOCUS"), dataF)
