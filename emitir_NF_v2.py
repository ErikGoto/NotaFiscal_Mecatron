import pyautogui as pya
import time
import pandas


def le_plan(num_projeto, inputParcela):
    plan = pandas.read_excel(r"Nota_fiscal.xlsx", converters={'CPF/CNPJ cliente':str, 'Total':str, 'Parcela 1':str,'Parcela 2':str,'Parcela 3':str,'Parcela 4':str,'Parcela 5':str,'Parcela 6':str,'Parcela 7':str,'Parcela 8':str,'Parcela 9':str,'Parcela 10':str,'Parcela 11':str,'Parcela 12':str})
    cpf = str(plan['CPF/CNPJ cliente'][num_projeto])
    parcelas = plan['Total'][num_projeto]
    descricao = str(plan['Descrição Projeto'][num_projeto])
    valor = str(plan['Parcela ' + str(inputParcela)][num_projeto])
    return cpf, parcelas, descricao, valor, plan

projeto = int(input("Qual o código do projeto? "))
inputParcela = int(input("Qual o a parcela da NF? "))

cpf, parcelas, descricao, valor, plan = le_plan(projeto, inputParcela)

print(cpf)
print(str(inputParcela) + " de " + parcelas)
print(descricao)
print(valor)

#---------------------------------------------------------------------------------------------------------------------------------------



#---------------------------------------------------------------------------------------------------------------------------------------


#Clica na caixinha do CPF
try:    
    pya.click('imagens/cpf.png')
except:
    print("Erro: Caixa de CPF não encontrada")
    pya.alert(text='', title='Finalizado', button='OK')
    quit()
    
time.sleep(1.5)
pya.write(cpf[:], 0.25)
pya.move(50, 0)
pya.click()
time.sleep(1.5)

#Botão avançar
pya.scroll(700)
time.sleep(0.5)
pya.scroll(-10)
time.sleep(0.5)
try:    
    pya.click('imagens/avancar_1.png')
except:
    print("Erro: Caixa de Avançar não encontrada")
    pya.alert(text='', title='Finalizado', button='OK')
    quit()


#Selecionar o serviço de engenharia
time.sleep(1.5)
pya.moveTo(700, 596)
pya.click()
time.sleep(0.25)
pya.move(0, 60) 
pya.click()
time.sleep(2)

#Localidade da prestação de serviço
pya.scroll(-10)
time.sleep(1.5)
try:
    pya.click('imagens/local.png')
except:
    print("Erro: Caixa de localidade não encontrada")
    pya.alert(text='', title='Finalizado', button='OK')
    quit()
    

pya.move(0, 45)
pya.click()
time.sleep(1.5)
pya.scroll(-10)
pya.moveTo(815, 408)
time.sleep(0.25)
pya.click()
time.sleep(1.5)

#Última página
pya.scroll(-6)
time.sleep(0.25)
pya.click()
pya.write("Nota referente a parcela: " + str(inputParcela) + " de " + parcelas + "\n" + descricao + "\n" + "Associacao civil sem fins lucrativos, de carater estudantil, isenta de retencao na fonte quando das contribuicoes PIS, CONFINS e CSSL. Conforme Lei 9.532/1997 quando desenvolve atividades fins estatutarias / nao esta sujeita a retencao previdenciaria prevista in rfb 971/2009 por nao se tratar de cessao de mao-de-obra / nao estao sujeitas a desconto do imposto de renda na fonte sobre rendimentos pela prestação de servicos pela atividade fim conforme in rfb 23/1986")
time.sleep(0.25)
try:
    pya.click('imagens/projeto.png')
except:
    print("Erro: Caixa de Item não encontrada")
    pya.alert(text='', title='Finalizado', button='OK')
    quit()
    
pya.write("Projeto")
pya.move(175, 0)
pya.click()
pya.write("1")
pya.move(100, 0)
pya.click()
pya.write(valor)
pya.press("tab")

pya.alert(text='', title='Finalizado', button='OK')