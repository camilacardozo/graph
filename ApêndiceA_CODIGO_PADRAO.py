#-*- coding: cp1252 -*-


#importa as bibliotecas necess�rias
import serial
import numpy
import matplotlib.pyplot as plt  
import pandas as pd
from pandas import DataFrame
import openpyxl
from drawnow import*

X=[] # guarda os valores coletados atrav�s do Arduino
Y=[] # guarda os valores coletados atrav�s do Arduino
dados = [] # guarda os valores coletados atrav�s do Arduino
arduinoData= serial.Serial('COM3', 115200)
#estabelece a comunica��o serial com a porta COM3 (OBS: para outros sistemas operacionais que n�o o Windows, verificar o nome da porta)
plt.ion() #ativa o modo interativo no matplotlib
z=0  #inicia a vari�vel x em 0
i=input("N� de pontos plotados:")  #abre uma entrada de dados. ao rodar o programa, o usu�rio dever� inserir o n�mero de pontos que dever�o ser plotados, evitando um ac�mulo desnecess�rio de dados

def makeFig():    #cria a fun��o makeFig
    plt.plot(X, Y, 'b.:')  # constr�i o gr�fico de y(x)
    plt.grid(True)  #ativa a grade de fundo do gr�fico, facilitando a visualiza��o dos pontos
    plt.grid('r', linestyle='--', linewidth=1)  #seleciona a grade na cor vermelha, tracejada e de tamanho 1
    plt.title('y(t))')  #insere o t�tulo do gr�fico
    plt.ylim(0)  #insere o limites inferior do eixo y
    plt.xlim(0)    #insere o limite inferior do eixo x
    plt.ylabel('y')  # insere um t�tulo para o eixo y


while True:   #inicia um loop
    while ( arduinoData.inWaiting()==0): #insere outro loop, colocando como condi��o a comunica��o serial.
         pass  #caso n�o haja comunica��o serial, n�o executa nenhuma tarefa
    arduinoString = arduinoData.readline() #define arduinoString como a leitura realizada pela porta serial (definida acima)
    print arduinoString #escreve os valores lidos na porta serial
    objeto = arduinoString.split(',') #separa o primeiro objeto em dois novos objetos, utilizando v�rgula como crit�rio de separa��o
    valorX = long (objeto[0])  #guarda o primeiro valor enviado via porta serial (�ndice 0)
    valorY =float(objeto[1]) #guarda o segundo valor enviado via porta serial (�ndice 1)
    X.append(valorX)  #insere os valores guardados em valorX dentro da vari�vel x, contido fora do loop
    Y.append(valorY) #insere os valores guardados em valorY dentro da vari�vel Y, contido fora do loop
    dados.append(arduinoString) #insere os valores coletados via porta serial dentro da vari�vel dados
    z=z+1 #conta o n�mero de intera��es com a porta serial
    if (z<=i):  #compara o n� de intera��es com o n�dados selecionados pelo usu�rio em i. Se x<=i, realiza a seguinte tarefa:
        df = pd.DataFrame(dados) #indere os valores guardados em dados dentro da fun��o df
        df.to_excel('Aquisi��o.xlsx', 'Sheet1', index=False)  #gera um arquivo xlsx com os valores guardados em df
        drawnow(makeFig)  #executa a ferramenta drawnow, ou seja, atualiza o gr�fico em cada intera��o da porta serial
    else: #caso z>i, realiza a seguinte tarefa:
        plt.show(makeFig)  #constr�i o gr�fico com os dados adquiridos at� ent�o, sem mais atualiz�-lo
