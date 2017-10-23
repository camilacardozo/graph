#-*- coding: cp1252 -*-


#importa as bibliotecas necessárias
import serial
import numpy
import matplotlib.pyplot as plt  
import pandas as pd
from pandas import DataFrame
import openpyxl
from drawnow import*

X=[] # guarda os valores coletados através do Arduino
Y=[] # guarda os valores coletados através do Arduino
dados = [] # guarda os valores coletados através do Arduino
arduinoData= serial.Serial('COM3', 115200)
#estabelece a comunicação serial com a porta COM3 (OBS: para outros sistemas operacionais que não o Windows, verificar o nome da porta)
plt.ion() #ativa o modo interativo no matplotlib
z=0  #inicia a variável x em 0
i=input("Nº de pontos plotados:")  #abre uma entrada de dados. ao rodar o programa, o usuário deverá inserir o número de pontos que deverão ser plotados, evitando um acúmulo desnecessário de dados

def makeFig():    #cria a função makeFig
    plt.plot(X, Y, 'b.:')  # constrói o gráfico de y(x)
    plt.grid(True)  #ativa a grade de fundo do gráfico, facilitando a visualização dos pontos
    plt.grid('r', linestyle='--', linewidth=1)  #seleciona a grade na cor vermelha, tracejada e de tamanho 1
    plt.title('y(t))')  #insere o título do gráfico
    plt.ylim(0)  #insere o limites inferior do eixo y
    plt.xlim(0)    #insere o limite inferior do eixo x
    plt.ylabel('y')  # insere um título para o eixo y


while True:   #inicia um loop
    while ( arduinoData.inWaiting()==0): #insere outro loop, colocando como condição a comunicação serial.
         pass  #caso não haja comunicação serial, não executa nenhuma tarefa
    arduinoString = arduinoData.readline() #define arduinoString como a leitura realizada pela porta serial (definida acima)
    print arduinoString #escreve os valores lidos na porta serial
    objeto = arduinoString.split(',') #separa o primeiro objeto em dois novos objetos, utilizando vírgula como critério de separação
    valorX = long (objeto[0])  #guarda o primeiro valor enviado via porta serial (índice 0)
    valorY =float(objeto[1]) #guarda o segundo valor enviado via porta serial (índice 1)
    X.append(valorX)  #insere os valores guardados em valorX dentro da variável x, contido fora do loop
    Y.append(valorY) #insere os valores guardados em valorY dentro da variável Y, contido fora do loop
    dados.append(arduinoString) #insere os valores coletados via porta serial dentro da variável dados
    z=z+1 #conta o número de interações com a porta serial
    if (z<=i):  #compara o nº de interações com o nºdados selecionados pelo usuário em i. Se x<=i, realiza a seguinte tarefa:
        df = pd.DataFrame(dados) #indere os valores guardados em dados dentro da função df
        df.to_excel('Aquisição.xlsx', 'Sheet1', index=False)  #gera um arquivo xlsx com os valores guardados em df
        drawnow(makeFig)  #executa a ferramenta drawnow, ou seja, atualiza o gráfico em cada interação da porta serial
    else: #caso z>i, realiza a seguinte tarefa:
        plt.show(makeFig)  #constrói o gráfico com os dados adquiridos até então, sem mais atualizá-lo
