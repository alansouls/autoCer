#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys
import xlrd
import os
import commands





def getNames(file,number):
	book = xlrd.open_workbook(file)
	sh = book.sheet_by_index(number)
	nomes = []
	for i in range(1,sh.nrows):
		nomes.append(sh.cell_value(rowx=i, colx=0))
	matricula = []
	for i in range(1,sh.nrows):
		matricula.append(int(sh.cell_value(rowx=i, colx=2)))

	
	
	return nomes,matricula


def gerarPDF(name,file):
	command = 'inkscape --file=modificado.svg --export-area-drawing --without-gui --export-pdf='+name+'.pdf'
	os.system(command)

def preencherCertificado(nome,matricula,atividade,cargaHor,file):
	with open(file, 'r') as file:
		filedata = file.read()

	filedata = filedata.replace('_NOME_',nome)
	filedata = filedata.replace('_MAT_', str(matricula))
	filedata = filedata.replace('_ATIVIDADE_', atividade)
	filedata = filedata.replace('_TEMPO_', cargaHor)
	with open('modificado.svg', 'w') as file:
		file.write(filedata)

#example
nomes,matriculas = getNames('file.xlsx',0)
for i in range(len(nomes)):
	preencherCertificado(nomes[i].encode('utf-8'), matriculas[i],'Minicurso de Octave','4','certificadoModelo.svg')
	gerarPDF(str(matriculas[i])+'_0','modificado.svg')
#example
