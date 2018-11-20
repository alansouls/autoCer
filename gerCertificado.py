import sys
import xlrd
import os
import command


def getNames(file,number):
	book = xlrd.open_workbook(file)
	sh = book.sheet_by_index(number)
	nomes = []
	for i in range(1,sh.nrows):
		nomes.append(sh.cell_value(rowx=i, colx=0))
	matricula = []
	for i in range(1,sh.nrows):
		matricula.append(sh.cell_value(rowx=i, colx=2))

	return nomes, matricula


def gerarPNG(matricula,file,atividade):
	if str(matricula).endswith('.0'):
		matricula = str(matricula)[:-2]
	output = str(matricula)+'-1.png';
	os.system("inkscape --file="+file+" --export-area-drawing --without-gui --export-png="+output);

def preencherCertificado(nome,matricula,atividade,cargaHor,file):
	with open(file, 'r') as file:
		filedata = file.read()
	if str(matricula).endswith('.0'):
		matricula = str(matricula)[:-2]
	filedata = filedata.replace('_NOME_',nome)
	filedata = filedata.replace('_MAT_', str(matricula))
	filedata = filedata.replace('_ATIVIDADE_', atividade)
	filedata = filedata.replace('_TEMPO_', str(cargaHor))
	with open('modificado.svg', 'w') as file:
		file.write(filedata)


nomes,matricula = getNames('frequencia.xlsx',0);
for i in range(len(nomes)):
	preencherCertificado(nomes[i],matricula[i],'Minicurso Octave',4,'certificadoModelo.svg')
	gerarPNG(matricula[i],'modificado.svg','Minicurso Octave')




