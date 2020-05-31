import openpyxl
import datetime

def change_value(str, new_value, sheet):
	sheet[str] = sheet[str].value.format(new_value)
	
def remake_shablon(shifr, name, diametr, tolshina, typ, number, quality):
	dirname = 'app/static/shablon.xlsx'
	wb = openpyxl.load_workbook(dirname)
	sheet= wb['sheet1']
	data= datetime.date.today().strftime("%d-%m-%Y")
	parametr1='≤' + str(int(tolshina)*0.1)
	parametr2='≤' + str(int(tolshina)*0.2)
	
	## Меняем шифр в шаблоне ##
	change_value('O46', shifr, sheet)
	
	## Меняем наименование объекта ##
	change_value('F48', name, sheet)
	
	## Меняем диаметр трубы ##
	change_value('O51', diametr, sheet)
	
	if int(diametr) < 300:
		sheet['A56'] = sheet['A56'].value.replace('x1', 'по всей длине шва.')
	else:
		sheet['A56'] = sheet['A56'].value.replace('x1', 'участка сварного шва после ремонта должна превышать длину отремонтированного участка на 100 мм в обе стороны.')
	if int(diametr) < 720:
		sheet['J83'] = sheet['J83'].value.format('.-')
	else:
		sheet['J83'] = sheet['J83'].value.format('8.-10')
	
	## Меняем толщину стенки ##
	change_value('O52', tolshina, sheet)
	
	## ТРЕБОВАНИЯ К ВЫПОЛНЕНИЮ ВИЗУАЛЬНОГО И ИЗМЕРИТЕЛЬНОГО КОНТРОЛЯ ##
	if int(tolshina) > 20:
		sheet['A56'] = sheet['A56'].value.replace('x2', tolshina)
	else:
		sheet['A56'] = sheet['A56'].value.replace('x2', '20')
	
	## Меняем глубину подреза ##
	if int(tolshina)*0.1 >= 0.5:
		sheet['J86'] = sheet['J86'].value.format('≤ 0,5')
	else:
		sheet['J86'] = sheet['J86'].value.format(parametr1) 
	
	## Глубина вогнутости корня (для сварных швов без подварки) ##
	if quality == 'A':
		if int(tolshina)*0.1 >= 1:
			sheet['J88'] = sheet['J88'].value.format('≤1')
		else:
			sheet['J88'] = sheet['J88'].value.format(a)
	if (quality == 'B') or (quality =='C'):
		if int(tolshina)*0.1 >= 2:
			sheet['J88'] = sheet['J88'].value.format('≤2')
		else:
			sheet['J88'] = sheet['J88'].value.format(parametr2)
	
	## Меняем номер удостоверения ##
	change_value('N115', number, sheet)
	
	## Меняем тип сварного соединения ##
	change_value('O54', typ, sheet)

	## Ширина облицовочного шва ##
	if typ == 'РД':
		if 6 < int(tolshina) <= 8:
			sheet['J82']= '11-18'
		elif 8.1 < int(tolshina) <= 12:
			sheet['J82']= '14-24'
		elif 12.1 < int(tolshina) <= 15:
			sheet['J82']= '18-28'
		elif 15.1 < int(tolshina) <= 20:
			sheet['J82']= '15-27'
		elif 20.1 < int(tolshina) <= 24:
			sheet['J82']= '18-31'
		elif 24.1 < int(tolshina) <= 27:
			sheet['J82']= '21-35'
		elif int(tolshina)>27:
			sheet['J82']= '21-35'
	if (typ == 'РД,МП+АПИ') or (typ == 'РД,МП+АПИ,МП+МПС+АФ'):
		if 6 < int(tolshina) <= 8:
			sheet['J82']= '11-18 для РД     11-17 для авт. Сварки'
		elif 8.1 < int(tolshina) <= 12:
			sheet['J82']= '14-24 для РД     16-24 для авт. Сварки'
		elif 12.1 < int(tolshina) <= 16:
			sheet['J82']= '18-28 для РД     19-27 для авт. Сварки'
		elif 16.1 < int(tolshina) <= 20.5:
			sheet['J82']= '15-27 для РД     20-28 для авт. Сварки'
		elif 20.6 < int(tolshina) <= 27:
			sheet['J82']= '18-31 для РД     22-30 для авт. Сварки'
		elif int(tolshina)>27:
			sheet['J82']= '21-35 для РД     22-30 для авт. Сварки'
	
	## Меняем уровень качества ##
	change_value('O50', quality, sheet)
	
	## Меняем дату ##
	change_value('L115', data, sheet)
	
	## Сохраняем новый ексель файл ##
	wb.save(dirname.replace('shablon', str(shifr)))