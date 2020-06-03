import openpyxl
from openpyxl.drawing.image import Image
import datetime

dirname = 'app/static/shablon.xlsx'
wb = openpyxl.load_workbook(dirname)
data= datetime.date.today().strftime("%d-%m-%Y")

def change_value(str, new_value, sheet):
	sheet[str] = sheet[str].value.format(new_value)

def vik(shifr, name, diametr, tolshina, typ, number, quality):
	sheet= wb['sheet1']
	parametr1='≤' + str(float(tolshina)*0.1)
	parametr2='≤' + str(float(tolshina)*0.2)
	
	## Меняем шифр в шаблоне ##
	change_value('O46', shifr, sheet)
	
	## Меняем наименование объекта ##
	change_value('F48', name, sheet)
	
	## Меняем диаметр трубы ##
	change_value('O51', diametr, sheet)
	
	if float(diametr) < 300:
		sheet['A56'] = sheet['A56'].value.replace('x1', 'по всей длине шва.')
	else:
		sheet['A56'] = sheet['A56'].value.replace('x1', 'участка сварного шва после ремонта должна превышать длину отремонтированного участка на 100 мм в обе стороны.')
	if float(diametr) < 720:
		sheet['J83'] = sheet['J83'].value.format('.-')
	else:
		sheet['J83'] = sheet['J83'].value.format('8.-10')
	
	## Меняем толщину стенки ##
	change_value('O52', tolshina, sheet)
	
	## ТРЕБОВАНИЯ К ВЫПОЛНЕНИЮ ВИЗУАЛЬНОГО И ИЗМЕРИТЕЛЬНОГО КОНТРОЛЯ ##
	if float(tolshina) > 20:
		sheet['A56'] = sheet['A56'].value.replace('x2', tolshina)
	else:
		sheet['A56'] = sheet['A56'].value.replace('x2', '20')
	
	## Меняем глубину подреза ##
	if float(tolshina)*0.1 >= 0.5:
		sheet['J86'] = sheet['J86'].value.format('≤ 0,5')
	else:
		sheet['J86'] = sheet['J86'].value.format(parametr1) 
	
	## Глубина вогнутости корня (для сварных швов без подварки) ##
	if quality == 'A':
		if float(tolshina)*0.1 >= 1:
			sheet['J88'] = sheet['J88'].value.format('≤1')
		else:
			sheet['J88'] = sheet['J88'].value.format(a)
	if (quality == 'B') or (quality =='C'):
		if float(tolshina)*0.1 >= 2:
			sheet['J88'] = sheet['J88'].value.format('≤2')
		else:
			sheet['J88'] = sheet['J88'].value.format(parametr2)
	
	## Меняем номер удостоверения ##
	change_value('N115', number, sheet)
	
	## Меняем тип сварного соединения ##
	change_value('O54', typ, sheet)

	## Ширина облицовочного шва ##
	if typ == 'РД':
		if 6 < float(tolshina) <= 8:
			sheet['J82']= '11-18'
		elif 8.1 < float(tolshina) <= 12:
			sheet['J82']= '14-24'
		elif 12.1 < float(tolshina) <= 15:
			sheet['J82']= '18-28'
		elif 15.1 < float(tolshina) <= 20:
			sheet['J82']= '15-27'
		elif 20.1 < float(tolshina) <= 24:
			sheet['J82']= '18-31'
		elif 24.1 < float(tolshina) <= 27:
			sheet['J82']= '21-35'
		elif float(tolshina)>27:
			sheet['J82']= '21-35'
	if (typ == 'РД,МП+АПИ') or (typ == 'РД,МП+АПИ,МП+МПС+АФ'):
		if 6 < float(tolshina) <= 8:
			sheet['J82']= '11-18 для РД     11-17 для авт. Сварки'
		elif 8.1 < float(tolshina) <= 12:
			sheet['J82']= '14-24 для РД     16-24 для авт. Сварки'
		elif 12.1 < float(tolshina) <= 16:
			sheet['J82']= '18-28 для РД     19-27 для авт. Сварки'
		elif 16.1 < float(tolshina) <= 20.5:
			sheet['J82']= '15-27 для РД     20-28 для авт. Сварки'
		elif 20.6 < float(tolshina) <= 27:
			sheet['J82']= '18-31 для РД     22-30 для авт. Сварки'
		elif float(tolshina)>27:
			sheet['J82']= '21-35 для РД     22-30 для авт. Сварки'
	
	## Меняем уровень качества ##
	change_value('O50', quality, sheet)
	
	## Меняем дату ##
	change_value('L115', data, sheet)

	wb.save(dirname.replace('shablon', str(shifr)))
def uzk (shifr, name, diametr, tolshina, typ, number, quality):
		wb_name='app/static/' + shifr + '.xlsx'
		wb = openpyxl.load_workbook(wb_name)
		sheet = wb['sheet2']

		# Меняет шифр
		change_value ('N9', shifr, sheet)

		# Меняет наименование объекта
		change_value ('D22', name, sheet)

		# Меняет диаметр
		change_value ('A58', diametr, sheet)

		# Меняет толщину стенки
		change_value ('B58', tolshina, sheet)

		# Меняет номер удостоверения
		change_value ('L183', number, sheet)

		# Меняет тип сварки
		change_value ('C60', typ, sheet)

		# Меняет уровень качества
		change_value ('Q49', quality, sheet)

		# Меняет дату составления документа
		change_value ('N183', data, sheet)

		if quality == 'A':
			if 4 <= float (tolshina) < 6:
				tg_a = 2.74
				ztv = 10
				zona_max = round (2 * float (tolshina) * tg_a + ztv + 25, 2)
				zona_n = round (2 * float(tolshina) * tg_a + ztv, 2)
				sheet['I58'] = sheet['I58'].value.replace (sheet['I58'].value, '0.7')
				sheet['J58'] = sheet['J58'].value.replace (sheet['J58'].value, '1.4')
				sheet['K58'] = sheet['K58'].value.replace (sheet['K58'].value, '1.0')
				sheet['G58'] = sheet['G58'].value.replace (sheet['G58'].value, '5.0')
				sheet['H58'] = sheet['H58'].value.replace (sheet['H58'].value, '70.0°')
				sheet['B66'] = sheet['B66'].value.replace ('zona1', str (zona_max))
				sheet['D85'] = sheet['D85'].value.replace ('zona2', str (zona_n))
			elif 6 <= float (tolshina) < 8:
				tg_a = 2.74
				ztv = 10
				zona_max = round (2 * float(tolshina) * tg_a + ztv + 25, 2)
				zona_n = round (2 * float(tolshina) * tg_a + ztv, 2)
				sheet['I58'] = sheet['I58'].value.replace (sheet['I58'].value, '0.85')
				sheet['J58'] = sheet['J58'].value.replace (sheet['J58'].value, '1.4')
				sheet['K58'] = sheet['K58'].value.replace (sheet['K58'].value, '1.2')
				sheet['G58'] = sheet['G58'].value.replace (sheet['G58'].value, '5.0')
				sheet['H58'] = sheet['H58'].value.replace (sheet['H58'].value, '70.0°')
				sheet['B66'] = sheet['B66'].value.replace ('zona1', str (zona_max))
				sheet['D85'] = sheet['D85'].value.replace ('zona2', str (zona_n))
			elif 8 <= int(tolshina) < 12:
				tg_a = 2.144
				ztv = 10
				zona_max = round (2 * float(tolshina) * tg_a + ztv + 25, 2)
				zona_n = round (2 * float(tolshina) * tg_a + ztv, 2)
				sheet['I58'] = sheet['I58'].value.replace (sheet['I58'].value, '1.05')
				sheet['J58'] = sheet['J58'].value.replace (sheet['J58'].value, '2.0')
				sheet['K58'] = sheet['K58'].value.replace (sheet['K58'].value, '1.5')
				sheet['G58'] = sheet['G58'].value.replace (sheet['G58'].value, '5.0')
				sheet['H58'] = sheet['H58'].value.replace (sheet['H58'].value, '65.0°')
				sheet['B66'] = sheet['B66'].value.replace ('zona1', str (zona_max))
				sheet['D85'] = sheet['D85'].value.replace ('zona2', str (zona_n))
			elif 12 <= float(tolshina) < 15:
				tg_a = 2.144
				ztv = 10
				zona_max = round (2 * float(tolshina) * tg_a + ztv + 25, 2)
				zona_n = round (2 * float(tolshina) * tg_a + ztv, 2)
				sheet['I58'] = sheet['I58'].value.replace (sheet['I58'].value, '1.40')
				sheet['J58'] = sheet['J58'].value.replace (sheet['J58'].value, '2.0')
				sheet['K58'] = sheet['K58'].value.replace (sheet['K58'].value, '2.0')
				sheet['G58'] = sheet['G58'].value.replace (sheet['G58'].value, '2.5')
				sheet['H58'] = sheet['H58'].value.replace (sheet['H58'].value, '65.0°')
				sheet['B66'] = sheet['B66'].value.replace ('zona1', str (zona_max))
				sheet['D85'] = sheet['D85'].value.replace ('zona2', str (zona_n))
			elif 15 <= float(tolshina) < 20:
				tg_a = 2.144
				ztv = 10
				zona_max = round (2 * float(tolshina) * tg_a + ztv + 25, 2)
				zona_n = round (2 * float(tolshina) * tg_a + ztv, 2)
				sheet['I58'] = sheet['I58'].value.replace (sheet['I58'].value, '1.75')
				sheet['J58'] = sheet['J58'].value.replace (sheet['J58'].value, '2.5')
				sheet['K58'] = sheet['K58'].value.replace (sheet['K58'].value, '2.0')
				sheet['G58'] = sheet['G58'].value.replace (sheet['G58'].value, '2.5')
				sheet['H58'] = sheet['H58'].value.replace (sheet['H58'].value, '65.0°')
				sheet['B66'] = sheet['B66'].value.replace ('zona1', str (zona_max))
				sheet['D85'] = sheet['D85'].value.replace ('zona2', str (zona_n))
			elif 20 <= float(tolshina) < 26:
				tg_a = 2.144
				ztv = float(tolshina) * 0.5
				zona_max = round (2 * float(tolshina) * tg_a + ztv + 25, 2)
				zona_n = round (2 * float(tolshina) * tg_a + ztv, 2)
				sheet['I58'] = sheet['I58'].value.replace (sheet['I58'].value, '2.5')
				sheet['J58'] = sheet['J58'].value.replace (sheet['J58'].value, '3.5')
				sheet['K58'] = sheet['K58'].value.replace (sheet['K58'].value, '2.0')
				sheet['G58'] = sheet['G58'].value.replace (sheet['G58'].value, '2.5')
				sheet['H58'] = sheet['H58'].value.replace (sheet['H58'].value, '65.0°')
				sheet['B66'] = sheet['B66'].value.replace ('zona1', str (zona_max))
				sheet['D85'] = sheet['D85'].value.replace ('zona2', str (zona_n))
			elif 26 <= float(tolshina) < 40:
				sheet['I58'] = sheet['I58'].value.replace ('3.5')

		if quality == 'B' or 'C':
			if 4 <= (float(tolshina)) < 6:
				tg_a = 2.74
				ztv = 10
				zona_max = round (2 * float (tolshina) * tg_a + ztv + 25, 2)
				zona_n = round (2 * float(tolshina) * tg_a + ztv, 2)
				sheet['I58'] = sheet['I58'].value.replace (sheet['I58'].value, '1.0')
				sheet['J58'] = sheet['J58'].value.replace (sheet['J58'].value, '1.4')
				sheet['K58'] = sheet['K58'].value.replace (sheet['K58'].value, '1.0')
				sheet['G58'] = sheet['G58'].value.replace (sheet['G58'].value, '5.0')
				sheet['H58'] = sheet['H58'].value.replace (sheet['H58'].value, '70.0°')
				sheet['B66'] = sheet['B66'].value.replace ('zona1', str (zona_max))
				sheet['D85'] = sheet['D85'].value.replace ('zona2', str (zona_n))
			elif 6 <= float(tolshina) < 8:
				tg_a = 2.74
				ztv = 10
				zona_max = round (2 * float (tolshina) * tg_a + ztv + 25, 2)
				zona_n = round (2 * float(tolshina) * tg_a + ztv, 2)
				sheet['I58'] = sheet['I58'].value.replace (sheet['I58'].value, '1.2')
				sheet['J58'] = sheet['J58'].value.replace (sheet['J58'].value, '1.4')
				sheet['K58'] = sheet['K58'].value.replace (sheet['K58'].value, '1.2')
				sheet['G58'] = sheet['G58'].value.replace (sheet['G58'].value, '5.0')
				sheet['H58'] = sheet['H58'].value.replace (sheet['H58'].value, '70.0°')
				sheet['B66'] = sheet['B66'].value.replace ('zona1', str (zona_max))
				sheet['D85'] = sheet['D85'].value.replace ('zona2', str (zona_n))
			elif 8 <= float(tolshina) < 12:
				tg_a = 2.144
				ztv = 10
				zona_max = round (2 * float(tolshina) * tg_a + ztv + 25, 2)
				zona_n = round (2 * float(tolshina) * tg_a + ztv, 2)
				sheet['I58'] = sheet['I58'].value.format ('1.5')
				sheet['J58'] = sheet['J58'].value.format ('2.0')
				sheet['K58'] = sheet['K58'].value.format ('1.5')
				sheet['G58'] = sheet['G58'].value.format ('5.0')
				sheet['H58'] = sheet['H58'].value.replace (sheet['H58'].value, '65.0°')
				sheet['B66'] = sheet['B66'].value.replace ('zona1', str (zona_max))
				sheet['D85'] = sheet['D85'].value.replace ('zona2', str (zona_n))
			elif 12 <= float(tolshina) < 15:
				tg_a = 2.144
				ztv = 10
				zona_max = round (2 * float(tolshina) * tg_a + ztv + 25, 2)
				zona_n = round (2 * float(tolshina) * tg_a + ztv, 2)
				sheet['I58'] = sheet['I58'].value.replace (sheet['I58'].value, '2.0')
				sheet['J58'] = sheet['J58'].value.replace (sheet['J58'].value, '2.0')
				sheet['K58'] = sheet['K58'].value.replace (sheet['K58'].value, '2.0')
				sheet['G58'] = sheet['G58'].value.replace (sheet['G58'].value, '2.5')
				sheet['H58'] = sheet['H58'].value.replace (sheet['H58'].value, '65.0°')
				sheet['B66'] = sheet['B66'].value.replace ('zona1', str (zona_max))
				sheet['D85'] = sheet['D85'].value.replace ('zona2', str (zona_n))
			elif 15 <= float(tolshina) < 20:
				tg_a = 2.144
				ztv = 10
				zona_max = round (2 * float(tolshina) * tg_a + ztv + 25, 2)
				zona_n = round (2 * float(tolshina) * tg_a + ztv, 2)
				sheet['I58'] = sheet['I58'].value.replace (sheet['I58'].value, '2.5')
				sheet['J58'] = sheet['J58'].value.replace (sheet['J58'].value, '2.5')
				sheet['K58'] = sheet['K58'].value.replace (sheet['K58'].value, '2.0')
				sheet['G58'] = sheet['G58'].value.replace (sheet['G58'].value, '2.5')
				sheet['H58'] = sheet['H58'].value.replace (sheet['H58'].value, '65.0°')
				sheet['B66'] = sheet['B66'].value.replace ('zona1', str (zona_max))
				sheet['D85'] = sheet['D85'].value.replace ('zona2', str (zona_n))
			elif 20 <= int(tolshina) < 26:
				tg_a = 2.144
				ztv = float(tolshina) * 0.5
				zona_max = round (2 * float(tolshina) * tg_a + ztv + 25, 2)
				zona_n = round (2 * float(tolshina) * tg_a + ztv, 2)
				sheet['I58'] = sheet['I58'].value.replace (sheet['I58'].value, '3.5')
				sheet['J58'] = sheet['J58'].value.replace (sheet['J58'].value, '3.5')
				sheet['K58'] = sheet['K58'].value.replace (sheet['K58'].value, '2.0')
				sheet['G58'] = sheet['G58'].value.replace (sheet['G58'].value, '2.5')
				sheet['H58'] = sheet['H58'].value.replace (sheet['H58'].value, '65.0°')
				sheet['B66'] = sheet['B66'].value.replace ('zona1', str (zona_max))
				sheet['D85'] = sheet['D85'].value.replace ('zona2', str (zona_n))
			elif 26 <= float(tolshina) < 40:
				sheet['I58'] = sheet['I58'].value.replace (sheet['H58'].value, '5.0')

				
			sheet['O152'] = sheet['O152'].value.format ('50')
			if float (tolshina) * 2 >= 25:
				sheet['O153'] = sheet['O153'].value.format ('25')
			else:
				sheet['O153'] = sheet['O153'].value.format (str(float(2 * float (tolshina))))
			if float (tolshina) * 2 >= 15:
				sheet['O154'] = sheet['O154'].value.format ('15')
			else:
				sheet['O154'] = sheet['O154'].value.format (str(float(2 * float (tolshina))))
			if float (tolshina) * 2 >= 30:
				sheet['O155'] = sheet['O155'].value.format ('30')
			else:
				sheet['O155'] = sheet['O155'].value.format (str(float(2 * float (tolshina))))
			if float (tolshina) * 2 >= 15:
				sheet['M156'] = sheet['M156'].value.format ('15')
			else:
				sheet['M156'] = sheet['M156'].value.format (str(float(2 * float (tolshina))))
			if float (tolshina) * 2 >= 25:
				sheet['P156'] = sheet['P156'].value.format ('25')
			else:
				sheet['P156'] = sheet['P156'].value.format (str(float(2 * float (tolshina))))
			if float (tolshina) * 2 >= 50:
				sheet['O158'] = sheet['O158'].value.format ('50')
			else:
				sheet['O158'] = sheet['O158'].value.format (str(float(2 * float (tolshina))))
			if float (tolshina) * 2 >= 50:
				sheet['O159'] = sheet['O159'].value.format ('50')
			else:
				sheet['O159'] = sheet['O159'].value.format (str(float(2 * float (tolshina))))
			if float (tolshina) >= 15:
				sheet['O160'] = sheet['O160'].value.format ('15')
			else:
				sheet['O160'] = sheet['O160'].value.format (str(float(tolshina)))
			if float (tolshina) >= 15:
				sheet['O161'] = sheet['O161'].value.format ('15')
			else:
				sheet['O161'] = sheet['O161'].value.format (str(float(tolshina)))

		if 4 < float (tolshina) <= 6:
			sheet['I145'] = sheet['I145'].value.replace (sheet['I145'].value, '5.0')
		elif 6 < float (tolshina) <= 9:
			sheet['I145'] = sheet['I145'].value.format ('7.0')
		elif 9 < float (tolshina) <= 12:
			sheet['I145'] = sheet['I145'].value.replace (sheet['I145'].value, '10.0')
		elif (float (tolshina)) > 12:
			sheet['I145'] = sheet['I145'].value.replace (sheet['I145'].value, '12.0')

		glubina = round (2 * float (tolshina) / 3, 2)
		sheet['A162'] = sheet['A162'].value.replace ('perimetr', str (glubina))
		wb.save(wb_name)		
def remake_shablon(shifr, name, diametr, tolshina, typ, number, quality):
	vik(shifr, name, diametr, tolshina, typ, number, quality)
	uzk (shifr, name, diametr, tolshina, typ, number, quality)


	