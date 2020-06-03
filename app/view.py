from .import app
from flask import render_template, request, url_for, redirect, flash
from .forms import Wform
from .remake_xls import remake_shablon
from .emailxls import send_techcard
import os
@app.route('/', methods =['GET', 'POST'])
def index():
	form = Wform()
	if form.validate_on_submit():
		shifr = form.shifr.data
		name = form.name.data
		diametr= form.diametr.data
		tolshina= form.tolshina.data
		number= form.number.data
		typ = form.typ.data
		quality = form.quality.data
		remake_shablon(shifr, name, diametr, tolshina, typ, number, quality)

		filename = 'static/' + shifr + '.xlsx'
		##  отправить файл по почте ##
		#send_techcard(app.config['MAIL_RECIPIENT'], name, filename)

		## Удалить после отправки ##
		#path = os.path.join(os.path.abspath(os.path.dirname(__file__)), filename)
		#os.remove(path)

		form.shifr.data = shifr
		form.name.data = name
		form.diametr.data = diametr
		form.tolshina.data = tolshina
		form.number.data = number
		form.typ.data = typ
		form.quality.data = quality
		return render_template('index.html', form=form)	
	return render_template('index.html', form=form)