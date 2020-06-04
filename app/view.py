from .import app
from flask import render_template, request, url_for, redirect, flash, g, session
from .forms import Wform
from .remake_xls import remake_shablon
from .emailxls import send_techcard
import os


@app.before_request
def before_request():
	if 'shifr' in session:
		form= Wform()
		form.shifr.data = session['shifr']
		form.name.data = session['name']
		form.diametr.data = session['diametr']
		form.tolshina.data = session['tolshina']
		form.number.data = session['number']
		form.typ.data = session['typ']
		form.quality.data = session['quality']
		g.form = form
		session.pop('shifr', None)
		session.pop('name', None)
		session.pop('diametr', None)
		session.pop('tolshina', None)
		session.pop('number', None)
		session.pop('typ', None)
		session.pop('quality', None)
	else:
		g.form = Wform()
@app.route('/', methods =['GET', 'POST'])
def index():
	form=g.form
	if form.validate_on_submit():
		shifr = form.shifr.data
		name = form.name.data
		diametr= form.diametr.data
		tolshina= form.tolshina.data
		number= form.number.data
		typ = form.typ.data
		quality = form.quality.data
		remake_shablon(shifr, name, diametr, tolshina, typ, number, quality)
		flash(f'Техкарта c шифром "{shifr}" отправлена')
		filename = 'static/' + shifr + '.xlsx'
		##  отправить файл по почте ##
		#send_techcard(app.config['MAIL_RECIPIENT'], name, filename)

		## Удалить после отправки ##
		#path = os.path.join(os.path.abspath(os.path.dirname(__file__)), filename)
		#os.remove(path)
		session['shifr'] = shifr
		session['name'] = name
		session['diametr']= diametr
		session['tolshina'] = tolshina
		session['number'] = number
		session['typ'] = typ
		session['quality'] = quality
		return redirect(url_for('index'))	
	return render_template('index.html', form=form)
	