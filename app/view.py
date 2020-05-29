from .import app
from flask import render_template, request, url_for, redirect
from .forms import Wform
from .remake_xls import remake_shablon
from emailxls import send_techcard

@app.route('/', methods =['GET', 'POST'])
def index():
	form = Wform()
	if request.method == 'POST':
		shifr = form.shifr.data
		name = form.name.data
		diametr= form.diametr.data
		tolshina= form.tolshina.data
		number= form.number.data
		typ = form.typ.data
		quality = form.quality.data
		remake_shablon(shifr, name, diametr, tolshina, typ, number, quality)
		return redirect(url_for('index'))
	return render_template('index.html', form=form)