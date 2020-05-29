from flask_wtf import FlaskForm
from wtforms import StringField, SelectField,SubmitField

class Wform(FlaskForm):

	shifr= StringField(u'Шифр')
	name= StringField(u'Название объекта')
	diametr= StringField(u'Диаметр трубы')
	tolshina= StringField(u'Толщина стенки')
	number= StringField(u'Номер удостоверения')
	typ= SelectField(u'Тип сварного соединения', choices=[('РД', 'РД'), ('РД,МП+АПИ', 'РД,МП+АПИ'), ('РД,МП+АПИ,МП+МПС+АФ', 'РД,МП+АПИ,МП+МПС+АФ')])
	quality = SelectField(u'Уровень качества', choices=[('A', 'A'),('B', 'B'), ('C', 'C')])