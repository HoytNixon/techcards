import os


class Configuration(object):
	SECRET_KEY =os.environ.get('SECRET_KEY') or 'verysecretkey'
	MAIL_SERVER = 'smtp.googlemail.com'
	MAIL_PORT = 465
	MAIL_USE_SSL = True
	MAIL_USERNAME =os.environ.get('MAIL_USERNAME') or 'techkards@gmail.com'
	MAIL_PASSWORD =os.environ.get('MAIL_PASSWORD') or 'Zuw78pi4'
	MAIL_DEFAULT_SENDER =os.environ.get('MAIL_USERNAME') or 'techkards@gmail.com'
	MAIL_RECIPIENT =os.environ.get('MAIL_RECIPIENT') or 'ueks@svartr.ru'