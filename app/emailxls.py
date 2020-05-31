from flask_mail import Message
from app import mail
from threading import Thread

def send_async_email(app, msg):
	with app.app_context():
		mail.send(msg)
def send_techcard(to, subject, filename, **kwargs):
	msg = Message(subject, recipients = [to])
	with app.open_resource(filename) as file:
		msg.attach(filename,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', file.read())
	async_mail = Thread(target=send_async_email, args = [app, msg])
	async_mail.start()
	return async_mail