from flask import Flask
from flask_mail import Mail
from .config import Configuration

app= Flask(__name__)
app.config.from_object(Configuration)
mail = Mail(app)
from .import view