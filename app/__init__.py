from flask import Flask
from flask_mail import Mail
from .config import Configuration

app= Flask(__name__)
mail = Mail(app)
app.config.from_object(Configuration)
from .import view