# config.py
import os

class Config:
    # Gantilah username dan password dengan yang sesuai dari Laragon
    SQLALCHEMY_DATABASE_URI = 'mysql://root:@localhost/vida'  # username=root, password=kosong
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SECRET_KEY = os.urandom(24)