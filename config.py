# config.py
import os

class Config:
    # Gantilah username dan password dengan yang sesuai dari Laragon
    SQLALCHEMY_DATABASE_URI = 'mysql://root:@localhost/vida'  # username=root, password=kosong
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    SQLALCHEMY_ENGINE_OPTIONS = {
        'pool_pre_ping': True,
        'pool_recycle' : 280
    }
    SECRET_KEY = os.urandom(24)