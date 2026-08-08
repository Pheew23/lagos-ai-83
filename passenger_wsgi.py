import os
import sys

# Mengarahkan server ke direktori aplikasi saat ini
sys.path.insert(0, os.path.dirname(__file__))

# Mengimpor aplikasi Flask dari app.py
from app import app as application
