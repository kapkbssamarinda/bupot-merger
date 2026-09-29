import os
import sys

# Tambahkan root path ke sys.path agar Vercel runtime dapat menemukan app.py
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from app import app

# Vercel serverless WSGI entry point
# Objek 'app' akan otomatis dikenali dan dijalankan oleh runtime Python Vercel
