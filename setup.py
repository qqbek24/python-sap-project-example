import os
import sys

os.system('py -m pip install --user virtualenv')
os.system('py -m venv env')
home = os.path.dirname(os.path.abspath(__file__))
env_python = f"{home}\\env\\Scripts\\python.exe"
os.system(f'{env_python} -m pip install --upgrade pip')
os.system(f'{env_python} -m pip install -r requirements.txt')
os.system(f'{env_python} -m pip install pytest')
os.system(f'{env_python} -m pip install freezegun')
