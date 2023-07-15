## Create VirtualEnv
    virtualenv env

## Activate VirtualEnv
    venv/Scripts/activate

## Libraries
pip install openpyxl
pip install pandas
pip install auto-py-to-exe
pip install PyQt5

### To generate the UI file 
pyuic5 UI.ui -o UI.py

## Execute following to generate the exe file
auto-py-to-exe