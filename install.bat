set venv_root_dir=%~F1

set PYTHON_MAJOR_VERSION=0
for /f %%i in ('python -c "import sys; print(sys.version_info[0])"') do set PYTHON_MAJOR_VERSION=%%i

If %PYTHON_MAJOR_VERSION%==3 GOTO PYTHON_3

ECHO "Need Python 3+"
EXIT 1
:PYTHON_3

python -m venv %venv_root_dir%

%venv_root_dir%\Scripts\python.exe -m pip install --upgrade pip

%venv_root_dir%\Scripts\pip.exe install -r %py_ocal_to_gcal_root_dir%\requirements.txt
