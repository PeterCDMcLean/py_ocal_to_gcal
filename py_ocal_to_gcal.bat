set venv_root_dir=%~F1
set py_ocal_to_gcal_root_dir=%~F2

%venv_root_dir%\Scripts\python.exe %py_ocal_to_gcal_root_dir%\py_ocal_to_gcal.py --work %3

exit /B 0