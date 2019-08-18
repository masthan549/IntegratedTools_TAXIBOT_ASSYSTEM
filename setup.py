from cx_Freeze import setup, Executable

EXE = 'IntegratedTaxiTools_GUI_Main'
filename = EXE+'.py'

setup(
    name = EXE ,
    version = "0.1" ,
    description = "" ,
    executables = [Executable(filename)] ,
    )							 