@echo off
rem "Renombra los archivos para que tengan siempre el mismo nombre."

rename D:\Dropbox\Aplicaciones\Producteca\price*.* price.*
rename D:\Dropbox\Aplicaciones\Producteca\stock*.* stock.*
rename D:\Dropbox\TUCUMAN\stock*.* stock_tucuman.*
del /S /F /Q D:\Dropbox\TUCUMAN\price*.*
copy D:\Dropbox\TUCUMAN\stock_tucuman.* D:\Dropbox\Aplicaciones\Producteca\stock_tucuman.*
