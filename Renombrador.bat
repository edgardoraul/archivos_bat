@echo off
rem "Renombra los archivos para que tengan siempre el mismo nombre."

rename D:\Dropbox\Aplicaciones\Producteca\price*.* price.*
rename D:\Dropbox\Aplicaciones\Producteca\stock*.* stock.*
rename D:\Dropbox\TUCUMAN\stock*.* stock_tucuman.*
del /S /F /Q D:\Dropbox\TUCUMAN\price*.*
copy D:\Dropbox\TUCUMAN\stock_tucuman.* D:\Dropbox\Aplicaciones\Producteca\stock_tucuman.*

rename D:\Dropbox\MOVIL5\stock*.* stock_movil5.*
del /S /F /Q D:\Dropbox\MOVIL5\price*.*
copy D:\Dropbox\MOVIL5\stock_movil5.* D:\Dropbox\Aplicaciones\Producteca\stock_movil5.*

rename D:\Dropbox\MOVIL2\stock*.* stock_movil2.*
del /S /F /Q D:\Dropbox\MOVIL2\price*.*
copy D:\Dropbox\MOVIL2\stock_movil2.* D:\Dropbox\Aplicaciones\Producteca\stock_movil2.*