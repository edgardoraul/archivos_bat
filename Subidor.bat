@echo off
echo user web@rerda.com>> ftpcmd.dat
echo 919q3Nnhremfrursmv>> ftpcmd.dat
echo bin>> ftpcmd.dat
echo cd upload/>> ftpcmd.dat
echo delete price.csv>> ftpcmd.dat
echo delete stock.csv>> ftpcmd.dat
echo put D:\Dropbox\Aplicaciones\Producteca\price.csv>> ftpcmd.dat
echo put D:\Dropbox\Aplicaciones\Producteca\stock.csv>> ftpcmd.dat
echo quit>> ftpcmd.dat
ftp -n -s:ftpcmd.dat ftp.rerda.com
del ftpcmd.dat