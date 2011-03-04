@echo off
rem ==$Workfile: mtbin.cmd $============================================
rem  mexbin creates a binary distribution for MapiExport.
rem  Version    : $Id$
rem  Application: MAPI
rem  Platform   : Win32, MAPI
rem --------------------------------------------------------------------
rem Copyright  : 2009, Enter AG, Z�rich, Switzerland
rem Created    : 08.10.2009, Hartwig Thomas
rem ====================================================================
if "%~1"=="-?" goto help
if "%~1"=="/?" goto help
if "%~1"=="-h" goto help
if "%~1"=="/h" goto help
if "%~1"=="-H" goto help
if "%~1"=="/h" goto help
chcp 1252

set version=2_08

rem --------------------------------------------------------------------
rem execution directory from which cmd is called
rem --------------------------------------------------------------------
set execdir=%~dp0
rem detach the trailing backslash
set execdir=%execdir:~0,-1%

rem --------------------------------------------------------------------
rem list of distribution files for pkzip command
rem --------------------------------------------------------------------
set distlist=%execdir%\mtbin.lst

rem --------------------------------------------------------------------
rem target file
rem --------------------------------------------------------------------
set zipfile=%execdir%\mtbin_%version%.zip
del "%zipfile%"

rem --------------------------------------------------------------------
rem change to directory above msgtext
rem --------------------------------------------------------------------
cd "%execdir%\..\.."

rem --------------------------------------------------------------------
rem prepare distribution
rem --------------------------------------------------------------------
copy /b/y MsgText\Release\MsgText.exe MsgText\bin\MsgText.exe
@echo pkzipc -add -attr=all -dir=current "%zipfile%" "@%distlist%"
pkzipc -add -attr=all -dir=current "%zipfile%" "@%distlist%"
cd "%execdir%"
@echo Distribution zipped
goto exit

:help
rem --------------------------------------------------------------------
rem help for usage
rem --------------------------------------------------------------------
rem we need the quotes for protecting the angular brackets
@echo "Usage                                                               "
@echo "  mtbin.cmd                                                         "
@echo "creates the distribution mtbin.zip in the same folder as this cmd   "
@echo "using the file list mtbin.lst, also on that folder.                 "
@echo "                                                                    "

rem --------------------------------------------------------------------
rem exit
rem --------------------------------------------------------------------
:exit
