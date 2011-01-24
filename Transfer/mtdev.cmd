@echo off
rem ==$Workfile: mtdev.cmd $============================================
rem  mtxdev creates a binary distribution for MsgText.
rem  Version    : $Id$
rem  Application: MAPI
rem  Platform   : Win32, MAPI
rem --------------------------------------------------------------------
rem Copyright  : 2009, Enter AG, Zürich, Switzerland
rem Created    : 08.10.2009, Hartwig Thomas
rem ====================================================================
if "%~1"=="-?" goto help
if "%~1"=="/?" goto help
if "%~1"=="-h" goto help
if "%~1"=="/h" goto help
if "%~1"=="-H" goto help
if "%~1"=="/h" goto help
chcp 1252

set version=2_01

rem --------------------------------------------------------------------
rem execution directory from which cmd is called
rem --------------------------------------------------------------------
set execdir=%~dp0
rem detach the trailing backslash
set execdir=%execdir:~0,-1%

rem --------------------------------------------------------------------
rem list of distribution files for pkzip command
rem --------------------------------------------------------------------
set distlist=%execdir%\mtdev.lst

rem --------------------------------------------------------------------
rem target file
rem --------------------------------------------------------------------
set zipfile=%execdir%\mtdev_%version%.zip
del "%zipfile%"

rem --------------------------------------------------------------------
rem change to directory above MsgText
rem --------------------------------------------------------------------
cd "%execdir%\..\.."

rem --------------------------------------------------------------------
rem prepare distribution
rem --------------------------------------------------------------------
copy /b/y MsgText\Release\MsgText.exe MsgText\bin\MsgText.exe
pkzipc -add -attr=all -dir=current "%zipfile%" "@%distlist%"
cd "%execdir%"
@echo Source zipped
goto exit

:help
rem --------------------------------------------------------------------
rem help for usage
rem --------------------------------------------------------------------
rem we need the quotes for protecting the angular brackets
@echo "Usage                                                               "
@echo "  mtdev.cmd                                                         "
@echo "creates the distribution mtbin.zip in the same folder as this cmd   "
@echo "using the file list mtdev.lst, also on that folder.                 "
@echo "                                                                    "

rem --------------------------------------------------------------------
rem exit
rem --------------------------------------------------------------------
:exit
