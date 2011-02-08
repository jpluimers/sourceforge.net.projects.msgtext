/*-- $Workfile: Fs.h $ -----------------------------------------------
Object for file system handling.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Filesystem handling.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
This file is part of MsgText.
MsgText is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
any later version.
MsgText is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
You should have received a copy of the GNU General Public License
along with MsgText.  If not, see <http://www.gnu.org/licenses/>.
----------------------------------------------------------------------*/
#ifndef FS_H
#define FS_H

#include "FileTime.h"

/*======================================================================
Utilites
======================================================================*/
/*----------------------------------------------------------------------
trim
  returns a freshly allocated copy of the trimmed portion of the input
  string.
heap is not invariant - caller must delete trimmed string.
----------------------------------------------------------------------*/
char* trim(char* ps);

/*----------------------------------------------------------------------
newMultiByte
  p string to be converted from UNICODE to ISO-Latin.
  returns: converted string or NULL, if an error occurred.
heap is not invariant - caller must delete converted string.
----------------------------------------------------------------------*/
char* newMultiByte(wchar_t* pWc);

/*----------------------------------------------------------------------
newWideChar
  p string to be converted from ISO-Latin to UNICODE.
  returns: converted string or NULL, if an error occurred.
heap is not invariant - caller must delete converted string.
----------------------------------------------------------------------*/
wchar_t* newWideChar(char* p);

/*======================================================================
Fs wraps a name file or folder
======================================================================*/
class Fs
{
public:
  /*-------------------------------------------------------------------
  constructor
    pName: if terminated by a slash, the Fs object represents a directory,
           otherwise a file name.
  heap size invariant.
  -------------------------------------------------------------------*/
  Fs(char* pName);
  /*-------------------------------------------------------------------
  destructor
  heap size invariant.
  -------------------------------------------------------------------*/
  ~Fs();

  /*-------------------------------------------------------------------
  name returns the file name.
    returns: stored file name.
  heap size not invariant: caller must free resulting string.
  -------------------------------------------------------------------*/
  char* name();

  /*-------------------------------------------------------------------
  extension returns the extension of the file name.
    returns: extension of file name, "" if it had no extension, or NULL if an error occurred.
  heap size not invariant: caller must delete result.
  -------------------------------------------------------------------*/
  char* extension();

  /*-------------------------------------------------------------------
  toWideChar converts file/directory name to UNICODE
    returns: converted name or NULL, if an error occurred.
  heap size not invariant: caller must delete result.
  -------------------------------------------------------------------*/
  wchar_t* toWideChar();

  /*-------------------------------------------------------------------
  normalize converts file/directory name to "safe" ASCII characters.
    returns: converted name or NULL, if an error occurred.
  heap size not invariant: caller must delete result.
  -------------------------------------------------------------------*/
  char* normalize();

  /*-------------------------------------------------------------------
  isFolder checks, if name terminates on a backslash.
    returns: true, if name terminates on a backslash.
  heap size invariant.
  -------------------------------------------------------------------*/
  BOOL isFolder();

  /*-------------------------------------------------------------------
  exists checks, if file or folder exist
    returns: true, if file or folder exist.
  heap size invariant.
  -------------------------------------------------------------------*/
  BOOL exists();

  /*-------------------------------------------------------------------
  newParentFolder returns the name of the parent folder.
    returns: full parent folder name or NULL if an error occurred.
  heap size not invariant: caller must delete result.
  -------------------------------------------------------------------*/
  char* newParentFolder();

  /*-------------------------------------------------------------------
  newFilename returns the filename portion of the full filename.
    returns: filename portion of the full filename.
  heap size not invariant: caller must delete result.
  -------------------------------------------------------------------*/
  char* newFilename();

  /*-------------------------------------------------------------------
  newSubFolder concatenates directory name with subfolder name.
    pSubFolder: name of subfolder
    returns: full subfolder name or NULL if an error occurred.
  heap size not invariant: caller must delete result.
  -------------------------------------------------------------------*/
  char* newSubFolder(char* pSubFolder);

  /*-------------------------------------------------------------------
  newSubFolder concatenates directory name with subfolder name truncating
  the latter to a given maximum length and disambiguating it.
    pSubFolder: name of subfolder
    ulMaxFilenameLength: maximum length for subfolder name
    returns: full subfolder name or NULL if an error occurred.
  heap size not invariant: caller must delete result.
  -------------------------------------------------------------------*/
  char* newSubFolder(char* pSubFolder, ULONG ulMaxFilenameLength);

  /*-------------------------------------------------------------------
  normalizedSubFolder concatenates directory name with normalized subfolder 
    name truncating the latter to a given maximum length and 
    disambiguating it.
    pSubFolder: name of subfolder
    ulMaxFilenameLength: maximum length for subfolder name
    returns: full subfolder name or NULL if an error occurred.
  heap size not invariant: caller must delete result.
  -------------------------------------------------------------------*/
  char* normalizedSubFolder(char* pSubFolder, ULONG ulMaxFilenameLength);

  /*-------------------------------------------------------------------
  newFile concatenates directory name with file name.
    pFile: name of file
    returns: full file name or NULL if an error occurred.
  heap size not invariant: caller must delete result.
  -------------------------------------------------------------------*/
  char* newFile(char* pFile);

  /*-------------------------------------------------------------------
  newFile concatenates directory name with file name and extension
  truncating the file name to a given maximum length and disambiguating it.
    pFile: name of file
    ulMaxFilenameLength: maximum length for file name
    pExtension: extension, may be NULL if no extension is desired.
    returns: full file name or NULL if an error occurred.
  heap size not invariant: caller must delete result.
  -------------------------------------------------------------------*/
  char* newFile(char* pFile, ULONG ulMaxFilenameLength, char* pExtension);

  /*-------------------------------------------------------------------
  normalizeFile concatenates directory name with normalized file name 
  and extension truncating the file name to a given maximum length and 
  disambiguating it.
    pFile: name of file
    ulMaxFilenameLength: maximum length for file name
    pExtension: extension, may be NULL if no extension is desired.
    returns: full file name or NULL if an error occurred.
  heap size not invariant: caller must delete result.
  -------------------------------------------------------------------*/
  char* normalizeFile(char* pFile, ULONG ulMaxFilenameLength, char* pExtension);

  /*-------------------------------------------------------------------
  derivedFile detaches the extension of this file, attaches the
    given extension (including '.') and returns the result.
    pExtension: extension to be attached (or NULL)
    returns: derived file name or NULL if an error occurred.
  heap size not invariant: caller must delete result.
  -------------------------------------------------------------------*/
  char* deriveFile(char* pExtension);

  /*-------------------------------------------------------------------
  setTimes sets the file times.
    returns: 0, if no error occurred.
  heap size invariant.
  -------------------------------------------------------------------*/
  DWORD setTimes(FileTime* pftCreation, 
                 FileTime* pftModification, 
                 FileTime* pftAccess);

  /*-------------------------------------------------------------------
  deleteFile deletes the file.
    returns: 0, if no error occurred.
  heap size invariant.
  -------------------------------------------------------------------*/
  DWORD deleteFile();

  /*-------------------------------------------------------------------
  createFolder creates the directory.
    returns: 0, if no error occurred.
  heap size invariant.
  -------------------------------------------------------------------*/
  DWORD createFolder();

  /*-------------------------------------------------------------------
  deleteFolder creates the directory.
    returns: 0, if no error occurred.
  heap size invariant.
  -------------------------------------------------------------------*/
  DWORD deleteFolder(BOOL bDeleteContents);

private:
  /*-------------------------------------------------------------------
  exists checks, whether subfolder or file exist within this
    sub folder.
    returns: true, if subfolder or file does exist.
  heap size invariant.
  -------------------------------------------------------------------*/
  BOOL exists(char* pSubFolderOrFile);

  /*-------------------------------------------------------------------
  disambiguate creates a file name truncated to maximum length and
    disambiguated with a counter suffix to avoid name collisions
    with existing files/directories.
    returns: disambiguated file name or NULL if an error occurred.
  heap size not invariant: caller must delete result.
  -------------------------------------------------------------------*/
  char* disambiguate(char* pFile, ULONG ulMaxFilenameLength, char* pExtension); /* must be deleted by caller */

  char* m_pName;
}; /* class Fs */

#endif // FS_H not defined
/*==End of File=======================================================*/
