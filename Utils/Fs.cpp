/*-- $Workfile: Fs.cpp $ -----------------------------------------------
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
#define _CRT_SECURE_NO_WARNINGS
#define STRICT
#include <windows.h>
#include <stdio.h>
#include "Fs.h"
#include "Heap.h"

/* translation table for normalization */
static char* apXlat[256] =
{
  /* 00 */ "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_",
  /* 10 */ "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", 
  /* 20 */ "_", "!", "_", "#", "$", "%", "_", "_", "(", ")", "_", "+", ",", "-", ".", "_",
  /* 30 */ "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "_", "_", "_", "=", "_", "_",
  /* 40 */ "@", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O",
  /* 50 */ "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "[", "_", "]", "_", "_",
  /* 60 */ "_", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o",
  /* 70 */ "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "{", "_", "}", "~", "_",
  /* 80 */ "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_",
  /* 90 */ "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", "_", 
  /* A0 */ "_", "_", "c", "L=", "E=", "Y=", "S", "SS", "s", "(c)", "a", "_", "_", "_", "(r)", "_",
  /* B0 */ "deg", "+-", "2", "3", "Z", "u", "P", ".", "z", "1", "o", "_", "OE", "oe", "Y", "_",
  /* C0 */ "A", "A", "A", "A", "Ae", "A", "AE", "C", "E", "E", "E", "E", "I", "I", "I", "I",
  /* D0 */ "Dh", "N", "O", "O", "O", "O", "Oe", "x", "O", "U", "U", "U", "Ue", "Y", "th", "ss",
  /* E0 */ "a", "a", "a", "a", "ae", "a", "ae", "c", "e", "e", "e", "e", "i", "i", "i", "i",
  /* F0 */ "dh", "n", "o", "o", "o", "o", "oe", "_", "o", "u", "u", "u", "ue", "y", "Th", "y"
};

/*----------------------------------------------------------------------
trim
  returns a freshly allocated copy of the trimmed portion of the input
  string.
heap is not invariant - caller must delete trimmed string.
----------------------------------------------------------------------*/
char* trim(char* ps)
{
  char* pt = NULL;
  char* p;
  if (ps != NULL)
  {
    for (p = ps; isspace((unsigned char)*p); p++) {}
    pt = _strdup(p);
    if ((pt != NULL))
    {
      for (p = pt + strlen(pt); (p != pt) && (*p == '\0');)
      {
        p--;
        if (isspace((unsigned char)*p))
          *p = '\0';
      }
    }
  }
  return pt;
} /* trim */

/*----------------------------------------------------------------------
newWideChar
  p string to be converted from ISO-Latin to UNICODE.
  returns: converted string or NULL, if an error occurred.
heap is not invariant - caller must delete converted string.
----------------------------------------------------------------------*/
wchar_t* newWideChar(char* p)
{
  int iLength = strlen(p);
  wchar_t* pWc = new wchar_t[iLength+1];
  memset(pWc,0,(iLength+1)*sizeof(wchar_t));
  DWORD dwError = MultiByteToWideChar(CP_ACP, 0, p, -1, pWc, iLength+1);
  if (dwError == 0)
  {
    dwError = GetLastError();
    if (dwError != 0)
    {
      delete pWc;
      pWc = NULL;
    }
  }
  else
    dwError = 0;
  return pWc;
} /* newWideChar */

/*----------------------------------------------------------------------
newMultiByte
  p string to be converted from UNICODE to ISO-Latin.
  returns: converted string or NULL, if an error occurred.
heap is not invariant - caller must delete converted string.
----------------------------------------------------------------------*/
char* newMultiByte(wchar_t* pWc)
{
  BOOL bDefaultUsed = false;
  int iLength = wcslen(pWc);
  char* p = new char[iLength+1];
  memset(p,0,iLength+1);
  DWORD dwError = WideCharToMultiByte(CP_ACP, 0, pWc, -1, p, iLength+1, NULL, NULL);
  if (dwError == 0)
  {
    dwError = GetLastError();
    if (dwError != 0)
    {
      delete p;
      p = NULL;
    }
  }
  else
    dwError = 0;
  return p;
} /* newMultiByte */


/*--------------------------------------------------------------------*/
Fs::Fs(char* pName)
{
  m_pName = pName;
} /* Fs */

/*--------------------------------------------------------------------*/
Fs::~Fs()
{
} /* ~Fs */

/*--------------------------------------------------------------------*/
char* Fs::name()
{
  return _strdup(this->m_pName);
} /* name */

/*--------------------------------------------------------------------*/
char* Fs::extension()
{
  char* pExtension = NULL;
  unsigned u;
  if (!this->isFolder())
  {
    char* pName = _strdup(m_pName);
    /* find last '.', '/' or '\\' */
    for (u = strlen(pName) - 1;
      (u >= 0) && (pName[u] != '.') && (pName[u] != '\\') && (pName[u] != '/');
      u--) {}
    /* extension */
    if (pName[u] == '.')
      pExtension = _strdup(&pName[u]);
    else
      pExtension = _strdup("");
  }
  return pExtension;
} /* extension */

/*--------------------------------------------------------------------*/
BOOL Fs::isFolder()
{
  return (m_pName[strlen(m_pName)-1] == '\\');
} /* isFolder */

/*--------------------------------------------------------------------*/
BOOL Fs::exists()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  BOOL bExists = FALSE;
  DWORD dwError = 0;
  /* detach trailing backslash */
  char* pName = _strdup(m_pName);
  if (this->isFolder())
    pName[strlen(pName)-1] = '\0';
  wchar_t* pwName = newWideChar(pName);
  DWORD dwAttr = GetFileAttributes(pwName);
  if (dwAttr != 0xFFFFFFFF)
    bExists = TRUE;
  else
    dwError = GetLastError();
  delete pwName;
  free(pName);
#ifdef _DEBUG
  if (!(pHeap->isValid() && pHeap->hasInitialSize()))
    bExists = FALSE;
  delete pHeap;
#endif
  return bExists;
} /* exists */

/*--------------------------------------------------------------------*/
BOOL Fs::exists(char* pSubFolderOrFile)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  BOOL bExists = false;
  if (this->isFolder())
  {
    char* pNew = (char*)malloc(strlen(m_pName)+1+strlen(pSubFolderOrFile)+1);
    if (pNew != NULL)
    {
      strcpy(pNew,m_pName);
      if (pNew[strlen(pNew)-1] != '\\')
        strcat(pNew,"\\");
      strcat(pNew,pSubFolderOrFile);
      Fs* pFsNew = new Fs(pNew);
      bExists = pFsNew->exists();
      delete pFsNew;
      delete pNew;
    }
  }
#ifdef _DEBUG
  if (!(pHeap->isValid() && pHeap->hasInitialSize()))
    bExists = FALSE;
  delete pHeap;
#endif
  return bExists;
} /* exists */

/*--------------------------------------------------------------------*/
char* Fs::disambiguate(char* pFile, ULONG ulMaxFilenameLength, char* pExtension)
{
  char* pExt;
  if (pExtension != NULL)
  {
    pExt = new char[strlen(pExtension)+2];
    if (pExtension[0] != '.')
    {
      pExt[0] = '.';
      strcpy(&pExt[1],pExtension);
    }
    else
      strcpy(pExt,pExtension);
  }
  else
    pExt = _strdup("");
  char* pSuffix = new char[12+strlen(pExt)];
  if (strlen(pFile) > ulMaxFilenameLength)
    pFile[ulMaxFilenameLength] = '\0';
  for (int iLength = strlen(pFile); (iLength > 0) && (pFile[iLength-1] == '.'); iLength--)
    pFile[iLength-1] = '\0';
  ULONG ulLength = strlen(pFile);
  char* pNew = new char[ulMaxFilenameLength+12+strlen(pExt)];
  strcpy(pNew,pFile);
  strcpy(&pNew[strlen(pFile)],pExt);
  for (ULONG ul = 0; this->exists(pNew); ul++)
  {
    sprintf(pSuffix,"%u",ul);
    if (strlen(pFile)+strlen(pSuffix) < ulMaxFilenameLength)
      strcpy(&pNew[strlen(pFile)],pSuffix); // append suffix
    else
      strcpy(&pNew[ulMaxFilenameLength-strlen(pSuffix)],pSuffix); // squeeze suffix in
    // append extension
    strcat(pNew,pExt);
  }
  delete pSuffix;
  delete pExt;
  return pNew;
} /* disambiguate */

/*--------------------------------------------------------------------*/
wchar_t* Fs::toWideChar()
{
  return newWideChar(this->m_pName);
} /* toWideChar */

/*--------------------------------------------------------------------*/
char* Fs::normalize()
{
  char* pNormalized = new char[3*strlen(m_pName)];
  if (pNormalized != NULL)
  {
    char* p = pNormalized;
    for (size_t i = 0; i < strlen(m_pName); i++)
    {
      /* C++ has these stupid signed characters ... */
      int iCode = (unsigned char)m_pName[i];
      strcpy(p,apXlat[iCode]);
      p += strlen(apXlat[iCode]);
    }
  }
  return pNormalized;
} /* normalize */

/*--------------------------------------------------------------------*/
char* Fs::newParentFolder()
{
  char* pNewParentFolder = NULL;
  int i;
  for (i = strlen(m_pName) - 2; /* skip last character, which might by a backslash */
       (i >= 0) && (m_pName[i] != '\\');
       i--) {}
  if (i >= 0)
  {
    pNewParentFolder = (char*)malloc(i+2);
    memcpy(pNewParentFolder,m_pName,i+1);
    pNewParentFolder[i+1] = '\0';
  }
  return pNewParentFolder;
} /* newParentFolder */

/*--------------------------------------------------------------------*/
char* Fs::newFilename()
{
  char* pNewFilename = NULL;
  int i;
  for (i = strlen(m_pName) - 1;
       (i >= 0) && (m_pName[i] != '\\');
       i--) {}
  i++;
  pNewFilename = _strdup(&m_pName[i]);
  return pNewFilename;
} /* newFilename */

/*--------------------------------------------------------------------*/
char* Fs::newSubFolder(char* pSubFolder)
{
  char* pNewSubFolder = NULL;
  if (this->isFolder())
  {
    pNewSubFolder = new char[strlen(m_pName)+strlen(pSubFolder)+2];
    if (pNewSubFolder)
    {
      strcpy(pNewSubFolder,m_pName);
      strcat(pNewSubFolder,pSubFolder);
      if (pSubFolder[strlen(pSubFolder)-1] != '\\')
        strcat(pNewSubFolder,"\\");
    }
  }
  return pNewSubFolder;
} /* newSubFolder */

/*--------------------------------------------------------------------*/
char* Fs::newSubFolder(char* pSubFolder, ULONG ulMaxFilenameLength)
{
  char* pNewSubFolder = NULL;
  char* pDisambiguatedSubFolder = this->disambiguate(pSubFolder,ulMaxFilenameLength,NULL);
  if (pDisambiguatedSubFolder != NULL)
  {
    pNewSubFolder = newSubFolder(pDisambiguatedSubFolder);
    delete pDisambiguatedSubFolder;
  }
  return pNewSubFolder;
} /* newSubFolder */

/*--------------------------------------------------------------------*/
char* Fs::normalizedSubFolder(char* pSubFolder, ULONG ulMaxFilenameLength)
{
  Fs* pFsFolder = new Fs(pSubFolder);
  char* pNormalizedFolder = pFsFolder->normalize();
  delete pFsFolder;
  char* pNormalizedSubFolder = this->newSubFolder(pNormalizedFolder,ulMaxFilenameLength);
  free(pNormalizedFolder);
  return pNormalizedSubFolder;
} /* normalizedSubFolder */

/*--------------------------------------------------------------------*/
char* Fs::newFile(char* pFile)
{
  char* pNewFile = NULL;
  if (this->isFolder())
  {
    pNewFile = new char[strlen(m_pName)+strlen(pFile)+1];
    if (pNewFile != NULL)
    {
      strcpy(pNewFile,m_pName);
      strcat(pNewFile,pFile);
    }
  }
  return pNewFile;
} /* newFile */

/*--------------------------------------------------------------------*/
char* Fs::newFile(char* pFile, ULONG ulMaxFilenameLength, char* pExtension)
{
  char* pDisambiguatedFile = this->disambiguate(pFile,ulMaxFilenameLength,pExtension);
  char* pNewFile = newFile(pDisambiguatedFile);
  delete pDisambiguatedFile;
  return pNewFile;
} /* newFile */

/*--------------------------------------------------------------------*/
char* Fs::normalizeFile(char* pFile, ULONG ulMaxFilenameLength, char* pExtension)
{
  Fs* pFsFile = new Fs(pFile);
  char* pNormalizedFile = pFsFile->normalize();
  delete pFsFile;
  char* pNewFile = newFile(pNormalizedFile, ulMaxFilenameLength, pExtension);
  delete pNormalizedFile;
  return pNewFile;
} /* normalizeFile */

/*--------------------------------------------------------------------*/
char* Fs::deriveFile(char* pExtension)
{
  char* pDerivedFile = NULL;
  unsigned u;
  if (!this->isFolder())
  {
    char* pName = _strdup(m_pName);
    /* find last '.', '/' or '\\' */
    for (u = strlen(pName) - 1;
      (u >= 0) && (pName[u] != '.') && (pName[u] != '\\') && (pName[u] != '/');
      u--) {}
    /* detach extension */
    for(;(u >= 0) && (pName[u] == '.');u--)
      pName[u] = '\0';
    u = 0;
    if (pExtension != NULL)
      u = strlen(pExtension);
    char* pTrimmed = trim(pName);
    free(pName);
    pDerivedFile = (char*)malloc(strlen(pTrimmed)+u+1);
    strcpy(pDerivedFile,pTrimmed);
    if (pExtension != NULL)
      strcat(pDerivedFile,pExtension);
    free(pTrimmed);
  }
  return pDerivedFile;
} /* deriveFile */

/*--------------------------------------------------------------------*/
DWORD Fs::setTimes(FileTime* pftCreation, 
                 FileTime* pftModification, 
                 FileTime* pftAccess)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  DWORD dwError = ERROR_BAD_FORMAT;
  wchar_t* pWcName = this->toWideChar();
  if (pWcName != NULL)
  {
    /* create the file */
    HANDLE hFile = CreateFile(pWcName,GENERIC_ALL,0,NULL,OPEN_EXISTING,0,NULL);
    if (hFile != INVALID_HANDLE_VALUE)
    {
      BOOL bSuccess = SetFileTime(hFile,pftCreation->getFt(),pftAccess->getFt(),pftModification->getFt());
      /* close the file */
      CloseHandle(hFile);
    }
    else
      dwError = GetLastError();
    delete pWcName;
  }
#ifdef _DEBUG
  if (!((dwError == 0) && pHeap->isValid() && (pHeap->hasInitialSize())))
    dwError = ERROR_NOT_ENOUGH_MEMORY;
  delete pHeap;
#endif
  return dwError;
} /* setTimes */

/*--------------------------------------------------------------------*/
DWORD Fs::deleteFile()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  DWORD dwError = ERROR_BAD_FORMAT;
  if (!this->isFolder())
  {
    dwError = ERROR_NO_UNICODE_TRANSLATION;
    wchar_t* pWcName = this->toWideChar();
    if (pWcName != NULL)
    {
      if (!DeleteFile(pWcName))
        dwError = GetLastError();
      delete pWcName;
    }
  }
#ifdef _DEBUG
  if (!((dwError == 0) && pHeap->isValid() && (pHeap->hasInitialSize())))
    dwError = ERROR_NOT_ENOUGH_MEMORY;
  delete pHeap;
#endif
  return dwError;
} /* deleteFile */

/*--------------------------------------------------------------------*/
DWORD Fs::createFolder()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  DWORD dwError = ERROR_NO_UNICODE_TRANSLATION;
  /* convert to wchar */
  wchar_t* pWcName = this->toWideChar();
  if (pWcName != NULL)
  {
    /* remove final backslash */
    pWcName[wcslen(pWcName)-1] = 0;
    if (CreateDirectory(pWcName,NULL))
      dwError = 0;
    else
      dwError = GetLastError();
    delete pWcName;
  }
#ifdef _DEBUG
  if (!((dwError == 0) && pHeap->isValid() && (pHeap->hasInitialSize())))
    dwError = ERROR_NOT_ENOUGH_MEMORY;
  delete pHeap;
#endif
  return dwError;
} /* createFolder */

/*--------------------------------------------------------------------*/
DWORD Fs::deleteFolder(BOOL bDeleteContents)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  DWORD dwError = ERROR_BAD_FORMAT;
  if (this->isFolder())
  {
    dwError = ERROR_NO_UNICODE_TRANSLATION;
    wchar_t* pWcName = this->toWideChar();
    if (pWcName != NULL)
    {
      /* remove final backslash (thus doubling terminating zero) */
      pWcName[wcslen(pWcName)-1] = L'\0';
      if (bDeleteContents)
      {
        /* must use SHFileOperation because enumeration through FindFirstFile/FindNextFile
           "locks" enumerated set and prevents removal. */
        SHFILEOPSTRUCT sfo;
        sfo.hwnd = NULL;
        sfo.wFunc = FO_DELETE;
        sfo.pFrom = pWcName;
        sfo.pTo = NULL;
        sfo.fFlags = FOF_SILENT | FOF_NOCONFIRMATION | FOF_NOERRORUI;
        sfo.fAnyOperationsAborted = false;
        sfo.hNameMappings = NULL;
        sfo.lpszProgressTitle = NULL;
        dwError = SHFileOperation(&sfo);
      }
      else
      {
        if (RemoveDirectory(pWcName))
          dwError = 0;
        else
          dwError = GetLastError();
      }
      delete pWcName;
    }
  }
#ifdef _DEBUG
  if (!((dwError == 0) && pHeap->isValid() && (pHeap->hasInitialSize())))
    dwError = ERROR_NOT_ENOUGH_MEMORY;
  delete pHeap;
#endif
  return dwError;
} /* deleteFolder */

/*==End of File=======================================================*/
