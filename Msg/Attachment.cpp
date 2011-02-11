/*-- $Workfile: Attachment.cpp $ ---------------------------------------
Class representiong an Attachment object in a Storage file.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Attachment class extends a PropStorage object.
             See MS-OXMSG (http://msdn.microsoft.com/en-us/library/cc463912.aspx)
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
#include <stdio.h>
#define STRICT
#include <windows.h>
#include "../Utils/Fs.h"
#include "../Utils/Heap.h"
#include "Attachment.h"

/*--------------------------------------------------------------------*/
ULONG Attachment::getMethod()
{
  ULONG ulMethod = ATTACH_BY_VALUE;
  Property* pMethod = this->getProperty(PR_ATTACH_METHOD);
  if (pMethod != NULL)
    ulMethod = pMethod->getLongValue();
  return ulMethod;
} /* getMethod */

/*--------------------------------------------------------------------*/
LPBYTE Attachment::header(LPBYTE pHeader)
{
  /* 2 4-byte values:
    dwReserved1,
    dwReserved2,
  */
  pHeader += 8;
  return pHeader;
} /* header */

/*--------------------------------------------------------------------*/
Attachment::Attachment(IStorage* pStorage, BOOL bUnicode) 
  : PropStorage(pStorage, bUnicode)
{
} /* Atachment */

/*--------------------------------------------------------------------*/
Attachment::~Attachment()
{
} /* ~Attachment */

/*--------------------------------------------------------------------*/
BOOL Attachment::isEmbeddedMsg()
{
  return (getMethod() == ATTACH_EMBEDDED_MSG);
} /* isEmbeddedMsg */

/*--------------------------------------------------------------------*/
char* Attachment::getFilename()
{
  char* pFilename = NULL;
  char* pName = getStringValue(PR_ATTACH_LONG_FILENAME);
  if ((pName == NULL) || (*pName == '\0'))
  {
    pName = getStringValue(PR_ATTACH_FILENAME);
    if ((pName == NULL) || (*pName == '\0'))
    {
      pName = getStringValue(PR_DISPLAY_NAME);
      if ((pName != NULL) && (*pName != '\0'))
      {
        // printf("--- DisplayName: %s\n",pName);
        /* detach size */
        int i = strlen(pName) - 1;
        if (pName[i] == ')')
        {
          for (i--; (i >= 0) && (pName[i] != '('); i--) {}
          pName[i] = '\0';
        }
        // printf("--- DisplayName without size: %s\n",pName);
      }
      else
      {
        pName = _strdup("No name");
        printf("No name!");
      }
    }
    //else
    //  printf("--- AttachFilename: %s\n",pName);
  }
  //else
  //  printf("--- AttachLongFilename: %s\n",pName);
  pFilename = trim(pName);
  free(pName);
  // printf("--- Trimmed file name: %s\n",pFilename);
  /* append extension */
  char* pExtension = NULL;
  if (getMethod() == ATTACH_EMBEDDED_MSG)
    pExtension = _strdup(".msg");
  else
    pExtension = getStringValue(PR_ATTACH_EXTENSION);
  if ((pExtension != NULL) && (*pExtension != '\0'))
  {
    // printf("--- Extension: %s\n",pExtension);
    pFilename = (char*)realloc(pFilename, strlen(pFilename)+strlen(pExtension)+1);
    strcat(pFilename,pExtension);
    // printf("--- Filename with extension: %s\n",pFilename);
    free(pExtension);
  }
  /* replace invalid characters: special characters and ':*?\/"<>|' by spaces */
  char* pSrc;
  char* pDest = pFilename;
  for (pSrc = pFilename; (pSrc != NULL) && (*pSrc != '\0'); pSrc++)
  {
    if (((unsigned char)*pSrc >= ' ') &&
        (*pSrc != ':') && (*pSrc != '*') &&  (*pSrc != '?') &&
        (*pSrc != '\\') && (*pSrc != '/') && (*pSrc != '"') &&
        (*pSrc != '<') && (*pSrc != '>') && (*pSrc != '|'))
      *pDest = *pSrc;
    else
      *pDest = ' ';
    pDest++;
  }
  *pDest = '\0';
  printf("Attachment file name: %s\n",pFilename);
  return pFilename;
} /* getFilename */

/*--------------------------------------------------------------------*/
BOOL Attachment::save(char* pFilename)
{
  BOOL bSucceeded = false;
  /* PR_ATTACH_DATA_BIN */
  Property* pPropAttachData = getProperty(PR_ATTACH_DATA_BIN);
  if (pPropAttachData != NULL)
  {
    Stream* pAttachStream = pPropAttachData->getBinaryStream();
    if (pAttachStream != NULL)
    {
      FILE* pfileAttach = fopen(pFilename,"wb");
      if (pfileAttach != NULL)
      {
        char* pBuf = (char*)malloc(BUFSIZ);
        setvbuf(pfileAttach,pBuf,_IOFBF,BUFSIZ);
#ifdef _DEBUG
        Heap* pHeap = new Heap();
#endif
        bSucceeded = true;
        LPBYTE pBuffer = (LPBYTE)malloc(BUFSIZ);
        ULONG ulRead = BUFSIZ;
        for (ULONG ulSize = pAttachStream->size(); bSucceeded && (ulSize > 0); ulSize -= ulRead)
        {
          if (ulRead > ulSize)
            ulRead = ulSize;
          bSucceeded = pAttachStream->read(ulRead,pBuffer);
          if (bSucceeded)
            bSucceeded = (fwrite(pBuffer,1,ulRead,pfileAttach) == ulRead);
        }
        free(pBuffer);
#ifdef _DEBUG
        if (!(bSucceeded && pHeap->isValid() && pHeap->hasInitialSize()))
        {
          printf("Attachment::save: Heap inconsistent in save()!\n");
          bSucceeded = false;
        }
        delete pHeap;
#endif
        fclose(pfileAttach);
        free(pBuf);
      }
      delete pAttachStream;
    }
  }
  else
    printf("Attachment::save: PR_ATTACH_DATA_BIN could not be opened!\n");
  return bSucceeded;
} /* save */

/*--------------------------------------------------------------------*/
IStorage* Attachment::getMsgStorage()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  IStorage* pAttachment = NULL;
  Property* pPropAttachData = getProperty(PR_ATTACH_DATA_OBJ);
  if (pPropAttachData != NULL)
    pAttachment = pPropAttachData->getBinaryIStorage();
#ifdef _DEBUG
  if (!((pAttachment != NULL) && pHeap->isValid() && pHeap->hasInitialSize()))
  {
    printf("Attachment::getMsgStorage: Heap inconsistent after getMsgStorage()!\n");
    if (pAttachment != NULL)
      pAttachment->Release();
    pAttachment = NULL;
  }
  delete pHeap;
#endif
  return pAttachment;
} /* getMsgStorage */

/*==End of File=======================================================*/
