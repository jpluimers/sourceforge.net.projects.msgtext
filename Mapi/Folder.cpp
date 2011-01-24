/*-- $Workfile: Folder.cpp $ -------------------------------------------
Wrapper for an IMAPIFolder.
Version    : $Id$
             $NoKeywords: $
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Folder class wraps an IMAPIFolder.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#define _CRT_SECURE_NO_WARNINGS
#include <stdio.h>
#define STRICT
#include <mapix.h>
#include "../Utils/Heap.h"
#include "../Utils/Fs.h"
#include "Folder.h"

/*--------------------------------------------------------------------*/
Folder::Folder(IMAPIFolder* pMapiFolder)
  : Container((IMAPIContainer*)pMapiFolder)
{
  m_pTableFolders = NULL;
  m_pTableMessages = NULL;
  /* get hierarchy table */
  LPMAPITABLE pHierarchyTable;
  HRESULT hr = pMapiFolder->GetHierarchyTable(0,&pHierarchyTable);
  if (!FAILED(hr))
    m_pTableFolders = new Table(pHierarchyTable);
  else
    printf("Folder::Folder: GetHierarchyTable failed!\n");
  /* get contents table */
  LPMAPITABLE pContentsTable;
  hr = pMapiFolder->GetContentsTable(0,&pContentsTable);
  if (!FAILED(hr))
    m_pTableMessages = new Table(pContentsTable);
  else
    printf("Folder::Folder: GetContentsTable failed!\n");
} /* Folder */

/*--------------------------------------------------------------------*/
Folder::~Folder()
{
  if (m_pTableFolders != NULL)
    delete m_pTableFolders;
  if (m_pTableMessages != NULL)
    delete m_pTableMessages;
} /* ~Folder */

/*--------------------------------------------------------------------*/
char* Folder::getDisplayName()
{
  char* pDisplayName = NULL;
  PropValue* pPv = this->getPropValue(PR_DISPLAY_NAME_A);
  if (pPv != NULL)
  {
    pDisplayName = pPv->getString8Value();
    delete pPv;
  }
  return pDisplayName;
} /* getDisplayName */

/*--------------------------------------------------------------------*/
ULONG Folder::folders()
{
  ULONG ulFolders = 0;
  if (m_pTableFolders)
    ulFolders = m_pTableFolders->entries();
  return ulFolders;
} /* folders */

/*--------------------------------------------------------------------*/
Folder* Folder::getFolder(ULONG ulIndex)
{
  Folder* pFolder = NULL;
  IMAPIFolder* pThisMapiFolder = (IMAPIFolder*)this->m_pMapiProp;
  IMAPIFolder* pMapiFolder;
  ULONG ulType;
  ULONG ulLength = m_pTableFolders->getEntryIdLength(ulIndex);
  LPENTRYID pEntryId = m_pTableFolders->getEntryId(ulIndex);
  HRESULT hr = pThisMapiFolder->OpenEntry(ulLength, pEntryId, 
    NULL, 0, &ulType, (LPUNKNOWN*)&pMapiFolder);
  if (!FAILED(hr) && (ulType == MAPI_FOLDER))
    pFolder = new Folder(pMapiFolder);
  else
    printf("Folder::getFolder: OpenEntry failed for subfolder!\n");
  return pFolder;
} /* getFolder */

/*--------------------------------------------------------------------*/
ULONG Folder::messages()
{
  ULONG ulMessages = 0;
  if (m_pTableMessages)
    ulMessages = m_pTableMessages->entries();
  return ulMessages;
} /* messages */

/*--------------------------------------------------------------------*/
Message* Folder::getMessage(ULONG ulIndex)
{
  Message* pMessage = NULL;
  IMAPIFolder* pThisMapiFolder = (IMAPIFolder*)this->m_pMapiProp;
  IMessage* pMsg;
  ULONG ulType;
  if (ulIndex < this->messages())
  {
    ULONG ulLength = m_pTableMessages->getEntryIdLength(ulIndex);
    LPENTRYID pEntryId = m_pTableMessages->getEntryId(ulIndex);
    HRESULT hr = pThisMapiFolder->OpenEntry(ulLength, pEntryId, 
      NULL, 0, &ulType, (LPUNKNOWN*)&pMsg);
    if (!FAILED(hr)  && (ulType == MAPI_MESSAGE))
      pMessage = new Message(pMsg);
    else
      printf("Folder::getMessage: OpenEntry failed in for message!\n");
  }
  return pMessage;
} /* getMessage */

/*--------------------------------------------------------------------*/
BOOL Folder::isEmpty()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  /* it would be more efficient to cache this property ... */
  BOOL bEmpty = TRUE;
  for (ULONG ulMessage = 0; bEmpty && (ulMessage < this->messages()); ulMessage++)
  {
    Message* pMessage = this->getMessage(ulMessage);
    if (pMessage != NULL)
    {
      if (pMessage->isIpmNote())
        bEmpty = FALSE;
      delete pMessage;
    }
  }
  for (ULONG ulFolder = 0; bEmpty && (ulFolder < this->folders()); ulFolder++)
  {
    Folder* pFolder = this->getFolder(ulFolder);
    if (pFolder != NULL)
    {
      if (!pFolder->isEmpty())
        bEmpty = FALSE;
      delete pFolder;
    }
  }
#ifdef _DEBUG
  if (!(pHeap->isValid() && (pHeap->hasInitialSize())))
    printf("Folder::isEmpty: Inconsistent heap!\n");
  delete pHeap;
#endif
  return bEmpty;
} /* isEmpty */

/*--------------------------------------------------------------------*/
HRESULT Folder::list()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = NOERROR;
  char* pDisplayName = this->getDisplayName();
  printf("Message Folder: %s, Subfolders: %u, Objects: %u\n",
    pDisplayName,this->folders(),this->messages());
  free(pDisplayName);
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
    hr = E_FAIL;
  delete pHeap;
#endif
  return hr;
} /* list */

/*--------------------------------------------------------------------*/
HRESULT Folder::listMapi(BOOL bEmpty, char* pMapiFolder, char* pMapiPath)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = NOERROR;
  char* pDisplayName;
  char* pDisplayPath;
  char* pMapiSubPath;
  for (ULONG ulFolder = 0; SUCCEEDED(hr) && (ulFolder < this->folders()); ulFolder++)
  {
    Folder* pFolder = this->getFolder(ulFolder);
    if ((pFolder != NULL) && (bEmpty || !pFolder->isEmpty()))
    {
      pDisplayName = pFolder->getDisplayName();
      if (pMapiPath != NULL)
      {
        pDisplayPath = (char*)malloc(strlen(pMapiPath)+1+strlen(pDisplayName)+1);
        strcpy(pDisplayPath,pMapiPath);
        strcat(pDisplayPath,"/");
        strcat(pDisplayPath,pDisplayName);
      }
      else
        pDisplayPath = pDisplayName;
      pMapiSubPath = NULL;
      if ((pMapiFolder == NULL) ||(strcmp(pMapiFolder,pDisplayPath) == 0))
      {
        printf("%s\n",pDisplayPath);
        pMapiSubPath = pDisplayPath;
        pMapiFolder = NULL;
      }
      hr = pFolder->listMapi(bEmpty, pMapiFolder, pMapiSubPath);
      if (pDisplayPath != pDisplayName)
        free(pDisplayPath);
      free(pDisplayName);
    }
    if (pFolder != NULL)
      delete pFolder;
  }
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
    hr = E_FAIL;
  delete pHeap;
#endif
  return hr;
} /* listMapi */

/*--------------------------------------------------------------------*/
HRESULT Folder::listFiles(BOOL bEmpty, char* pMapiFolder, char* pMapiPath,
  ULONG ulMaxFilenameLength, char* pFolderPath)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = NOERROR;
  char* pDisplayName;
  char* pDisplayPath;
  char* pSubFolderPath;
  char* pSubMapiFolder;
  DWORD dwError;
  if (pMapiFolder == NULL)
  {
    /* list all msg files */
    for (ULONG ulMessage = 0; SUCCEEDED(hr) && (ulMessage < this->messages()); ulMessage++)
    {
      Message* pMessage = this->getMessage(ulMessage);
      if (pMessage != NULL)
      {
        if (pMessage->isIpmNote())
          hr = pMessage->listFile(ulMaxFilenameLength,pFolderPath);
        delete pMessage;
      }
    }
  }
  /* list files in sub folders */
  for (ULONG ulFolder = 0; SUCCEEDED(hr) && (ulFolder < this->folders()); ulFolder++)
  {
    Folder* pFolder = this->getFolder(ulFolder);
    if ((pFolder != NULL) && (bEmpty || !pFolder->isEmpty()))
    {
      /* pDisplayPath from pDisplayName and pMapiPath */
      pDisplayName = pFolder->getDisplayName();
      if (pMapiPath != NULL)
      {
        pDisplayPath = (char*)malloc(strlen(pMapiPath)+1+strlen(pDisplayName)+1);
        strcpy(pDisplayPath,pMapiPath);
        strcat(pDisplayPath,"/");
        strcat(pDisplayPath,pDisplayName);
      }
      else
        pDisplayPath = pDisplayName;
      pSubFolderPath = pFolderPath;
      pSubMapiFolder = pMapiFolder;
      dwError = 0;
      if ((pMapiFolder == NULL) ||(strcmp(pMapiFolder,pDisplayPath) == 0))
      {
        /* create pSubFolder from pFolder and pDisplayName */
        Fs* pFsFolder = new Fs(pFolderPath);
        pSubFolderPath = pFsFolder->normalizedSubFolder(pDisplayName, ulMaxFilenameLength);
        delete pFsFolder;
        /* folder creation is needed for listFile because of disambiguation */
        Fs* pFsSubFolder = new Fs(pSubFolderPath);
        dwError = pFsSubFolder->createFolder();
        delete pFsSubFolder;
        if (dwError == 0)
          printf("%s\n",pSubFolderPath);
        pSubMapiFolder = NULL;
      }
      if (dwError == 0)
        hr = pFolder->listFiles(bEmpty, pSubMapiFolder, pDisplayPath, ulMaxFilenameLength, pSubFolderPath);
      if (pSubFolderPath != pFolderPath)
        free(pSubFolderPath);
      if (pDisplayPath != pDisplayName)
        free(pDisplayPath);
      free(pDisplayName);
    }
    if (pFolder != NULL)
      delete pFolder;
  }
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
    hr = E_FAIL;
  delete pHeap;
#endif
  return hr;
} /* listFiles */

/*--------------------------------------------------------------------*/
HRESULT Folder::archive(BOOL bEmpty, char* pMapiFolder, char* pMapiPath,
  ULONG ulMaxFilenameLength, char* pFolderPath)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = NOERROR;
  char* pDisplayName;
  char* pDisplayPath;
  char* pSubFolderPath;
  char* pSubMapiFolder;
  DWORD dwError;
  if (pMapiFolder == NULL)
  {
    /* list all msg files */
    for (ULONG ulMessage = 0; SUCCEEDED(hr) && (ulMessage < this->messages()); ulMessage++)
    {
      Message* pMessage = this->getMessage(ulMessage);
      if (pMessage != NULL)
      {
        if (pMessage->isIpmNote())
          hr = pMessage->archive(ulMaxFilenameLength,pFolderPath);
        delete pMessage;
      }
    }
  }
  /* list files in sub folders */
  for (ULONG ulFolder = 0; SUCCEEDED(hr) && (ulFolder < this->folders()); ulFolder++)
  {
    Folder* pFolder = this->getFolder(ulFolder);
    if ((pFolder != NULL) && (bEmpty || !pFolder->isEmpty()))
    {
      /* pDisplayPath from pDisplayName and pMapiPath */
      pDisplayName = pFolder->getDisplayName();
      if (pMapiPath != NULL)
      {
        pDisplayPath = (char*)malloc(strlen(pMapiPath)+1+strlen(pDisplayName)+1);
        strcpy(pDisplayPath,pMapiPath);
        strcat(pDisplayPath,"/");
        strcat(pDisplayPath,pDisplayName);
      }
      else
        pDisplayPath = pDisplayName;
      pSubFolderPath = pFolderPath;
      pSubMapiFolder = pMapiFolder;
      dwError = 0;
      if ((pMapiFolder == NULL) ||(strcmp(pMapiFolder,pDisplayPath) == 0))
      {
        /* create pSubFolder from pFolder and pDisplayName */
        Fs* pFsFolder = new Fs(pFolderPath);
        pSubFolderPath = pFsFolder->normalizedSubFolder(pDisplayName, ulMaxFilenameLength);
        delete pFsFolder;
        /* folder creation is needed for listFile because of disambiguation */
        Fs* pFsSubFolder = new Fs(pSubFolderPath);
        dwError = pFsSubFolder->createFolder();
        delete pFsSubFolder;
        pSubMapiFolder = NULL;
      }
      if (dwError == 0)
        hr = pFolder->archive(bEmpty, pSubMapiFolder, pDisplayPath, ulMaxFilenameLength, pSubFolderPath);
      if (pSubFolderPath != pFolderPath)
        free(pSubFolderPath);
      if (pDisplayPath != pDisplayName)
        free(pDisplayPath);
      free(pDisplayName);
    }
    if (pFolder != NULL)
      delete pFolder;
  }
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
    hr = E_FAIL;
  delete pHeap;
#endif
  return hr;
} /* archive */

/*==End of File=======================================================*/
