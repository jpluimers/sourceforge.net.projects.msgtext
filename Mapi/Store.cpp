/*-- $Workfile: Store.cpp $ --------------------------------------------
Wrapper for an IMsgStore.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Store class wraps an IMsgStore.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#include <stdio.h>
#define STRICT
#include <mapix.h>
#include "../Utils/Fs.h"
#include "../Utils/Heap.h"
#include "Store.h"

/*--------------------------------------------------------------------*/
Store::Store(IMsgStore* pMsgStore) 
  : Properties((IMAPIProp*)pMsgStore)
{
} /* Store */

/*--------------------------------------------------------------------*/
Store::~Store()
{
} /* ~Store */

/*--------------------------------------------------------------------*/
char* Store::getDisplayName()
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
BOOL Store::getDefaultStore()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  BOOL bDefaultStore = FALSE;
  PropValue* pPv = this->getPropValue(PR_DEFAULT_STORE);
  if (pPv != NULL)
  {
    bDefaultStore = pPv->getBooleanValue();
    delete pPv;
  }
#ifdef _DEBUG
  if (!(pHeap->isValid() && (pHeap->hasInitialSize())))
    bDefaultStore = false;
  delete pHeap;
#endif
  return bDefaultStore;
} /* getDefaultStore */

/*--------------------------------------------------------------------*/
Folder* Store::getIpmRootFolder()
{
  Folder* pFolder = NULL;
  IMAPIFolder* pMapiFolder = NULL;
  ULONG ulObjType;
  /* get the PR_IPM_SUBTREE_ENTRYID property */
  PropValue* pPv = this->getPropValue(PR_IPM_SUBTREE_ENTRYID);
  if (pPv != NULL)
  {
    LPENTRYID pEntryId = (LPENTRYID)pPv->getBinaryPointer();
    ULONG ulLength = pPv->getBinaryLength();
    IMsgStore* pMsgStore = (IMsgStore*)this->m_pMapiProp;
    HRESULT hr = pMsgStore->OpenEntry(ulLength,pEntryId,&IID_IMAPIFolder,0,&ulObjType,(LPUNKNOWN*)&pMapiFolder);
    if ((!FAILED(hr)) && (ulObjType == MAPI_FOLDER))
      pFolder = new Folder(pMapiFolder);
    delete pPv;
  }
  return pFolder;
} /* getIpmRootFolder */

/*--------------------------------------------------------------------*/
HRESULT Store::list()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = NOERROR;
  /* list display name and default */
  char* pDisplayName = this->getDisplayName();
  printf("Message Store: %s%s\n",pDisplayName,this->getDefaultStore()?" (default)":"");
  free(pDisplayName);
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
    hr = E_FAIL;
  delete pHeap;
#endif
  return hr;
} /* list */

/*--------------------------------------------------------------------*/
HRESULT Store::listMapi(BOOL bEmpty, char* pMapiFolder)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = NOERROR;
  Folder* pRootFolder = this->getIpmRootFolder();
  if (bEmpty || !pRootFolder->isEmpty())
  {
    /* the store's display name is like the root folder's name */
    char* pDisplayName = this->getDisplayName();
    char* pMapiPath = NULL;
    if ((pMapiFolder == NULL) || (strcmp(pMapiFolder,pDisplayName) == 0))
    {
      printf("%s\n",pDisplayName);
      pMapiPath = pDisplayName;
      pMapiFolder = NULL;
    }
    hr = pRootFolder->listMapi(bEmpty, pMapiFolder, pMapiPath);
    free(pDisplayName);
  }
  delete pRootFolder;
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
    hr = E_FAIL;
  delete pHeap;
#endif
  return hr;
} /* listMapi */

/*--------------------------------------------------------------------*/
HRESULT Store::listFiles(BOOL bEmpty, char* pMapiFolder,
  ULONG ulMaxFilenameLength, char* pArchiveFolder)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = NOERROR;
  Folder* pRootFolder = this->getIpmRootFolder();
  if (bEmpty || !pRootFolder->isEmpty())
  {
    /* the store's display name is like the root folder's name */
    char* pDisplayName = this->getDisplayName();
    char* pSubFolder = pArchiveFolder;
    DWORD dwError = 0;
    if ((pMapiFolder == NULL) || (strcmp(pMapiFolder,pDisplayName) == 0))
    {
      Fs* pFsArchive = new Fs(pArchiveFolder);
      pSubFolder = pFsArchive->normalizedSubFolder(pDisplayName, ulMaxFilenameLength);
      delete pFsArchive;
      /* folder creation is needed for listFile because of disambiguation */
      Fs* pFsSubFolder = new Fs(pSubFolder);
      dwError = pFsSubFolder->createFolder();
      delete pFsSubFolder;
      if (dwError == 0)
        printf("%s\n",pSubFolder);
      pMapiFolder = NULL;
    }
    if (dwError == 0)
      hr = pRootFolder->listFiles(bEmpty, pMapiFolder, pDisplayName,
        ulMaxFilenameLength, pSubFolder);
    if (pSubFolder != pArchiveFolder)
      free(pSubFolder);
    free(pDisplayName);
  }
  delete pRootFolder;
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
    hr = E_FAIL;
  delete pHeap;
#endif
  return hr;
} /* listFiles */

/*--------------------------------------------------------------------*/
HRESULT Store::archive(BOOL bEmpty, char* pMapiFolder, 
  ULONG ulMaxFilenameLength, char* pArchiveFolder)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = NOERROR;
  Folder* pRootFolder = this->getIpmRootFolder();
  if (bEmpty || !pRootFolder->isEmpty())
  {
    /* the store's display name is like the root folder's name */
    char* pDisplayName = this->getDisplayName();
    char* pSubFolder = pArchiveFolder;
    DWORD dwError = 0;
    if ((pMapiFolder == NULL) || (strcmp(pMapiFolder,pDisplayName) == 0))
    {
      Fs* pFsArchive = new Fs(pArchiveFolder);
      pSubFolder = pFsArchive->normalizedSubFolder(pDisplayName, ulMaxFilenameLength);
      delete pFsArchive;
      /* folder creation is needed for listFile because of disambiguation */
      Fs* pFsSubFolder = new Fs(pSubFolder);
      dwError = pFsSubFolder->createFolder();
      delete pFsSubFolder;
      if (dwError == 0)
        printf("%s\n",pSubFolder);
      pMapiFolder = NULL;
    }
    if (dwError == 0)
      hr = pRootFolder->archive(bEmpty, pMapiFolder, pDisplayName,
        ulMaxFilenameLength, pSubFolder);
    if (pSubFolder != pArchiveFolder)
      free(pSubFolder);
    free(pDisplayName);
  }
  delete pRootFolder;
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
    hr = E_FAIL;
  delete pHeap;
#endif
  return hr;
} /* archive */

/*==End of File=======================================================*/
