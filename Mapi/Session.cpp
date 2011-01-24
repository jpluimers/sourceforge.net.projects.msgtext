/*-- $Workfile: Session.cpp $ ------------------------------------------
Wrapper for an IMAPISession for archiving MSG files.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Session class wraps an IMAPISession.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#include <stdio.h>
#define STRICT
#include <mapix.h>
#include "../Utils/Heap.h"
#include "../Utils/Fs.h"
#include "Session.h"
#include "Table.h"
#include "Store.h"

/*--------------------------------------------------------------------*/
Session::Session(LPMAPISESSION pMapiSession)
{
  LPMAPITABLE pMapiTable;
  m_pMapiSession = pMapiSession;
  HRESULT hr = m_pMapiSession->GetMsgStoresTable(0,&pMapiTable);
  if (!FAILED(hr))
    m_pTableMsgStores = new Table(pMapiTable);
  else
    printf("GetMsgStoresTable failed!\n");
} /* Session */

/*--------------------------------------------------------------------*/
Session::~Session()
{
  if (m_pTableMsgStores != NULL)
    delete m_pTableMsgStores;
  m_pMapiSession->Logoff(NULL,NULL,NULL);
  m_pMapiSession->Release();
} /* ~Session */

/*--------------------------------------------------------------------*/
ULONG Session::stores()
{
  ULONG ulStores = -1;
  if (m_pTableMsgStores)
    ulStores = m_pTableMsgStores->entries();
  return ulStores;
} /* getStores */

/*--------------------------------------------------------------------*/
Store* Session::getStore(ULONG ulIndex)
{
  Store* pStore = NULL;
  if (m_pTableMsgStores)
  {
    IMsgStore* pMsgStore;
    ULONG ulEntryIdLength = m_pTableMsgStores->getEntryIdLength(ulIndex);
    LPENTRYID pEntryId = m_pTableMsgStores->getEntryId(ulIndex);
    HRESULT hr = m_pMapiSession->OpenMsgStore(NULL, 
      ulEntryIdLength, 
      pEntryId, NULL, 0, &pMsgStore);
    if (!FAILED(hr))
      pStore = new Store(pMsgStore);
    else
      printf("OpenMsgStore failed!\n");
  }
  return pStore;
} /* getStore */

/*--------------------------------------------------------------------*/
HRESULT Session::listMapi(
  BOOL bEmpty,
  char* pMapiFolder)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = NOERROR;
  if (m_pTableMsgStores)
  {
    for (ULONG ulStore = 0; (ulStore < this->stores()) && (!FAILED(hr)); ulStore++)
    {
      Store* pStore = this->getStore(ulStore);
      if (pStore)
      {
        hr = pStore->listMapi(bEmpty, pMapiFolder);
        delete pStore;
      }
      else
        printf("Store object not instantiated!\n");
    }
  }
  else
    printf("m_pTableMsgStores not initialized!\n");
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
    hr = E_FAIL;
  delete pHeap;
#endif
  return hr;
} /* listMapi */

/*--------------------------------------------------------------------*/
HRESULT Session::listFiles(
  BOOL bEmpty,
  char* pMapiFolder,
  ULONG ulMaxFilenameLength,
  char* pArchiveFolder)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = NOERROR;
  if (m_pTableMsgStores)
  {
    for (ULONG ulStore = 0; (ulStore < this->stores()) && (!FAILED(hr)); ulStore++)
    {
      Store* pStore = this->getStore(ulStore);
      if (pStore)
      {
        hr = pStore->listFiles(bEmpty, pMapiFolder, ulMaxFilenameLength, pArchiveFolder);
        delete pStore;
      }
      else
        printf("Store object not instantiated!\n");
    }
  }
  else
    printf("m_pTableMsgStores not initialized!\n");
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
    hr = E_FAIL;
  delete pHeap;
#endif
  return hr;
} /* listFiles */

/*--------------------------------------------------------------------*/
HRESULT Session::archive(
  BOOL bEmpty,
  char* pMapiFolder,
  ULONG ulMaxFilenameLength,
  char* pArchiveFolder)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = NOERROR;
  if (m_pTableMsgStores)
  {
    for (ULONG ulStore = 0; (ulStore < this->stores()) && (!FAILED(hr)); ulStore++)
    {
      Store* pStore = this->getStore(ulStore);
      if (pStore)
      {
        hr = pStore->archive(bEmpty, pMapiFolder, ulMaxFilenameLength, pArchiveFolder);
        delete pStore;
      }
      else
        printf("Store object not instantiated!\n");
    }
  }
  else
    printf("m_pTableMsgStores not initialized!\n");
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
    hr = E_FAIL;
  delete pHeap;
#endif
  return hr;
} /* archive */

/*==End of File=======================================================*/
