/*-- $Workfile: Storage.cpp $ ------------------------------------------
Wrapper for a STORAGE object.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Storage class wraps a STORAGE object.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#include <stdio.h>
#define STRICT
#include <mapix.h>
#include "../Utils/Fs.h"
#include "../Utils/Heap.h"
#include "Storage.h"

/*--------------------------------------------------------------------*/
Storage::Storage(char* pName)
{
  m_pName = pName;
  m_pStorage = NULL;
} /* Storage */

/*--------------------------------------------------------------------*/
Storage::~Storage()
{
  if (m_pStorage != NULL)
    m_pStorage->Release();
} /* ~Storage */

/*--------------------------------------------------------------------*/
LPSTORAGE Storage::getStg()
{
  return m_pStorage;
} /* getStg */

/*--------------------------------------------------------------------*/
HRESULT Storage::create(GUID clsid)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  Fs* fs = new Fs(m_pName);
  wchar_t* pNameWc = fs->toWideChar();
  delete fs;
  HRESULT hr = StgCreateStorageEx(pNameWc,STGM_READWRITE | STGM_TRANSACTED | STGM_CREATE,
    STGFMT_DOCFILE, 0, NULL, NULL, IID_IStorage, (void**)&m_pStorage);
  if (SUCCEEDED(hr))
  {
    hr = WriteClassStg(m_pStorage, clsid);
    if (FAILED(hr))
    {
      printf("CLSID could not be written to \"%s\" (0x%08x)!\n",m_pName,hr);
      m_pStorage->Release();
      m_pStorage = NULL;
    }
  }
  else
  {
    printf("Storage \"%s\" could not be created (0x%08x)!\n",m_pName,hr);
    if (m_pStorage != NULL)
      m_pStorage->Release();
    m_pStorage = NULL;
  }
  delete pNameWc;
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
  {
    hr = E_FAIL;
    printf("Storage::create: Heap inconsistent!\n");
  }
  delete pHeap;
#endif
  return hr;
} /* create */

/*--------------------------------------------------------------------*/
HRESULT Storage::open()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  Fs* fs = new Fs(m_pName);
  wchar_t* pNameWc = fs->toWideChar();
  delete fs;
  HRESULT hr = StgOpenStorageEx(pNameWc, STGM_READ | STGM_SHARE_EXCLUSIVE,
    STGFMT_DOCFILE, 0, NULL, 0, IID_IStorage, (void**)&m_pStorage);
  if (FAILED(hr))
    printf("Storage \"%s\" could not be opened (0x%08x)!\n",m_pName);
  delete pNameWc;
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
  {
    hr = E_FAIL;
    printf("Storage::open: Heap inconsistent!\n");
  }
  delete pHeap;
#endif
  return hr;
} /* open */

/*--------------------------------------------------------------------*/
BOOL Storage::matches(GUID clsid)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  GUID clsidRead;
  HRESULT hr = ReadClassStg(m_pStorage, &clsidRead);
  BOOL bMatches = (memcmp(&clsid,&clsidRead,sizeof(GUID)) == 0);
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
  {
    bMatches = false;
    printf("Storage::matches: Heap inconsistent after ReadClassStg()!\n");
  }
  delete pHeap;
#endif
  return bMatches;
} /* matches */

/*--------------------------------------------------------------------*/
HRESULT Storage::commit()
{
  HRESULT hr = m_pStorage->Commit(STGC_DEFAULT);
  return hr;
} /* commit */

/*==End of File=======================================================*/
