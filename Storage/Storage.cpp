/*-- $Workfile: Storage.cpp $ ------------------------------------------
Wrapper for an IStorage object.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Storage class wraps an IStorage object.
             Used "EnumAll sample" for listing all properties in all 
             property sets:
             http://msdn.microsoft.com/en-us/library/aa379016(VS.85).aspx.
             But beware: That is a very incomplete example!
             Also used dfview.exe, which - unfortunately - is not included
             with Visual Studio any more (but can still be found in MSVC 6.0)!
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
#include <stdio.h>
#define STRICT
#include <windows.h>
#include "../Utils/Fs.h"
#include "../Utils/Heap.h"
#include "Storage.h"

/*--------------------------------------------------------------------*/
void Storage::clear()
{
  m_pFileName = NULL;
  m_pStorage = NULL;
  m_iStorageCount = 0;
  m_apwcsStorage = NULL;
  m_iStreamCount = 0;
  m_apwcsStream = NULL;
} /* clear */

/*--------------------------------------------------------------------*/
int Storage::append(void*** pap, int iCount, void* p)
{
  iCount++;
  *pap = (void**)realloc(*pap,iCount*sizeof(void*));
  (*pap)[iCount-1] = p;
  return iCount;
} /* append */

/*--------------------------------------------------------------------*/
HRESULT Storage::initialize()
{
  IEnumSTATSTG *pEnumContained;
  STATSTG statstg;
  HRESULT hr = m_pStorage->EnumElements(0, NULL, 0, &pEnumContained);
  if (SUCCEEDED(hr))
  {
    statstg.pwcsName = NULL;
    for (hr = pEnumContained->Next(1, &statstg, NULL);
         hr == S_OK;
         hr = pEnumContained->Next(1, &statstg, NULL))
    {
      if (statstg.type == STGTY_STORAGE)
        m_iStorageCount = append((void***)&m_apwcsStorage,m_iStorageCount,statstg.pwcsName);
      else if (statstg.type == STGTY_STREAM)
        m_iStreamCount = append((void***)&m_apwcsStream,m_iStreamCount,statstg.pwcsName);
      else
      {
        hr = E_FAIL;
        printf("Found unexpected child element in IStorage!\nType: %d, Name: %S\n",
          statstg.type,statstg.pwcsName);
      }
    }
    if (!FAILED(hr))
    {
#ifdef _DEBUG
      printf("Content ----------------------------\n");
      Storage::listContent();
#endif
      hr = S_OK;
    }
    pEnumContained->Release();
  }
  else
    printf("IStorage::EnumElements failed (0x%08x)!\n",hr);
  return hr;
} /* initialize */

/*--------------------------------------------------------------------*/
Storage::Storage(char* pName)
{
  clear();
  m_pFileName = pName;
} /* Storage */

/*--------------------------------------------------------------------*/
Storage::Storage(IStorage* pStorage)
{
  clear();
  m_pStorage = pStorage;
} /* Storage */

/*--------------------------------------------------------------------*/
Storage::~Storage()
{
  int iStorage;
  int iStream;
  /* release the names of storage children */
  for (iStorage = 0; iStorage < m_iStorageCount; iStorage++)
    CoTaskMemFree(m_apwcsStorage[iStorage]);
  if (m_apwcsStorage != NULL)
    free(m_apwcsStorage);
  /* release the names of stream children */
  for (iStream = 0; iStream < m_iStreamCount; iStream++)
    CoTaskMemFree(m_apwcsStream[iStream]);
  if (m_apwcsStream != NULL)
    free(m_apwcsStream);
  /* release the IStorage object */
  if (m_pStorage != NULL)
    m_pStorage->Release();
} /* ~Storage */

/*--------------------------------------------------------------------*/
HRESULT Storage::open()
{
  Fs* fs = new Fs(m_pFileName);
  wchar_t* pNameWc = fs->toWideChar();
  delete fs;
  HRESULT hr = StgOpenStorageEx(pNameWc, STGM_READ | STGM_SHARE_EXCLUSIVE,
    STGFMT_DOCFILE, 0, NULL, 0, IID_IStorage, (void**)&m_pStorage);
  if (SUCCEEDED(hr))
    hr = initialize();
  else
    printf("Storage \"%s\" could not be opened (0x%08x)!\n",m_pFileName,hr);
  free(pNameWc);
  return hr;
} /* open */

/*--------------------------------------------------------------------*/
HRESULT Storage::create(GUID clsid)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  Fs* fs = new Fs(m_pFileName);
  wchar_t* pNameWc = fs->toWideChar();
  delete fs;
  HRESULT hr = StgCreateStorageEx(pNameWc,STGM_READWRITE | STGM_TRANSACTED | STGM_CREATE,
    STGFMT_DOCFILE, 0, NULL, NULL, IID_IStorage, (void**)&m_pStorage);
  if (SUCCEEDED(hr))
  {
    hr = WriteClassStg(m_pStorage, clsid);
    if (FAILED(hr))
    {
      printf("CLSID could not be written to \"%s\" (0x%08x)!\n",m_pFileName,hr);
      m_pStorage->Release();
      m_pStorage = NULL;
    }
  }
  else
  {
    printf("Storage \"%s\" could not be created (0x%08x)!\n",m_pFileName,hr);
    if (m_pStorage != NULL)
      m_pStorage->Release();
    m_pStorage = NULL;
  }
  delete pNameWc;
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && pHeap->hasInitialSize()))
  {
    hr = E_FAIL;
    printf("Storage::create: Heap inconsistent!\n");
  }
  delete pHeap;
#endif
  return hr;
} /* create */

/*--------------------------------------------------------------------*/
HRESULT Storage::commit()
{
  HRESULT hr = m_pStorage->Commit(STGC_DEFAULT);
  return hr;
} /* commit */

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
  if (!(SUCCEEDED(hr) && pHeap->isValid() && pHeap->hasInitialSize()))
  {
    bMatches = false;
    printf("Storage::matches: Heap inconsistent after ReadClassStg()!\n");
  }
  delete pHeap;
#endif
  return bMatches;
} /* matches */

/*--------------------------------------------------------------------*/
LPOLESTR Storage::name()
{
  LPOLESTR pwcsName = NULL;
  STATSTG ststg;
  HRESULT hr =  m_pStorage->Stat(&ststg,STATFLAG_DEFAULT);
  if (SUCCEEDED(hr))
    pwcsName = ststg.pwcsName;
  return pwcsName;
} /* name */

/*--------------------------------------------------------------------*/
DWORD Storage::type()
{
  DWORD dwType = 0;
  STATSTG ststg;
  HRESULT hr =  m_pStorage->Stat(&ststg,STATFLAG_DEFAULT);
  if (SUCCEEDED(hr))
    dwType = ststg.type;
  return dwType;
} /* type */

/*--------------------------------------------------------------------*/
int Storage::getStorageCount()
{
  return m_iStorageCount;
} /* getStorageCount */

/*--------------------------------------------------------------------*/
LPOLESTR Storage::getStorageName(int iStorage)
{
  LPOLESTR pwcsName = NULL;
  if ((iStorage >= 0) && (iStorage < m_iStorageCount))
    pwcsName = m_apwcsStorage[iStorage];
  return pwcsName;
} /* getStorageName */

/*--------------------------------------------------------------------*/
IStorage* Storage::getIStorage(LPOLESTR pwcsName)
{
  IStorage* pStorage = NULL;
  HRESULT hr = m_pStorage->OpenStorage(pwcsName, NULL, 
    STGM_READ | STGM_SHARE_EXCLUSIVE, NULL, 0, &pStorage);
  if (FAILED(hr))
    printf("IStorage::OpenStorage() failed (0x%08x)!\n",hr);
  return pStorage;
} /* getIStorage */

/*--------------------------------------------------------------------*/
Storage* Storage::getStorage(LPOLESTR pwcsName)
{
  Storage* pStorage = NULL;
  IStorage* pStorageContained = getIStorage(pwcsName);
  if (pStorageContained != NULL)
  {
    pStorage = new Storage(pStorageContained);
    pStorage->initialize();
  }
  return pStorage;
} /* getStorage */

/*--------------------------------------------------------------------*/
int Storage::getStreamCount()
{
  return m_iStreamCount;
} /* getStreamCount */

/*--------------------------------------------------------------------*/
LPOLESTR Storage::getStreamName(int iStream)
{
  LPOLESTR pwcsName = NULL;
  if ((iStream >= 0) && (iStream < m_iStreamCount))
    pwcsName = m_apwcsStream[iStream];
  return pwcsName;
} /* getStreamName */

/*--------------------------------------------------------------------*/
IStream* Storage::getIStream(LPOLESTR pwcsName)
{
  IStream* pStream = NULL;
  HRESULT hr = m_pStorage->OpenStream(pwcsName, NULL, 
    STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &pStream);
  if (FAILED(hr))
    printf("IStorage::OpenStream() failed (0x%08x)!\n",hr);
  return pStream;
} /* getIStream */

/*--------------------------------------------------------------------*/
Stream* Storage::getStream(LPOLESTR pwcsName)
{
  Stream* pStream = NULL;
  IStream* pStreamContained = getIStream(pwcsName);
  if (pStreamContained != NULL)
    pStream = new Stream(pStreamContained);
  return pStream;
} /* getStream */

/*--------------------------------------------------------------------*/
HRESULT Storage::copyStorage(LPOLESTR pwcsStorage, Storage* pStorage)
{
  HRESULT hr = m_pStorage->MoveElementTo(pwcsStorage,
    pStorage->m_pStorage,pwcsStorage,STGMOVE_COPY);
  if (SUCCEEDED(hr))
  {
    IStorage* pCopiedStorage;
    hr = pStorage->m_pStorage->OpenStorage(pwcsStorage, NULL, 
      STGM_READ | STGM_SHARE_EXCLUSIVE, NULL, 0, &pCopiedStorage);
    if (SUCCEEDED(hr))
    {
      STATSTG statstg;
      hr = pCopiedStorage->Stat(&statstg,STATFLAG_DEFAULT);
      if (SUCCEEDED(hr))
      {
        pStorage->m_iStorageCount = pStorage->append((void***)&pStorage->m_apwcsStorage,
          pStorage->m_iStorageCount,statstg.pwcsName);
      }
      pCopiedStorage->Release();
    }
  }
  return hr;
} /* copyStorage */

/*--------------------------------------------------------------------*/
HRESULT Storage::copyStream(LPOLESTR pwcsStream, Storage* pStorage)
{
  HRESULT hr = m_pStorage->MoveElementTo(pwcsStream,
    pStorage->m_pStorage,pwcsStream,STGMOVE_COPY);
  if (SUCCEEDED(hr))
  {
    IStream* pCopiedStream;
    hr = pStorage->m_pStorage->OpenStream(pwcsStream, NULL, 
      STGM_READ | STGM_SHARE_EXCLUSIVE, 0, &pCopiedStream);
    if (SUCCEEDED(hr))
    {
      STATSTG statstg;
      hr = pCopiedStream->Stat(&statstg,STATFLAG_DEFAULT);
      if (SUCCEEDED(hr))
      {
        hr = pCopiedStream->Release();
        if (SUCCEEDED(hr))
        {
          pStorage->m_iStreamCount = pStorage->append((void***)&pStorage->m_apwcsStream,
            pStorage->m_iStreamCount,statstg.pwcsName);
        }
      }
    }
  }
  return hr;
} /* copyStream */

/*--------------------------------------------------------------------*/
Stream* Storage::createStream(LPOLESTR pwcsStream)
{
  Stream* pStream = NULL;
  IStream* pCreatedStream;
  HRESULT hr = m_pStorage->CreateStream(pwcsStream, 
    STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_FAILIFTHERE,
    0,0,&pCreatedStream);
  if (SUCCEEDED(hr))
  {
    STATSTG statstg;
    hr = pCreatedStream->Stat(&statstg,STATFLAG_DEFAULT);
    if (SUCCEEDED(hr))
    {
      m_iStreamCount = append((void***)&m_apwcsStream,m_iStreamCount,statstg.pwcsName);
      pStream = new Stream(pCreatedStream);
    }
  }
  return pStream;
} /* createStream */

/*--------------------------------------------------------------------*/
BOOL Storage::listContent()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  BOOL bSucceeded = true;
  int iStorage;
  int iStream;
  Stream* pstreamChild;
  void* pBuffer;
  LPOLESTR pwcsName;
  for (iStorage = 0; bSucceeded && (iStorage < m_iStorageCount); iStorage++)
  {
    bSucceeded = false;
    pwcsName = m_apwcsStorage[iStorage];
    if (pwcsName != NULL)
    {
      printf("+ %S\n",pwcsName);
      bSucceeded = true;
    }
  }
  for (iStream = 0; bSucceeded && (iStream < m_iStreamCount); iStream++)
  {
    bSucceeded = false;
    pwcsName = m_apwcsStream[iStream];
    if (pwcsName != NULL)
    {
      printf("- %S\n",pwcsName);
      pstreamChild = getStream(pwcsName);
      if (pstreamChild != NULL)
      {
        pBuffer = pstreamChild->content();
        free(pBuffer);
        delete pstreamChild;
        bSucceeded = true;
      }
    }
  }
#ifdef _DEBUG
  if (!(bSucceeded && pHeap->isValid() && pHeap->hasInitialSize()))
  {
    bSucceeded = false;
    printf("Storage::listContent: Heap inconsistent after listContent()!\n");
  }
  delete pHeap;
#endif
  return bSucceeded;
} /* listContent */

/*==End of File=======================================================*/
