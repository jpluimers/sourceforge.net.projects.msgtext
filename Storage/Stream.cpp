/*-- $Workfile: Stream.cpp $ -------------------------------------------
Wrapper for an IStream.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Stream class wraps an IStream.
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
#include "../Utils/Heap.h"
#include "Stream.h"

/*--------------------------------------------------------------------*/
Stream::Stream(IStream* pStream)
{
  m_pStream = pStream;
} /* Stream */

/*--------------------------------------------------------------------*/
Stream::~Stream()
{
  if (m_pStream != NULL)
    m_pStream->Release();
} /* ~Stream */

/*--------------------------------------------------------------------*/
LPOLESTR Stream::name()
{
  LPOLESTR pwcsName = NULL;
  STATSTG ststg;
  HRESULT hr =  m_pStream->Stat(&ststg,STATFLAG_DEFAULT);
  if (SUCCEEDED(hr))
    pwcsName = ststg.pwcsName;
  return pwcsName;
} /* name */

/*--------------------------------------------------------------------*/
ULONG Stream::size()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  ULONG ulSize = -1;
  STATSTG ststg;
  HRESULT hr =  m_pStream->Stat(&ststg,STATFLAG_DEFAULT);
  if (SUCCEEDED(hr) && (ststg.cbSize.HighPart == 0))
  {
    ulSize = ststg.cbSize.LowPart;
    CoTaskMemFree(ststg.pwcsName);
  }
#ifdef _DEBUG
  if (!((ulSize >= 0) && pHeap->isValid() && pHeap->hasInitialSize()))
  {
    printf("Stream::size: Heap inconsistent after size()!\n");
  }
  delete pHeap;
#endif
  return ulSize;
} /* size */

/*--------------------------------------------------------------------*/
HRESULT Stream::commit()
{
  HRESULT hr = m_pStream->Commit(STGC_DEFAULT);
  return hr;
} /* commit */

/*--------------------------------------------------------------------*/
ULONG Stream::read(ULONG ulSize, void* pBuffer)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  ULONG ulRead = -1;
  HRESULT hr = m_pStream->Read(pBuffer,ulSize,&ulRead);
  if (FAILED(hr))
    ulRead = -1;
#ifdef _DEBUG
  if (!((ulRead != -1) && pHeap->isValid() && pHeap->hasInitialSize()))
  {
    printf("Stream::read: Heap inconsistent after read()!\n");
    ulRead = -1;
  }
  delete pHeap;
#endif
  return ulRead;
} /* read */

/*--------------------------------------------------------------------*/
BOOL Stream::read(ULONG ulOffset, ULONG ulSize, void* pBuffer)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  BOOL bSucceeded = false;
  LARGE_INTEGER i64;
  i64.LowPart = ulOffset;
  i64.HighPart = 0;
  HRESULT hr = m_pStream->Seek(i64, STREAM_SEEK_SET, NULL);
  if (SUCCEEDED(hr))
    bSucceeded = (read(ulSize,pBuffer) == ulSize);
#ifdef _DEBUG
  if (!(bSucceeded && pHeap->isValid() && pHeap->hasInitialSize()))
  {
    printf("Stream::read: Heap inconsistent after read()!\n");
    bSucceeded = false;
  }
  delete pHeap;
#endif
  return bSucceeded;
} /* read */

/*--------------------------------------------------------------------*/
BOOL Stream::write(ULONG ulSize, void* pBuffer)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  BOOL bSucceeded = false;
  ULONG ulWritten = -1;
  HRESULT hr = m_pStream->Write(pBuffer,ulSize,&ulWritten);
  if (SUCCEEDED(hr) && (ulSize == ulWritten))
    bSucceeded = true;
#ifdef _DEBUG
  if (!(bSucceeded && pHeap->isValid() && pHeap->hasInitialSize()))
  {
    printf("Stream::write: Heap inconsistent after write()!\n");
    bSucceeded = false;
  }
  delete pHeap;
#endif
  return bSucceeded;
} /* write */

/*--------------------------------------------------------------------*/
LPBYTE Stream::content()
{
  LPBYTE pContent = NULL;
  ULONG ulSize = size();
  if (ulSize >= 0)
  {
    pContent = (LPBYTE)malloc(ulSize);
    if (pContent != NULL)
    {
      if (!read(0,ulSize,pContent))
      {
        free(pContent);
        pContent = NULL;
      }
    }
  }
  return pContent;
} /* content */

/*==End of File=======================================================*/
