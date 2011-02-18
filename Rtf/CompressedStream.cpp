/*-- $Workfile: CompressedStream.cpp $ ---------------------------------
Wrapper for an IStream with compressed RTF.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: CompressedStream class wraps an IStream object containing 
             compressed RTF.
             See MS-OXRTFCP.pdf (http://msdn.microsoft.com/en-us/library/cc463890.aspx)
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
#include "CompressedStream.h"

#define COMPRESSED 0x75465A4C
#define UNCOMPRESSED 0x414C454D

/*--------------------------------------------------------------------*/
void CompressedStream::initialize()
{
  Stream::read(sizeof(ULONG),&m_ulCompressedSize);
  if (m_ulCompressedSize != (Stream::size()-sizeof(ULONG)))
    printf("Unexpected Compressed Size %u (instead of %u)!\n",m_ulCompressedSize,Stream::size());
  Stream::read(sizeof(ULONG),&m_ulSize);
  Stream::read(sizeof(ULONG),&m_ulCompType);
  Stream::read(sizeof(ULONG),&m_ulCrc);
  m_ulCompressedSize = Stream::size() - 16;
  m_bEos = false;
  m_iRead = 0;
  m_iDecoded = 0;
  m_pDecoded = NULL;
  m_pDictionary = new Dictionary();
  m_pCrc = new Crc();
} /* initialize */

/*--------------------------------------------------------------------*/
CompressedStream::CompressedStream(IStream* pStream) : Stream(pStream)
{
  initialize();
} /* CompressedStream */

/*--------------------------------------------------------------------*/
CompressedStream::~CompressedStream()
{
  if (m_pDecoded != NULL)
    free(m_pDecoded);
  if (m_pDictionary != NULL)
    delete m_pDictionary;
  if (m_pCrc != NULL)
    delete m_pCrc;
} /* ~CompressedStream */

/*--------------------------------------------------------------------*/
ULONG CompressedStream::size()
{
  return m_ulSize;
} /* size */

/*--------------------------------------------------------------------*/
BOOL CompressedStream::isCompressed()
{
  return (m_ulCompType == COMPRESSED);
} /* isCompressed */

/*--------------------------------------------------------------------*/
BOOL CompressedStream::readByte(BYTE* pByte)
{
  BOOL bSucceeded = Stream::read(1,pByte);
  if (bSucceeded)
  {
    m_pCrc->digest(*pByte);
    m_ulCompressedSize -= 1;
  }
  return bSucceeded;
} /* readByte */

/*--------------------------------------------------------------------*/
BOOL CompressedStream::decode()
{
  BOOL bSucceeded = false;
  BYTE ctl;
  m_iDecoded = 0;
  m_iRead = 0;
  if (readByte(&ctl))
  {
    bSucceeded = true;
    for (int i = 0; bSucceeded && (i < 8); i++)
    {
      bSucceeded = false;
      if ((ctl & 0x01) == 0x01)
      {
        WORD wReference; /* big endian! */
        if (readByte((BYTE*)&wReference))
        {
          wReference = wReference << 8;
          if (readByte((BYTE*)&wReference))
          {
            /* split */
            int iLength = (wReference & 0x000F) + 2;
            int iOffset = wReference >> 4;
            /* read iLength characters from Dictionary */
            m_iDecoded += iLength;
            m_pDecoded = (char*)realloc(m_pDecoded,m_iDecoded+1);
            if (!m_pDictionary->get(iOffset,iLength,&m_pDecoded[m_iDecoded-iLength]))
            {
              m_iDecoded -= iLength;
              m_bEos = true;
            }
            m_pDecoded[m_iDecoded] = '\0';
            bSucceeded = true;
          }
        }
      }
      else
      {
        char c;
        if (readByte((BYTE*)&c))
        {
          m_iDecoded = m_iDecoded + 1;
          m_pDecoded = (char*)realloc(m_pDecoded,m_iDecoded+1);
          m_pDecoded[m_iDecoded-1] = c;
          m_pDecoded[m_iDecoded] = '\0';
          m_pDictionary->put(c);
          bSucceeded = true;
        }
      }
      ctl = ctl >> 1;
    }
  }
  if (m_bEos)
    bSucceeded = true;
  return bSucceeded;
} /* decode */

/*--------------------------------------------------------------------*/
int CompressedStream::get()
{
  int i = -1;
  if (m_iRead == m_iDecoded)
  {
    if (!m_bEos)
      decode();
  }
  if (m_iRead < m_iDecoded)
  {
    i = m_pDecoded[m_iRead];
    m_iRead++;
  }
  return i;
} /* get */

/*--------------------------------------------------------------------*/
ULONG CompressedStream::read(ULONG ulSize, void* pBuffer)
{
  ULONG ulRead = 0;
  if (ulSize > 0)
  {
    ulRead = -1;
    char* p = (char*)pBuffer;
    if (isCompressed())
    {
      ulRead = 0;
      for (int i = 0; (ulRead < ulSize) && (i != -1);)
      {
        i = get();
        if (i != -1)
        {
          p[ulRead] = (char)i;
          ulRead++;
        }
      }
    }
    else
      ulRead = Stream::read(ulSize,pBuffer);
  }
  return ulRead;
} /* read */

/*--------------------------------------------------------------------*/
BOOL CompressedStream::isCorrupt()
{
  BOOL bCorrupt = true;
  if (isCompressed())
  {
    /* read padding */
    for (BYTE b = 0; m_ulCompressedSize > 0; readByte(&b)) {}
    /* check CRC */
    if (m_pCrc->value() == m_ulCrc)
      bCorrupt = false;
  }
  else
  {
    /* CRC must be 0 and COMPTYPE 0x414C454D */
    if ((m_ulCompType == UNCOMPRESSED) && (m_ulCrc == 0))
      bCorrupt = false;
  }
  return bCorrupt;
} /* isCorrupt */

/*==End of File=======================================================*/
