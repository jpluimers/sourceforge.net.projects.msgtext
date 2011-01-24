/*-- $Workfile: CompressedStream.h $ -----------------------------------
Wrapper for an IStream with compressed RTF.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: CompressedStream class wraps an IStream object containing 
             compressed RTF.
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
#ifndef COMPRESSEDSTREAM_H
#define COMPRESSEDSTREAM_H

#include "../Storage/Stream.h"
#include "Dictionary.h"
#include "Crc.h"

/*======================================================================
CompressedStream extends Stream with decompression.
======================================================================*/
class CompressedStream : public Stream
{
public:
  /*--------------------------------------------------------------------
  constructor 
    pStream: pointer to IStream interface.
  heap size not invariant: Dictionary, Crc and buffer allocated.
  --------------------------------------------------------------------*/
  CompressedStream(IStream* pStream);

  /*--------------------------------------------------------------------
  destructor releases IStream interface.
  heap size not invariant: Dictionary, Crc and buffer freed.
  --------------------------------------------------------------------*/
  ~CompressedStream();

  /*--------------------------------------------------------------------
  size returns the (uncompressed) size in bytes or -1 if an error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  virtual ULONG size();

  /*--------------------------------------------------------------------
  isCompressed returns true, if m_CompType == 0x75465a4d
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL isCompressed();

  /*--------------------------------------------------------------------
  read copies portion of the uncompressed stream content to given buffer.
  heap size not invariant: decode buffer may have been reallocated.
  --------------------------------------------------------------------*/
  virtual ULONG read(ULONG ulSize, void* pBuffer);

  /*--------------------------------------------------------------------
  isCorrupt must be called after all bytes have been read
            and returns true, if Crc is correct.  
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL isCorrupt();

/*====================================================================*/
private:
  /*--------------------------------------------------------------------
  initialize reads the header.
  heap size not invariant: Dictionary, Crc and buffer allocated.
  --------------------------------------------------------------------*/
  void initialize();

  /*--------------------------------------------------------------------
  get returns the next character or -1 for EOS.
  heap size not invariant: may have called decode.
  --------------------------------------------------------------------*/
  int get();

  /*--------------------------------------------------------------------
  decode decodes a run.
  heap size not invariant: decoded buffer is reallocated.
  --------------------------------------------------------------------*/
  BOOL decode();

  /*--------------------------------------------------------------------
  readByte reads the next (compressed) byte and updates the Crc.
  heap size.
  --------------------------------------------------------------------*/
  BOOL readByte(BYTE* pByte);

  /* Crc object */
  Crc* m_pCrc;
  /* Dictionary object */
  Dictionary* m_pDictionary;
  ULONG m_ulCompressedSize;
  ULONG m_ulSize;
  ULONG m_ulCompType;
  ULONG m_ulCrc;
  char* m_pDecoded; /* decoded characters */
  int m_iDecoded; /* size of m_pDecoded */
  int m_iRead; /* number of characters already read in m_pDecoded */
  BOOL m_bEos; /* end-of-stream indicator */

}; /* class CompressedStream */
#endif // COMPRESSEDSTREAM_H not defined
/*==End of File=======================================================*/
