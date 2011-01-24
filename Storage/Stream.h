/*-- $Workfile: Stream.h $ ---------------------------------------------
Wrapper for an IStream.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Stream class wraps an IStream object.
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
#ifndef STREAM_H
#define STREAM_H

/*======================================================================
Stream is the IStream wrapper
======================================================================*/
class Stream
{
public:
  /*--------------------------------------------------------------------
  constructor 
    pStream: pointer to IStream interface.
  heap size invariant.
  --------------------------------------------------------------------*/
  Stream(IStream* pStream);

  /*--------------------------------------------------------------------
  destructor releases IStream interface.
  heap size invariant.
  --------------------------------------------------------------------*/
  ~Stream();

  /*--------------------------------------------------------------------
  name returns the friendly name of the stream.
  not heap size invariant: Caller must free returned name using CoTaskMemFree.
  --------------------------------------------------------------------*/
  LPOLESTR name();

  /*--------------------------------------------------------------------
  size returns the size in bytes or -1 if an error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  virtual ULONG size();

  /*--------------------------------------------------------------------
  commit commits the changes.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT commit();

  /*--------------------------------------------------------------------
  read copies portion of stream content to given buffer.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL read(ULONG ulOffset, ULONG ulSize, void* pBuffer);

  /*--------------------------------------------------------------------
  read copies portion of the uncompressed stream content to given buffer.
  heap size invariant.
  --------------------------------------------------------------------*/
  virtual ULONG read(ULONG ulSize, void* pBuffer);

  /*--------------------------------------------------------------------
  write copies a given buffer to the stream.
  not heap size invariant.
  --------------------------------------------------------------------*/
  BOOL write(ULONG ulSize, void* pBuffer);

  /*--------------------------------------------------------------------
  content allocates a buffer of the necessary size and copies the content
  of the stream to it.
  not heap size invariant: returned buffer must be freed by caller.
  --------------------------------------------------------------------*/
  LPBYTE content();

/*====================================================================*/
protected:
  /* pointer to IStream interface */
  IStream* m_pStream;
}; /* class Stream */

#endif // STREAM_H not defined
/*==End of File=======================================================*/
