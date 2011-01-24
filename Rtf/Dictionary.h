/*-- $Workfile: Dictionary.h $ -----------------------------------------
Dictionary for decompressing compressed RTF Streams.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Dictionary class for decompressing compressed RTF Streams.
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
#ifndef DICTIONARY_H
#define DICTIONARY_H

/*======================================================================
Dictionary implements a circular buffer of characters.
======================================================================*/
class Dictionary
{
public:
  /*--------------------------------------------------------------------
  constructor 
  heap size not invariant: dictionary buffer is released by destructor.
  --------------------------------------------------------------------*/
  Dictionary();

  /*--------------------------------------------------------------------
  destructor
  heap size not invariant: dictionary buffer is released by destructor.
  --------------------------------------------------------------------*/
  ~Dictionary();

  /*--------------------------------------------------------------------
  put writes a character to the write position and increments it.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL put(char c);

  /*--------------------------------------------------------------------
  get sets the read position to iOffset and reads iLength characters.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL get(int iOffset, int iLength, char* pBuffer);

/*====================================================================*/
private:

  /*--------------------------------------------------------------------
  inc increments i mod 4096.
  heap size invariant.
  --------------------------------------------------------------------*/
  int inc(int i);

  /*--------------------------------------------------------------------
  initialize allocates and initializes the circular buffer.
  heap size not invariant: dictionary is allocated.
  --------------------------------------------------------------------*/
  void initialize();

  char* m_pDictionary;
  int m_iWritePos;

}; /* class Dictionary */
#endif // DICTIONARY_H not defined
/*==End of File=======================================================*/
