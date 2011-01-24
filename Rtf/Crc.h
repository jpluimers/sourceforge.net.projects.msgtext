/*-- $Workfile: Crc.h $ ------------------------------------------------
Compute a CRC for an IStream with compressed RTF.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Crc class computes a CRC for an IStream with compressed RTF.
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
#ifndef CRC_H
#define CRC_H

/*======================================================================
Crc class computes a CRC for an IStream with compressed RTF.
======================================================================*/
class Crc
{
public:
  /*--------------------------------------------------------------------
  constructor starts CRC with 0 value.
  heap size invariant.
  --------------------------------------------------------------------*/
  Crc();

  /*--------------------------------------------------------------------
  destructor
  heap size invariant.
  --------------------------------------------------------------------*/
  ~Crc();

  /*--------------------------------------------------------------------
  value returns current CRC value.
  --------------------------------------------------------------------*/
  ULONG value();

  /*--------------------------------------------------------------------
  digest computes new CRC.
  heap size invariant.
  --------------------------------------------------------------------*/
  void digest(BYTE c);

/*====================================================================*/
private:
  ULONG m_ulCrc;

}; /* class Crc */

#endif // CRC_H not defined
/*==End of File=======================================================*/
