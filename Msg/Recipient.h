/*-- $Workfile: Recipient.h $ ------------------------------------------
Class representiong a Recipient object in a Storage file.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Recipient class extends a PropStorage object.
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
#ifndef RECIPIENT_H
#define RECIPIENT_H

#include "PropStorage.h"

/*======================================================================
Recipient extends PropStorage
======================================================================*/
class Recipient : public PropStorage
{
public:
  /*--------------------------------------------------------------------
  constructor 
    pStorage: open IStorage object.
  N.B.: resulting Recipient must be initialized after construction!
  heap size invariant.
  --------------------------------------------------------------------*/
  Recipient(IStorage* pStorage, BOOL bUnicode);

  /*--------------------------------------------------------------------
  destructor releases IStorage interface.
  heap size invariant.
  --------------------------------------------------------------------*/
  ~Recipient();

  /*--------------------------------------------------------------------
  getType
    returns: MAPI_ORIG, MAPI_TO, MAPI_CC, or MAPI_BCC
  heap size not invariant: Caller must free returned string.
  --------------------------------------------------------------------*/
  ULONG getType();

  /*--------------------------------------------------------------------
  getDisplayName
    returns: the display name of the recipient.
  heap size not invariant: Caller must free returned string.
  --------------------------------------------------------------------*/
  char* getDisplayName();

  /*--------------------------------------------------------------------
  getEmailAddress
    returns: the email address of the recipient.
  heap size not invariant: Caller must free returned string.
  --------------------------------------------------------------------*/
  char* getEmailAddress();

/*====================================================================*/
private:

  /*--------------------------------------------------------------------
  header analyzes header at pHeader and returns pointer to first property.
  heap size invariant.
  --------------------------------------------------------------------*/
  virtual LPBYTE header(LPBYTE pHeader);

}; /* class Recipient */

#endif // RECIPIENT_H not defined
/*==End of File=======================================================*/
