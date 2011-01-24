/*-- $Workfile: Recipient.cpp $ ----------------------------------------
Class representiong a Recipient object in a Storage file.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Recipient class extends a PropStorage object.
             See MS-OXMSG (http://msdn.microsoft.com/en-us/library/cc463912.aspx)
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
#include "../Utils/Heap.h"
#include "Recipient.h"

/*--------------------------------------------------------------------*/
LPBYTE Recipient::header(LPBYTE pHeader)
{
  /* 2 4-byte values:
    dwReserved1,
    dwReserved2.
  */
  pHeader += 8;
  return pHeader;
} /* header */

/*--------------------------------------------------------------------*/
Recipient::Recipient(IStorage* pStorage, BOOL bUnicode) 
  : PropStorage(pStorage, bUnicode)
{
} /* Recipient */

/*--------------------------------------------------------------------*/
Recipient::~Recipient()
{
} /* ~Recipient */

/*--------------------------------------------------------------------*/
ULONG Recipient::getType()
{
  ULONG ulType = MAPI_ORIG;
  Property* pPropType = getProperty(PR_RECIPIENT_TYPE);
  if (pPropType != NULL)
    ulType = pPropType->getLongValue();
  return ulType;
} /* getType */

/*--------------------------------------------------------------------*/
char* Recipient::getDisplayName()
{
  return getStringValue(PR_DISPLAY_NAME);
} /* getDisplayName */

/*--------------------------------------------------------------------*/
char* Recipient::getEmailAddress()
{
  return getStringValue(PR_EMAIL_ADDRESS);
} /* getEmailAddress */

/*==End of File=======================================================*/
