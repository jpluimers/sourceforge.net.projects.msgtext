/*-- $Workfile: Attachment.h $ -----------------------------------------
Class representiong an Attachment object in a Storage file.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Attachment class extends a PropStorage object.
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
#ifndef ATTACHMENT_H
#define ATTACHMENT_H

#include "PropStorage.h"

/*======================================================================
Attachment extends PropStorage
======================================================================*/
class Attachment : public PropStorage
{
  friend class Msg;

public:
  /*--------------------------------------------------------------------
  constructor 
    pStorage: open IStorage object.
  N.B.: resulting Attachment must be initialized after construction!
  heap size invariant.
  --------------------------------------------------------------------*/
  Attachment(IStorage* pStorage, BOOL bUnicode);

  /*--------------------------------------------------------------------
  destructor releases IStorage interface.
  heap size invariant.
  --------------------------------------------------------------------*/
  ~Attachment();

  /*--------------------------------------------------------------------
  isEmbeddedMsg returns true, if attachment is an embedded Msg.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL isEmbeddedMsg();

  /*--------------------------------------------------------------------
  getFilename returns the file name of the attachment.
  heap size not invariant: Caller must free returned file name.
  --------------------------------------------------------------------*/
  char* getFilename();

  /*--------------------------------------------------------------------
  save writes the content of the attachment to a file.
  heap size not invariant: System allocates irrecoverable space on heap
                           for stream.
  --------------------------------------------------------------------*/
  BOOL save(char* pFilename);

/*====================================================================*/
protected:

  /*--------------------------------------------------------------------
  getMsgStore returns a pointer to the IStorage interface of the embedded
  Msg object, if the attachment is an embedded Msg, NULL otherwise.
  heap size not invariant: Caller must release returned IStorage*.
  --------------------------------------------------------------------*/
  IStorage* getMsgStorage();

/*====================================================================*/
private:

  /*--------------------------------------------------------------------
  header analyzes header at pHeader and returns pointer to first property.
  heap size invariant.
  --------------------------------------------------------------------*/
  virtual LPBYTE header(LPBYTE pHeader);

  /*--------------------------------------------------------------------
  getMethod returns the value of PR_ATTACH_METHOD property:
  NO_ATTACHMENT, ATTACH_BY_VALUE, ATTACH_BY_REFERENCE,
  ATTACH_BY_REF_RESOLVE, ATTACH_BY_REF_ONLY, 
  ATTACH_EMBEDDED_MSG, or ATTACH_OLE
  We can only really handle NO_ATTACHMENT, ATTACH_BY_VALUE, 
  ATTACH_EMBEDDED_MSG and ATTACH_OLE.
  heap size invariant.
  --------------------------------------------------------------------*/
  ULONG getMethod();

}; /* class Attachment */

#endif // ATTACHMENT_H not defined
/*==End of File=======================================================*/
