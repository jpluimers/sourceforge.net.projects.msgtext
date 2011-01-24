/*-- $Workfile: PropStorage.h $ ----------------------------------------
Class representing a Msg object with Properties in a Storage file.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Class representing a Msg object (Msg, Attachment, Recipient)
             with Properties in a Storage file.
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
#ifndef PROPSTORAGE_H
#define PROPSTORAGE_H

#include "../Utils/FileTime.h"
#include "../Storage/Storage.h"
#include "Property.h"

/*======================================================================
PropStorage extends Storage
======================================================================*/
class PropStorage : public Storage
{
  friend class Msg; // Msg may call protected initialize()

public:

  /*--------------------------------------------------------------------
  constructor 
    pFilename: name of storage file.
  heap size invariant.
  --------------------------------------------------------------------*/
  PropStorage(char* pFilename);

  /*--------------------------------------------------------------------
  destructor releases IStorage interface.
  heap size invariant.
  --------------------------------------------------------------------*/
  ~PropStorage();

  /*--------------------------------------------------------------------
  getPropertyCount 
    returns: the number of properties.
  heap size invariant.
  --------------------------------------------------------------------*/
  int getPropertyCount();

  /*--------------------------------------------------------------------
  getProperty
    ulTag: tag of property
    returns the property or NULL.
  heap size invariant.
  --------------------------------------------------------------------*/
  Property* getProperty(ULONG ulTag);

  /*--------------------------------------------------------------------
  getStringValue
    returns: the string value or the unicode value converted to Multibyte.
  heap size not invariant: caller must delete returned value.
  --------------------------------------------------------------------*/
  char* getStringValue(ULONG ulTag);

  /*--------------------------------------------------------------------
  getSysTimeValue
    returns: the FileTime object representing the SysTime value.
  heap size not invariant: caller must delete returned value.
  --------------------------------------------------------------------*/
  FileTime* getSysTimeValue(ULONG ulTag);

  /*--------------------------------------------------------------------
  listProperties
    lists the properties of this PropStorage.
  heap size invariant.
  --------------------------------------------------------------------*/
  void listProperties();

/*====================================================================*/
protected:
  /*--------------------------------------------------------------------
  constructor 
    pStorage: open IStorage object.
  N.B.: resulting PropStorage must be initialized after construction!
  heap size invariant.
  --------------------------------------------------------------------*/
  PropStorage(IStorage* pStorage, BOOL bUnicode);

  /*--------------------------------------------------------------------
  initialize is called by open().
  It reads all properties and fills the array of properties.
    returns: 0, if no error occurred.
  not heap size invariant: property array is only released in destructor.
  --------------------------------------------------------------------*/
  virtual HRESULT initialize();

  /*--------------------------------------------------------------------
  header analyzes header at pHeader and returns pointer to first property.
  Must be overdefined by derived classes.
  heap size invariant.
  --------------------------------------------------------------------*/
  virtual LPBYTE header(LPBYTE pHeader);

  /* from PR_STORE_SUPPORT_MASK in Msg: */
  BOOL m_bUnicode;

/*====================================================================*/
private:

  /*--------------------------------------------------------------------
  clear clears the contained elements.
  heap size invariant.
  --------------------------------------------------------------------*/
  void clear();

  /* properties */
  Property** m_apProperty;
  int m_iPropertyCount;

}; /* class PropStorage */

#endif // PROPSTORAGE_H not defined
/*==End of File=======================================================*/
