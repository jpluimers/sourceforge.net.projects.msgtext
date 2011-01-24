/*-- $Workfile: Storage.h $ --------------------------------------------
Wrapper for a "Structured Storage" or "Doc" file.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Storage class wraps an IStorage object.
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
#ifndef STORAGE_H
#define STORAGE_H

#include "Stream.h"

// {00020D0B-0000-0000-C000-000000000046}
EXTERN_C DEFINE_OLEGUID(CLSID_MailMessage, 0x00020D0B, 0, 0);

/*======================================================================
Storage is the IStorage wrapper
======================================================================*/
class Storage
{

  friend class Property;

public:
  /*--------------------------------------------------------------------
  constructor 
    pFilename: name of storage file.
  heap size invariant.
  --------------------------------------------------------------------*/
  Storage(char* pFilename);

  /*--------------------------------------------------------------------
  destructor releases IStorage interface.
  heap size invariant.
  --------------------------------------------------------------------*/
  ~Storage();

  /*--------------------------------------------------------------------
  open opens an existing storage file.
    returns: 0, if no error occurred.
  not heap size invariant: arrays will only be released in destructor.
  --------------------------------------------------------------------*/
  HRESULT open();

  /*--------------------------------------------------------------------
  create creates an empty storage file and writes the CLSID to it.
    clsid: CLSID to be stored in storage.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT create(GUID clsid);

  /*--------------------------------------------------------------------
  commit commits the changes.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT commit();

  /*--------------------------------------------------------------------
  matches checks, whether CLSID matches.
    returns: true, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL matches(GUID clsid);

  /*--------------------------------------------------------------------
  name 
    returns: the name of the IStorage as a wide character string.
  not heap size invariant: caller must delete returned name using CoTaskMemFree.
  --------------------------------------------------------------------*/
  LPOLESTR name();

  /*--------------------------------------------------------------------
  type
    returns: the type of the IStorage (0=Error, 1=Storage, 2=Stream, 
  3=LockBytes, 4=Property).
  heap size invariant.
  --------------------------------------------------------------------*/
  DWORD type();

  /*--------------------------------------------------------------------
  getStorageCount
  returns: the number of contained Storage objects
  heap size invariant.
  --------------------------------------------------------------------*/
  int getStorageCount();

  /*--------------------------------------------------------------------
  getStorageName
    iStorage: index of contained Storage object
    returns: name of one contained Storage object.
  heap size invariant.
  --------------------------------------------------------------------*/
  LPOLESTR getStorageName(int iStorage);

  /*--------------------------------------------------------------------
  getStorage
    pwcsStorage: name of contained Storage object
    returns: open Storage object of that name or NULL.
  not heap size invariant: Caller must free returned object.
  --------------------------------------------------------------------*/
  Storage* getStorage(LPOLESTR pwcsStorage);

  /*--------------------------------------------------------------------
  getStreamCount 
    returns: the number of contained Stream objects
  heap size invariant.
  --------------------------------------------------------------------*/
  int getStreamCount();

  /*--------------------------------------------------------------------
  getStreamName 
    iStream: index of contained Stream object
    returns name of one contained Stream object.
  heap size invariant.
  --------------------------------------------------------------------*/
  LPOLESTR getStreamName(int iStream);

  /*--------------------------------------------------------------------
  getStream
    pwcsStream: name of contained Stream object
    returns: open Stream object of that name or NULL.
  not heap size invariant: Caller must free returned object.
  --------------------------------------------------------------------*/
  Stream* getStream(LPOLESTR pwcsStream);

  /*--------------------------------------------------------------------
  copyStorage
  copies the contained Storage object to the target Storage.
    pwcsStorage: name on contained Storage object.
    pStorage: target Storage.
    returns S_OK or an error indication.
  not heap size invariant, but added internal state will be removed on delete.
  --------------------------------------------------------------------*/
  HRESULT copyStorage(LPOLESTR pwcsStorage, Storage* pStorage);

  /*--------------------------------------------------------------------
  copyStream
  copies the contained Stream object to the target Storage.
    pwcsStream: name on contained Stream object.
    pStorage: target Storage.
    returns S_OK or an error indication.
  not heap size invariant, but added internal state will be removed on delete.
  --------------------------------------------------------------------*/
  HRESULT copyStream(LPOLESTR pwcsStream, Storage* pStorage);

  /*--------------------------------------------------------------------
  createStream
  create e new contained Stream object in this Storage.
    pwcsStream: name of new contained Stream object.
    returns the created Stream object.
  not heap size invariant, but added internal state will be removed on delete.
  --------------------------------------------------------------------*/
  Stream* createStream(LPOLESTR pwcsStream);

  /*--------------------------------------------------------------------
  listContent
  Lists the directly contained Storage and Stream objects.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL listContent();

/*====================================================================*/
protected:
  /*--------------------------------------------------------------------
  append an pointer to an array of pointers.
    pap: pointer to array of pointers (in/out).
    iCount: current size of array.
    p: new pointer to be added to array.
    returns: new size of array.
  heap size invariant, unless pap is NULL:
    Caller must free one copy of the array eventually.
  --------------------------------------------------------------------*/
  int append(void*** pap, int iCount, void* p);

  /*--------------------------------------------------------------------
  getIStorage
    pwcsStorage: name of contained IStorage interface.
    returns: open IStorage interface of that name or NULL.
  not heap size invariant: Caller must release returned interface.
  --------------------------------------------------------------------*/
  IStorage* getIStorage(LPOLESTR pwcsStorage);

  /*--------------------------------------------------------------------
  getIStream
    pwcsStream: name of contained IStream interface.
    returns: open IStream interface of that name or NULL.
  not heap size invariant: Caller must release returned interface.
  --------------------------------------------------------------------*/
  IStream* getIStream(LPOLESTR pwcsStorage);

  /*--------------------------------------------------------------------
  constructor 
    pStorage: open IStorage object.
  N.B.: resulting Storage must be initialized after construction!
  heap size invariant.
  --------------------------------------------------------------------*/
  Storage(IStorage* pStorage);

  /*--------------------------------------------------------------------
  initialize is called by open().
  It reads all directly contained Storage and Stream objects and fills 
  the name arrays.
    returns: 0, if no error occurred.
  not heap size invariant: arrays are only released in destructor.
  --------------------------------------------------------------------*/
  virtual HRESULT initialize();

/*====================================================================*/
private:
  /*--------------------------------------------------------------------
  clear clears the contained elements.
  heap size invariant.
  --------------------------------------------------------------------*/
  void clear();

  /* name of storage file */
  char* m_pFileName;
  /* pointer to STORAGE interface */
  IStorage* m_pStorage;

  /* list of names of contained Storage objects */
  LPOLESTR* m_apwcsStorage;
  /* number of contained Storage objects */
  int m_iStorageCount;
  /* list of names of contained Stream objects */
  LPOLESTR* m_apwcsStream;
  /* number of contained Stream objects */
  int m_iStreamCount;

}; /* class Storage */

#endif // STORAGE_H not defined
/*==End of File=======================================================*/
