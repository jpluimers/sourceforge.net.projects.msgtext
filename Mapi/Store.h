/*-- $Workfile: Store.h $ ----------------------------------------------
Wrapper for an IMsgStore.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Store class wraps an IMsgStore.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#ifndef STORE_H
#define STORE_H

#define STRICT
#include <mapix.h>
#include "Folder.h"

/*======================================================================
Store is the IMsgStore wrapper
======================================================================*/
class Store : public Properties
{
public:
  /*--------------------------------------------------------------------
  constructor 
    pMsgStore: pointer to IMsgStore interface.
  heap size not invariant: Properties establishes list of PropValues.
  --------------------------------------------------------------------*/
  Store(IMsgStore* pMsgStore);

  /*--------------------------------------------------------------------
  destructor releases the IMsgStore (implicitly through ~Properties).
  heap size not invariant: ~Properties cleans up list of PropValues.
  --------------------------------------------------------------------*/
  ~Store();

  /*--------------------------------------------------------------------
  list lists the display name of this store.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT list();

  /*--------------------------------------------------------------------
  listMapi lists all MAPI folders unter the given one.
    bEmpty        : true, if empty folders are to be listed.
    pMapiFolder   : Only mails in this folder are to be listed.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT listMapi(BOOL bEmpty, char* pMapiFolder);

  /*--------------------------------------------------------------------
  listFiles lists all MSG files that would be created in archive.
    bEmpty        : true, if empty folders are to be listed.
    pMapiFolder   : Only mails in this folder are to be listed.
    ulMaxFilenameLength: maximum folder/file name length
    pArchiveFolder: directory, where IPM folders would be archived.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT listFiles(BOOL bEmpty, char* pMapiFolder, 
    ULONG ulMaxFilenameLength, char* pArchiveFolder);

  /*--------------------------------------------------------------------
  archive write the store's IPM folders and messages to the given
    directory.
    bEmpty        : true, if empty folders are to be created.
    ulMaxFilenameLength: maximum folder/file name length
    pMapiFolder   : Only mails in this folder are to be archived.
    pArchiveFolder: directory, where IPM folders will be archived.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT archive(BOOL bEmpty, char* pMapiFolder, 
    ULONG ulMaxFilenameLength, char* pArchiveFolder);

private:
  /*--------------------------------------------------------------------
  getIpmRootFolder returns the IPM (IPM = interpersonal messages) root folder.
    returns: folder object or NULL, if an error occurred.
  heap size not invariant: caller must delete folder object.
  --------------------------------------------------------------------*/
  Folder* getIpmRootFolder();

  /*--------------------------------------------------------------------
  getDisplayName returns the display name of the message store.
    returns: display name of the message store.
  heap size not invariant: caller must delete display name.
  --------------------------------------------------------------------*/
  char* Store::getDisplayName();

  /*--------------------------------------------------------------------
  getDefaultStore returns true, if this message store is the default
    message store.
    returns: true, if this message store is the default message store.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL Store::getDefaultStore();

}; /* class Store */

#endif // STORE_H not defined
/*==End of File=======================================================*/
