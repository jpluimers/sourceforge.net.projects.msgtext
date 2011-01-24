/*-- $Workfile: Session.h $ --------------------------------------------
Wrapper for an IMAPISession for archiving MSG files.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Session class wraps an IMAPISession.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#ifndef SESSION_H
#define SESSION_H

#define STRICT
#include <mapix.h>
#include "Store.h"
#include "Table.h"

/*======================================================================
Session is the IMAPISession wrapper
======================================================================*/
class Session
{
public:
  /*--------------------------------------------------------------------
  constructor creates table of message stores.
    pMapiSession: pointer to a MAPISESSION interface
  heap size not invariant: internal table will be cleaned up on destruction.
  --------------------------------------------------------------------*/
  Session(LPMAPISESSION pMapiSession);

  /*--------------------------------------------------------------------
  destructor destroys table of message stores, logs off from MAPISESSION 
    and releases it.
  heap size not invariant: internal table is cleaned up.
  --------------------------------------------------------------------*/
  ~Session();

  /*--------------------------------------------------------------------
  stores 
    returns: number of message stores.
  heap size invariant.
  --------------------------------------------------------------------*/
  ULONG stores();

  /*--------------------------------------------------------------------
  getStore return message store of given index.
    ulIndex: index of message store to return.
    returns: the store with given index or NULL, if an error occurred.
  heap size not invariant: caller must delete returned object.
  --------------------------------------------------------------------*/
  Store* getStore(ULONG ulIndex);

  /*--------------------------------------------------------------------
  listMapi lists all MAPI folders unter the given one.
    bEmpty        : true, if empty folders are to be listed.
    pMapiFolder   : Only folders in this folder are to be listed.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT listMapi(
    BOOL bEmpty,
    char* pMapiFolder);

  /*--------------------------------------------------------------------
  listFiles lists all MSG files that would be created on archive.
    bEmpty        : true, if empty folders are to be listed.
    ulMaxFilenameLength: maximum folder/file name length
    pMapiFolder   : Only mails in this folder are to be listed.
    pArchiveFolder: directory, where IPM folders would be archived.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT listFiles(
    BOOL bEmpty,
    char* pMapiFolder,
    ULONG ulMaxFilenameLength,
    char* pArchiveFolder);

  /*--------------------------------------------------------------------
  archive write the session's IPM folders and messages to the given
    directory.
    bEmpty        : true, if empty folders are to be created.
    ulMaxFilenameLength: maximum folder/file name length
    pMapiFolder   : Only mails in this folder are to be archived.
    pArchiveFolder: directory, where IPM folders will be archived.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT archive(
    BOOL bEmpty,
    char* pMapiFolder,
    ULONG ulMaxFilenameLength,
    char* pArchiveFolder);

private:
  /* pointer to IMAPISession interface */
  LPMAPISESSION m_pMapiSession;
  /* table of message stores */
  Table* m_pTableMsgStores;

}; /* class Session */

#endif // SESSION_H not defined
/*==End of File=======================================================*/
