/*-- $Workfile: Storage.h $ --------------------------------------------
Wrapper for a "Structured Storage" or "Doc" file.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Storage class wraps a STORAGE object.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#ifndef STORAGE_H
#define STORAGE_H

#define STRICT
#include <mapix.h>
// {00020D0B-0000-0000-C000-000000000046}
EXTERN_C DEFINE_OLEGUID(CLSID_MailMessage, 0x00020D0B, 0, 0);

/*======================================================================
Storage is the STORAGE wrapper
======================================================================*/
class Storage
{
public:
  /*--------------------------------------------------------------------
  constructor 
    pFilename: name of storage file.
  heap size invariant.
  --------------------------------------------------------------------*/
  Storage(char* pFilename);

  /*--------------------------------------------------------------------
  destructor releases STORAGE interface.
  heap size invariant.
  --------------------------------------------------------------------*/
  ~Storage();

  /*--------------------------------------------------------------------
  getStg
    returns: pointer to STORAGE interface.
  heap size invariant.
  --------------------------------------------------------------------*/
  LPSTORAGE getStg();

  /*--------------------------------------------------------------------
  create creates an empty storage file and writes the CLSID to it.
    clsid: CLSID to be stored in storage.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT create(GUID clsid);

  /*--------------------------------------------------------------------
  open opens an existing storage file.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT open();

  /*--------------------------------------------------------------------
  matches checks, whether CLSID matches.
    returns: true, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL matches(GUID clsid);

  /*--------------------------------------------------------------------
  commit commits the changes.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT commit();

private:
  /* name of storage file */
  char* m_pName;
  /* pointer to STORAGE interface */
  LPSTORAGE m_pStorage;

}; /* class Storage */

#endif // STORAGE_H not defined
/*==End of File=======================================================*/
