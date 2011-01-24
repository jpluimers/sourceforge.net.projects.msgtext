/*-- $Workfile: Properties.h $ -----------------------------------------
Wrapper for an IMAPIProp.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Properties class wraps IMAPIProp.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#ifndef PROPERTIES_H
#define PROPERTIES_H

#define STRICT
#include <mapix.h>
#include "PropValue.h"

/*======================================================================
Properties is the IMAPIProp wrapper
======================================================================*/
class Properties
{
public:
  /*--------------------------------------------------------------------
  constructor 
    pMapiProp: pointer to IMAPIProp interface can be NULL,
               no properties are needed.
  heap size not invariant: establishes list of PropValues.
  --------------------------------------------------------------------*/
  Properties(IMAPIProp* pMapiProp);

  /*--------------------------------------------------------------------
  destructor releases IMAPIProp interface.
  heap size not invariant: cleans up list of PropValues.
  --------------------------------------------------------------------*/
  ~Properties();

  /*--------------------------------------------------------------------
  getCount returns number of PropValues.
    returns: number of PropValues
  heap size invariant.
  --------------------------------------------------------------------*/
  ULONG getCount();

  /*--------------------------------------------------------------------
  getTag returns tag of PropValue of given index.
    ulIndex: index of PropValue.
    returns: tag of PropValue of given index
  heap size invariant.
  --------------------------------------------------------------------*/
  ULONG getTag(ULONG ulIndex);

  /*--------------------------------------------------------------------
  getPropValue returns PropValue of given tag.
    returns: PropValue of given tag or NULL, if an error occurred.
  heap size not invariant: caller must delete returned object.
  --------------------------------------------------------------------*/
  PropValue* getPropValue(ULONG ulTag);

  /* Stream* getStream(ULONG ulTag); */
  /* Message* getMessage(ULONG ulTag); */
  /* StreamDocFile* getDocFile(ULONG ulTag); */

  /*--------------------------------------------------------------------
  list lists all properties.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT list();

protected:
  /* pointer to wrapped IMAPIProp interface */
  IMAPIProp* m_pMapiProp;

private:
  /* list of PropValues */
  LPSPropTagArray m_pPropTagArray;

}; /* class Properties */

#endif // PROPERTIES_H not defined
/*==End of File=======================================================*/
