/*-- $Workfile: PropValue.h $ ------------------------------------------
Wrapper for an SPropValue.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: PropValue class wraps an SPropValue.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#ifndef PROPVALUE_H
#define PROPVALUE_H

#define STRICT
#include <mapix.h>

/*======================================================================
PropValue is the SPropValue wrapper
======================================================================*/
class PropValue
{
public:
  /*--------------------------------------------------------------------
  constructor 
    pSPropValue: pointer to SPropValue struct on the heap.
  heap size invariant.
  --------------------------------------------------------------------*/
  PropValue(LPSPropValue pSPropValue);

  /*--------------------------------------------------------------------
  destructor 
    deletes the pointer to the SPropValue struct.
  heap size not invariant: m_pSPropValue is freed.
  --------------------------------------------------------------------*/
  ~PropValue();

  /*--------------------------------------------------------------------
  tag
    returns: tag of this PropValue.
  heap size invariant.
  --------------------------------------------------------------------*/
  ULONG tag();

  /*--------------------------------------------------------------------
  type
    returns: a display string describing the type of this PropValue.
  heap size invariant: returned string points to a "constant" in the
                       data segment.
  --------------------------------------------------------------------*/
  char* type();

  /*--------------------------------------------------------------------
  id
    returns: a display string describing the id of this PropValue.
  heap size invariant: returned string points to a "constant" in the
                       data segment.
  --------------------------------------------------------------------*/
  char* id();

  /*--------------------------------------------------------------------
  id
    returns: a display string describing the value of this PropValue.
  heap size invariant: returned string points to a "static" array in the
                       data segment. N.B.: Not MT-safe! Will be overwritten
                       on next call!
  --------------------------------------------------------------------*/
  char* value();

  /*--------------------------------------------------------------------
  getBooleanValue
    returns: the value of the PT_BOOLEAN PropValue.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL getBooleanValue();

  /*--------------------------------------------------------------------
  getLongValue
    returns: the value of the PT_LONG (i.e. PT_I4) PropValue.
  heap size invariant.
  --------------------------------------------------------------------*/
  ULONG getLongValue();

  /*--------------------------------------------------------------------
  getString8Value
    returns: the value of the PT_STRING8 PropValue.
  heap size not invariant: caller must delete returned value.
  --------------------------------------------------------------------*/
  char* getString8Value();

  /*--------------------------------------------------------------------
  getBinaryPointer
    returns: pointer to byte array of PropValues of type PT_BINARY.
  heap size invariant.
  --------------------------------------------------------------------*/
  LPBYTE getBinaryPointer();

  /*--------------------------------------------------------------------
  getBinaryLength
    returns: length of byte array of PropValues of type PT_BINARY.
  heap size invariant.
  --------------------------------------------------------------------*/
  ULONG getBinaryLength();

  /*--------------------------------------------------------------------
  getSysTimeValue
    returns: pointer to value of PropValues of type PT_SYSTIME.
  heap size invariant.
  --------------------------------------------------------------------*/
  FILETIME* getSysTimeValue();

  /*--------------------------------------------------------------------
  list displays id, type and value
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT list();
private:
  LPSPropValue m_pSPropValue;
}; /* class PropValue */

#endif // PROPVALUE_H not defined
/*==End of File=======================================================*/
