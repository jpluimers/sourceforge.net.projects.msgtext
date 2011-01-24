/*-- $Workfile: Row.h $ ------------------------------------------------
Wrapper for an SRow.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Row class wraps an SRow.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#ifndef ROW_H
#define ROW_H

#define STRICT
#include <mapix.h>
#include "PropValue.h"

/*======================================================================
Row is the SRow wrapper
======================================================================*/
class Row
{
public:
  /*--------------------------------------------------------------------
  constructor 
    pSRow: pointer to SRow struct on the heap.
  heap size invariant.
  --------------------------------------------------------------------*/
  Row(LPSRow pSRow);

  /*--------------------------------------------------------------------
  destructor 
    deletes the pointer to the SRow struct.
  heap size not invariant: m_pSRow is freed.
  --------------------------------------------------------------------*/
  ~Row();

  /*--------------------------------------------------------------------
  columns
    returns: number of PropValue "columns" of this row.
  heap size invariant.
  --------------------------------------------------------------------*/
  ULONG columns();

  /*--------------------------------------------------------------------
  getColumn return PropValue "column" of given index.
    ulIndex: index of PropValue "column" to return.
    returns: the PropValue "column" of the given index or NULL, if an error occurred.
  heap size not invariant: caller must delete returned object.
  --------------------------------------------------------------------*/
  PropValue* getColumn(ULONG ulTag);

  /*--------------------------------------------------------------------
  list lists all PropValue "columns" to stdout.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT list();

private:
  /* pointer to SRow struct */
  LPSRow m_pSRow;
}; /* class Row */

#endif // ROW_H not defined
/*==End of File=======================================================*/
