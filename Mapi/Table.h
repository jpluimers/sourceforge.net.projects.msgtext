/*-- $Workfile: Table.h $ ----------------------------------------------
Wrapper for an IMAPITable.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Table class wraps an IMAPITable.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#ifndef TABLE_H
#define TABLE_H

#define STRICT
#include <mapix.h>
#include "Row.h"

/*======================================================================
Table is the IMAPITable wrapper
======================================================================*/
class Table
{
public:
  /*--------------------------------------------------------------------
  constructor 
    pMapiTable: pointer to IMAPITable interface.
  heap size not invariant: establishes list of entry ids.
  --------------------------------------------------------------------*/
  Table(LPMAPITABLE pMapiTable);

  /*--------------------------------------------------------------------
  destructor releases IMAPITable interface.
  heap size not invariant: cleans up list of entry ids.
  --------------------------------------------------------------------*/
  ~Table();

  /*--------------------------------------------------------------------
  entries
    returns: number of rows of this MAPI table.
  heap size invariant.
  --------------------------------------------------------------------*/
  ULONG entries();

  /*--------------------------------------------------------------------
  getEntryIdLength returns length of entry id of given index.
    ulIndex: index of row for which length of entry id is to be returned.
    returns: length of entry id of given index.
  heap size invariant.
  --------------------------------------------------------------------*/
  ULONG getEntryIdLength(ULONG ulIndex);

  /*--------------------------------------------------------------------
  getEntryId returns pointer to entry id of given index.
    ulIndex: index of row for which entry id is to be returned.
    returns: pointer to entry id of given index.
  heap size invariant.
  --------------------------------------------------------------------*/
  LPENTRYID getEntryId(ULONG ulIndex);

private:
  /* pointer to IMAPITable interface */
  LPMAPITABLE m_pMapiTable;
  /* pointer to array of entry ids (each has a 4 byte length prefix) */
  BYTE** m_apEntryId; /* array of pointers to length/entryid components */

}; /* class Table */

#endif // TABLE_H not defined
/*==End of File=======================================================*/
