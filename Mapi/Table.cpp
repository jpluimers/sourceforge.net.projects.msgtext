/*-- $Workfile: Table.cpp $ --------------------------------------------
Wrapper for an IMAPITable.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Table class wraps an IMAPITable.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#include <stdio.h>
#define STRICT
#include <mapix.h>
#include <mapiutil.h>
#include "Table.h"
#include "Row.h"
#include "PropValue.h"

/*--------------------------------------------------------------------*/
Table::Table(LPMAPITABLE pMapiTable)
{
  SizedSPropTagArray(1,sptEntry) = { 1, {PR_ENTRYID} };
  LPSRowSet pSRowSet;
  Row* pRow;
  PropValue* pPv;
  m_pMapiTable = pMapiTable;
  m_apEntryId = NULL;
  ULONG ulLength;
  ULONG ulEntries = this->entries();
  HRESULT hr = HrQueryAllRows(m_pMapiTable,(LPSPropTagArray)&sptEntry,NULL,NULL,0,&pSRowSet);
  if (!FAILED(hr) && (pSRowSet->cRows == ulEntries))
  {
    m_apEntryId = (BYTE**)malloc(ulEntries*sizeof(BYTE*));
    for (ULONG ulEntry = 0; ulEntry < ulEntries; ulEntry++)
    {
      pRow = new Row(&pSRowSet->aRow[ulEntry]);
      pPv = pRow->getColumn(PR_ENTRYID);
      LPENTRYID pEntryId = (LPENTRYID)pPv->getBinaryPointer();
      ulLength = pPv->getBinaryLength();
      m_apEntryId[ulEntry] = (BYTE*)malloc(sizeof(ULONG)+ulLength);
      memcpy(m_apEntryId[ulEntry],&ulLength,sizeof(ULONG));
      memcpy(&m_apEntryId[ulEntry][sizeof(ULONG)],pEntryId,ulLength);
      delete pPv;
      delete pRow;
    }
  }
} /* Table */

/*--------------------------------------------------------------------*/
Table::~Table()
{
  if (m_apEntryId != NULL)
  {
    for (ULONG ulEntry = 0; ulEntry < this->entries(); ulEntry++)
    {
      if (m_apEntryId[ulEntry] != NULL)
        free(m_apEntryId[ulEntry]);
    }
    free(m_apEntryId);
  }
  if (m_pMapiTable != NULL)
    m_pMapiTable->Release();
} /* ~Table */

/*--------------------------------------------------------------------*/
ULONG Table::entries()
{
  ULONG ulCount = 0xFFFFFFFF;
  if (m_pMapiTable != NULL)
  {
    HRESULT hr = m_pMapiTable->GetRowCount(0,&ulCount);
    if (FAILED(hr))
      ulCount = 0xFFFFFFFF;
  }
  return ulCount;
} /* entries */

/*--------------------------------------------------------------------*/
LPENTRYID Table::getEntryId(ULONG ulIndex)
{
  LPENTRYID pEntryId = NULL;
  if ((m_apEntryId != NULL) && (ulIndex < this->entries()))
    pEntryId = (LPENTRYID)&m_apEntryId[ulIndex][sizeof(ULONG)];
  return pEntryId;
} /* getEntry */

/*--------------------------------------------------------------------*/
ULONG Table::getEntryIdLength(ULONG ulIndex)
{
  ULONG ulLength = NULL;
  if ((m_apEntryId != NULL) && (ulIndex < this->entries()))
    memcpy(&ulLength,m_apEntryId[ulIndex],sizeof(ULONG));
  return ulLength;
} /* getEntryLength */

/*==End of File=======================================================*/
