/*-- $Workfile: Row.cpp $ ----------------------------------------------
Wrapper for an SRow
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Row class wraps an SRow.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#include <stdio.h>
#define STRICT
#include <mapix.h>
#include "../Utils/Heap.h"
#include "Row.h"
#include "PropValue.h"

/*--------------------------------------------------------------------*/
Row::Row(LPSRow pSRow)
{
  m_pSRow = pSRow;
} /* Row */

/*--------------------------------------------------------------------*/
Row::~Row()
{
  MAPIFreeBuffer(m_pSRow);
} /* ~Row */

/*--------------------------------------------------------------------*/
ULONG Row::columns()
{
  return m_pSRow->cValues;
} /* columns */

/*--------------------------------------------------------------------*/
PropValue* Row::getColumn(ULONG ulTag)
{
  PropValue* pPv = NULL;
  for (ULONG ulColumn = 0; (ulColumn < this->columns()) && ((pPv == NULL) || (pPv->tag() != ulTag)); ulColumn++)
    pPv = new PropValue(&m_pSRow->lpProps[ulColumn]);
  if ((pPv != NULL) && (pPv->tag() != ulTag))
    pPv = NULL;
  return pPv;
} /* getColumn */

/*--------------------------------------------------------------------*/
HRESULT Row::list()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = NOERROR;
  PropValue* pPv;
  printf("row list\n");
  for (ULONG ulColumn = 0; ulColumn < this->columns(); ulColumn++)
  {
    pPv = new PropValue(&m_pSRow->lpProps[ulColumn]);
    pPv->list();
  }
  delete pPv;
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
    hr = E_FAIL;
  delete pHeap;
#endif
  return hr;
} /* list */

/*==End of File=======================================================*/
