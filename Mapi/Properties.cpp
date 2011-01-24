/*-- $Workfile: Properties.cpp $ ---------------------------------------
Wrapper for an IMAPIProp.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Properties class wraps IMAPIProp.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#include <stdio.h>
#define STRICT
#include <mapix.h>
#include "Properties.h"

/*--------------------------------------------------------------------*/
Properties::Properties(IMAPIProp* pMapiProp)
{
  m_pPropTagArray = NULL;
  m_pMapiProp = pMapiProp;
  if (pMapiProp != NULL)
    m_pMapiProp->GetPropList(0,&m_pPropTagArray);
} /* Properties */

/*--------------------------------------------------------------------*/
Properties::~Properties()
{
  if (m_pPropTagArray != NULL)
  {
    /* release all object properties!?! */
    MAPIFreeBuffer(m_pPropTagArray);
  }
  if (m_pMapiProp != NULL)
    m_pMapiProp->Release();
} /* ~Properties */

/*--------------------------------------------------------------------*/
ULONG Properties::getCount()
{
  ULONG ulCount = 0xFFFFFFFF;
  if (m_pPropTagArray != NULL)
    ulCount = m_pPropTagArray->cValues;
  return ulCount;
} /* getCount */

/*--------------------------------------------------------------------*/
ULONG Properties::getTag(ULONG ulIndex)
{
  ULONG ulTag = 0xFFFFFFFF;
  if ((m_pPropTagArray != NULL) && (ulIndex < this->getCount()))
    ulTag = m_pPropTagArray->aulPropTag[ulIndex];
  return ulTag;
} /* getTag */

/*--------------------------------------------------------------------*/
PropValue* Properties::getPropValue(ULONG ulTag)
{
  PropValue* pPv = NULL;
  SizedSPropTagArray(1, spt) = { 1, { 0 }};
  ULONG ulValues = spt.cValues;
  spt.aulPropTag[0] = ulTag;
  LPSPropValue pSPropValue = NULL;
  HRESULT hr = m_pMapiProp->GetProps((LPSPropTagArray)&spt,0,&ulValues,&pSPropValue);
  if ((!FAILED(hr)) && (hr != MAPI_W_ERRORS_RETURNED) && (ulValues == 1))
    pPv = new PropValue(pSPropValue);
  else if (hr == MAPI_W_ERRORS_RETURNED) // property not accessible
    MAPIFreeBuffer(pSPropValue);
  return pPv;
} /* getPropValue */

/*--------------------------------------------------------------------*/
HRESULT Properties::list()
{
  ULONG ulIndex;
  ULONG ulTag;
  PropValue* pPv;
  HRESULT hr = NOERROR;
  for (ulIndex = 0; (ulIndex < this->getCount()) && (!FAILED(hr)); ulIndex++)
  {
    ulTag = this->getTag(ulIndex);
    pPv = this->getPropValue(ulTag);
    if (pPv != NULL)
    {
      hr = pPv->list();
      delete pPv;
    }
  }
  return hr;
} /* list */

/*==End of File=======================================================*/
