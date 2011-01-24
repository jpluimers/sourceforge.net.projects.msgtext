/*-- $Workfile: PropStorage.cpp $ --------------------------------------
Class representing a Msg object with Properties in a Storage file.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Class representing a Msg object (Msg, Attachment, Recipient)
             with Properties in a Storage file.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
This file is part of MsgText.
MsgText is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
any later version.
MsgText is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
You should have received a copy of the GNU General Public License
along with MsgText.  If not, see <http://www.gnu.org/licenses/>.
----------------------------------------------------------------------*/
#include <stdio.h>
#define STRICT
#include <windows.h>
#include "../Utils/Fs.h"
#include "../Utils/Heap.h"
#include "PropStorage.h"

/*--------------------------------------------------------------------*/
LPBYTE PropStorage::header(LPBYTE pHeader)
{
  return NULL; // must be overdefined!
} /* header */

/*--------------------------------------------------------------------*/
void PropStorage::clear()
{
  m_bUnicode = false;
  m_apProperty = NULL;
  m_iPropertyCount = 0;
} /* clear */

/*--------------------------------------------------------------------*/
HRESULT PropStorage::initialize()
{
  int iDATA_SIZE = 16;
  int iPROPATTR_MANDATORY = 0x00000001;
  int iPROPATTR_READABLE = 0x00000002;
  int iPROPATTR_WRITABLE = 0x00000004;
  LPOLESTR awcPROPERTIES = L"__properties_version1.0";
  int iSize;
  Stream* pPropertyStream = NULL;
  LPBYTE pProperties = NULL;
  LPBYTE pProp = NULL;
  Property* pProperty;
  ULONG ulTag;
  ULONG ulType;
  HRESULT hr = Storage::initialize();
  if (SUCCEEDED(hr))
  {
    pPropertyStream = getStream(awcPROPERTIES);
    if (pPropertyStream != NULL)
    {
      iSize = pPropertyStream->size();
      pProperties = pPropertyStream->content();
      if (pProperties != NULL)
      {
        pProp = header((LPBYTE)pProperties);
        for (iSize -= (pProp - pProperties); iSize >= iDATA_SIZE; iSize -= iDATA_SIZE)
        {
          memcpy(&ulTag,pProp,sizeof(ulTag));
          pProp += sizeof(ulTag);
          memcpy(&ulType,pProp,sizeof(ulType));
          pProp += sizeof(ulType);
          if ((ulType & iPROPATTR_READABLE) == iPROPATTR_READABLE)
          {
            pProperty = new Property(this,ulTag,pProp);
            m_iPropertyCount = append((void***)&m_apProperty,m_iPropertyCount,pProperty);
          }
          pProp += iDATA_SIZE - sizeof(ulTag) - sizeof(ulType);
        }
        free(pProperties);
      }
      delete pPropertyStream;
#ifdef _DEBUG
      printf("Properties contained ----------------------\n");
      listProperties();
#endif
    }
    else
      printf("Stream %S could not be opened!",awcPROPERTIES);
  }
  return hr;
} /* initialize */

/*--------------------------------------------------------------------*/
PropStorage::PropStorage(char* pName): Storage(pName)
{
  clear();
} /* PropStorage */

/*--------------------------------------------------------------------*/
PropStorage::PropStorage(IStorage* pStorage, BOOL bUnicode) 
  : Storage(pStorage)
{
  clear();
  m_bUnicode = bUnicode;
} /* PropStorage */

/*--------------------------------------------------------------------*/
PropStorage::~PropStorage()
{
  for (int iProperty = 0; iProperty < m_iPropertyCount; iProperty++)
  {
    Property* pProperty = m_apProperty[iProperty];
    delete(pProperty);
  }
  free(m_apProperty);
} /* ~PropStorage */

/*--------------------------------------------------------------------*/
int PropStorage::getPropertyCount()
{
  return m_iPropertyCount;
} /* getPropertyCount */

/*--------------------------------------------------------------------*/
Property* PropStorage::getProperty(ULONG ulTag)
{
  Property* pProperty = NULL;
  int iProperty;
  for (iProperty = 0; (pProperty == NULL) && (iProperty < m_iPropertyCount); iProperty++)
  {
    if (m_apProperty[iProperty]->tag() == ulTag)
      pProperty = m_apProperty[iProperty];
  }
  return pProperty;
} /* getProperty */

/*--------------------------------------------------------------------*/
char* PropStorage::getStringValue(ULONG ulTag)
{
  char* pValue = NULL;
  LPOLESTR pwcsValue = NULL;
  Property* pProperty = NULL;
  if (m_bUnicode)
  {
    pProperty = getProperty((ulTag & 0xFFFF0000) | PT_UNICODE);
    if (pProperty != NULL)
    {
      pwcsValue = pProperty->getUnicodeValue();
      if (pwcsValue != NULL)
      {
        pValue = newMultiByte(pwcsValue);
        free(pwcsValue);
      }
    }
  }
  else
  {
    pProperty = getProperty((ulTag & 0xFFFF0000) | PT_STRING8);
    if (pProperty != NULL)
      pValue = pProperty->getString8Value();
  }
  return pValue;
} /* getStringValue */

/*--------------------------------------------------------------------*/
FileTime* PropStorage::getSysTimeValue(ULONG ulTag)
{
  FileTime* pTime = NULL;
  Property* pProperty = getProperty(ulTag);
  if (pProperty != NULL)
    pTime = new FileTime(&pProperty->getSysTimeValue());
  return pTime;
} /* getSysTimeValue */

/*--------------------------------------------------------------------*/
void PropStorage::listProperties()
{
  HRESULT hr = S_OK;
  int iProperty;
  for (iProperty = 0; SUCCEEDED(hr) && (iProperty < m_iPropertyCount); iProperty++)
    hr = m_apProperty[iProperty]->list();
} /* listProperties */

/*==End of File=======================================================*/
