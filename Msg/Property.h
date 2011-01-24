/*-- $Workfile: Property.h $ -------------------------------------------
Property represents a property in a Storage object.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Property class represents a property in a Storage object.
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
#ifndef PROPERTY_H
#define PROPERTY_H
#define STRICT
#include <mapix.h>
#include "../Storage/Storage.h"

#define PR_BODY_HTML					PROP_TAG( PT_LONG,		0x1013)

/*======================================================================
Property class represents a property in a Storage object.
======================================================================*/
class Property
{
public:
  /*--------------------------------------------------------------------
  constructor 
    pStorage: storage where variable length property values are stored.
    ulTag: tag value of property.
    pValue: pointer to 8 byte value array of property.
  heap size invariant.
  --------------------------------------------------------------------*/
  Property(Storage* pStorage, ULONG ulTag, void* pValue);

  /*--------------------------------------------------------------------
  destructor 
  heap size invariant.
  --------------------------------------------------------------------*/
  ~Property();

  /*--------------------------------------------------------------------
  tag
    returns: tag of this Property.
  heap size invariant.
  --------------------------------------------------------------------*/
  ULONG tag();

  /*--------------------------------------------------------------------
  type
    returns: a display string describing the type of this Property.
  heap size invariant: returned string points to a "constant" in the
                       data segment.
  --------------------------------------------------------------------*/
  char* type();

  /*--------------------------------------------------------------------
  id
    returns: a display string describing the id of this Property.
  heap size invariant: returned string points to a "constant" in the
                       data segment.
  --------------------------------------------------------------------*/
  char* id();

  /*--------------------------------------------------------------------
  value
    returns: a display string describing the value of this Property.
  heap size invariant: returned string points to a "static" array in the
                       data segment. N.B.: Not MT-safe! Will be overwritten
                       on next call!
  --------------------------------------------------------------------*/
  char* value();

  /*--------------------------------------------------------------------
  variable
    returns: true, if type is 
      PT_OBJECT
      PT_STRING8
      PT_UNICODE
      PT_CLSID
      PT_BINARY
   or a multi-valued variant of one of these.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL variable();

  /*--------------------------------------------------------------------
  multiValued
    returns: true, if multi-valued flag is set.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL multiValued();

  /*--------------------------------------------------------------------
  size
    returns: length of byte array of Propertys of variable length Property.
  heap size invariant.
  --------------------------------------------------------------------*/
  ULONG size();

  /*--------------------------------------------------------------------
  getBooleanValue
    returns: the value of the PT_BOOLEAN Property.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL getBooleanValue();

  /*--------------------------------------------------------------------
  getShortValue
    returns: the value of the PT_SHORT (i.e. PT_I2) Property.
  heap size invariant.
  --------------------------------------------------------------------*/
  short getShortValue();

  /*--------------------------------------------------------------------
  getLongValue
    returns: the value of the PT_LONG (i.e. PT_I4) Property.
  heap size invariant.
  --------------------------------------------------------------------*/
  ULONG getLongValue();

  /*--------------------------------------------------------------------
  getLargeValue
    returns: the value of the PT_LONGLONG (i.e. PT_I8) Property.
  heap size invariant.
  --------------------------------------------------------------------*/
  LARGE_INTEGER getLargeValue();

  /*--------------------------------------------------------------------
  getFloatValue
    returns: the value of the PT_FLOAT (i.e. PT_R4) Property.
  heap size invariant.
  --------------------------------------------------------------------*/
  float getFloatValue();

  /*--------------------------------------------------------------------
  getDoubleValue
    returns: the value of the PT_Double (i.e. PT_R8) Property.
  heap size invariant.
  --------------------------------------------------------------------*/
  double getDoubleValue();

  /*--------------------------------------------------------------------
  getSysTimeValue
    returns: pointer to value of Property of type PT_SYSTIME.
  heap size not invariant: Caller must free returned value.
  --------------------------------------------------------------------*/
  FILETIME getSysTimeValue();

  /*--------------------------------------------------------------------
  getString8Value
    returns: the value of the PT_STRING8 Property.
  heap size not invariant: caller must delete returned value.
  --------------------------------------------------------------------*/
  char* getString8Value();

  /*--------------------------------------------------------------------
  getUnicodeValue
    returns: the value of the PT_UNICODE Property.
  heap size not invariant: caller must delete returned value.
  --------------------------------------------------------------------*/
  LPOLESTR getUnicodeValue();

  /*--------------------------------------------------------------------
  getBinaryIStream
    returns: pointer to an IStream interface containing the binary value,
             or NULL.
  heap size not invariant: caller must release returned value.
  --------------------------------------------------------------------*/
  IStream* getBinaryIStream();

  /*--------------------------------------------------------------------
  getBinaryStream
    returns: pointer to Stream object containing the binary value,
             or NULL.
  heap size not invariant: caller must delete returned value.
  --------------------------------------------------------------------*/
  Stream* getBinaryStream();

  /*--------------------------------------------------------------------
  getBinaryValue
    returns: pointer to byte array of Propertys of type PT_BINARY.
  heap size not invariant: caller must free returned value.
  --------------------------------------------------------------------*/
  LPBYTE getBinaryValue();

  /*--------------------------------------------------------------------
  getBinaryIStorage
    returns: pointer to an IStorage object containing the binary value,
             or NULL.
  heap size not invariant: caller must release returned value.
  --------------------------------------------------------------------*/
  IStorage* getBinaryIStorage();

  /*--------------------------------------------------------------------
  list displays id, type and value
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT list();

/*====================================================================*/
private:
  /*--------------------------------------------------------------------
  getSubstgName
    returns: "__substg1.0_HHHHHHHH".
  heap size not invariant: Caller must free resulting name.
  --------------------------------------------------------------------*/
  LPOLESTR getSubstgName();

  /* tag of this property */
  ULONG m_ulTag;
  /* value of this property */
  char m_acValue[8];
  /* Storage object which contains the variable properties Streams 
     with names "__substg1.0_HHHHHHHH", where HHHHHHHH is the hex string
     of the tag */
  Storage* m_pStorage; 
}; /* class PropValue */

#endif // PROPVALUE_H not defined
/*==End of File=======================================================*/
