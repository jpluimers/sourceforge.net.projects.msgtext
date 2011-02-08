/*-- $Workfile: Heap.cpp $ ---------------------------------------------
Object for heap diagnostics.
Version    : $Id$
Application: Windows Utilities
Platform   : Windows, Extended MAPI
Description: Heap diagnostics.
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
#ifndef HEAP_H
#define HEAP_H
#include <malloc.h>
#include "Heap.h"

/*--------------------------------------------------------------------*/
Heap::Heap()
{
  m_sizeInitialUsed = this->bytesUsed();
} /* Heap */

/*--------------------------------------------------------------------*/
Heap::~Heap()
{
} /* ~Heap */

/*--------------------------------------------------------------------*/
int Heap::isValid()
{
  return (_heapchk() == _HEAPOK);
} /* isValid */

/*--------------------------------------------------------------------*/
int Heap::hasInitialSize()
{
  return (this->bytesUsed() == m_sizeInitialUsed);
} /* hasInitialSize */

/*--------------------------------------------------------------------*/
int Heap::minimize()
{
  return _heapmin();
} /* minimize */

/*--------------------------------------------------------------------*/
size_t Heap::bytesUsed()
{
  /* uses _heapwalk() in run time library */
  _HEAPINFO hinfo;
  int iHeapStatus;
  size_t sizeUsed = 0xFFFFFFFF;
  if (minimize() == 0)
  {
    hinfo._pentry = 0;
    sizeUsed = 0;
    for (iHeapStatus = _heapwalk(&hinfo); iHeapStatus == _HEAPOK; iHeapStatus = _heapwalk(&hinfo))
    {
      if (hinfo._useflag == _USEDENTRY)
        sizeUsed += hinfo._size;
    }
    if ((iHeapStatus != _HEAPEND) && (iHeapStatus != _HEAPEMPTY))
      sizeUsed = 0xFFFFFFFF;
  }
  return sizeUsed;
} /* bytesUsed */

#endif // HEAP_H not defined
/*==End of File=======================================================*/
