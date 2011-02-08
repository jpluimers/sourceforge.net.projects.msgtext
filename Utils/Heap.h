/*-- $Workfile: Heap.h $ -----------------------------------------------
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
#ifndef MYHEAP_H
#define MYHEAP_H

/*======================================================================
Heap wraps the heap
======================================================================*/
class Heap
{
public:
  /*--------------------------------------------------------------------
  constructor records initial used bytes
  heap size invariant.
  --------------------------------------------------------------------*/
  Heap();

  /*--------------------------------------------------------------------
  destructor
  heap size invariant.
  --------------------------------------------------------------------*/
  ~Heap();

  /*--------------------------------------------------------------------
  isValid 
    returns: false (0) if heap is inconsistent.
  heap size invariant.
  --------------------------------------------------------------------*/
  int isValid();

  /*--------------------------------------------------------------------
  hasInitialSize
    returns: false (0) if the heap does not have the same size as on 
             this object's creation.
  heap size invariant.
  --------------------------------------------------------------------*/
  int hasInitialSize();

  /*--------------------------------------------------------------------
  minimize minimizes the heap
    returns: returns 0 if heap is consistent.
  heap size invariant.
  --------------------------------------------------------------------*/
  int minimize();

  /*--------------------------------------------------------------------
  bytesUsed 
    returns: number of used bytes or 0xFFFFFFFF if heap is inconsistent.
  heap size invariant.
  --------------------------------------------------------------------*/
  size_t bytesUsed();

private:
  /* stores the number of used bytes at this object's creation */
  size_t m_sizeInitialUsed;
}; /* class Heap */

#endif // MYHEAP_H not defined
/*==End of File=======================================================*/
