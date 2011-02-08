/*-- $Workfile: FileTime.h $ -------------------------------------------
Wrapper for a FILETIME.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Container class wraps a FILETIME.
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
#ifndef FILETIME_H
#define FILETIME_H

#define STRICT
#include <windows.h>

/*======================================================================
FileTime is the FILETIME wrapper
======================================================================*/
class FileTime
{
public:
  /*--------------------------------------------------------------------
  constructor 
    pFileTime: pointer to FILETIME which is to be copied.
  heap size invariant.
  --------------------------------------------------------------------*/
  FileTime(FILETIME* pFileTime);

  /*--------------------------------------------------------------------
  destructor
  heap size invariant.
  --------------------------------------------------------------------*/
  ~FileTime();

  /*--------------------------------------------------------------------
  getFt returns the pointer to the FILETIME struct which will remain
        valid while this object is not deleted.
    returns: returns the pointer to the FILETIME struct.
  heap size invariant.
  --------------------------------------------------------------------*/
  FILETIME* getFt();

  /*--------------------------------------------------------------------
  convert FILETIME to RFC2882 Internet Date format.
    returns: returns the pointer to Internet Date.
  heap size not invariant: Caller must free returned string.
  --------------------------------------------------------------------*/
  char* getInternetTime();

  /*--------------------------------------------------------------------
  list displays the time in the format dd.MM.yyyy hh:mm:ss.fff
    returns: true, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL list();

private:
  /* pointer to FILETIME struct */
  FILETIME m_ft;
}; /* class FileTime */

#endif // FILETIME_H not defined
/*==End of File=======================================================*/
