/*-- $Workfile: FileTime.cpp $ -----------------------------------------
Wrapper for a FILETIME
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: FileTime class wraps aFILETIME.
             N.B. (http://msdn.microsoft.com/en-us/library/ms724284(VS.85).aspx)
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
#define _CRT_SECURE_NO_WARNINGS
#include <stdio.h>
#include <time.h>
#define STRICT
#include <windows.h>
#include "FileTime.h"

/*--------------------------------------------------------------------*/
FileTime::FileTime(FILETIME* pFileTime) 
{
  memcpy(&m_ft,pFileTime,sizeof(m_ft));
} /* FileTime */

/*--------------------------------------------------------------------*/
FileTime::~FileTime()
{
} /* ~FileTime */

/*--------------------------------------------------------------------*/
FILETIME* FileTime::getFt()
{
  return &m_ft;
} /* getFt */

/*--------------------------------------------------------------------*/
char* FileTime::getInternetTime()
{
  int iSIZE_INTERNET_TIME = 64;
  char* pInternetTime = (char*)malloc(iSIZE_INTERNET_TIME);
  FILETIME ftLocal;
  SYSTEMTIME stUtc;
  struct tm tmUtc;
  FileTimeToLocalFileTime(&m_ft,&ftLocal);
  /* FILETIME resolution is 100 nano seconds */
  LARGE_INTEGER li;
  LARGE_INTEGER liLocal;
  li.HighPart = m_ft.dwHighDateTime;
  li.LowPart = m_ft.dwLowDateTime;
  liLocal.HighPart = ftLocal.dwHighDateTime;
  liLocal.LowPart = ftLocal.dwLowDateTime;
  li.QuadPart = liLocal.QuadPart - li.QuadPart;
  /* this is the locale UTC offset in nano seconds */
  li.QuadPart = li.QuadPart/10000000; /* seconds */
  li.QuadPart = li.QuadPart/60; /* minutes */
  char cSign = '+';
  if (li.QuadPart < 0)
  {
    cSign = '-';
    li.QuadPart = -li.QuadPart;
  }
  int m = (int)li.QuadPart % 60; /* minutes */
  int h = (int)li.QuadPart / 60; /* hours */
  char* pTimezone = (char*)malloc(7);
  sprintf(pTimezone,"%c%02d:%02d",cSign,h,m);
  if (FileTimeToSystemTime(&m_ft,&stUtc))
  {
    tmUtc.tm_year = stUtc.wYear - 1900;
    tmUtc.tm_mon = stUtc.wMonth - 1;
    tmUtc.tm_mday = stUtc.wDay;
    tmUtc.tm_hour = stUtc.wHour;
    tmUtc.tm_min = stUtc.wMinute;
    tmUtc.tm_sec = stUtc.wSecond;
    tmUtc.tm_wday = stUtc.wDayOfWeek;
    strftime(pInternetTime,iSIZE_INTERNET_TIME,"%a, %#d %b %Y %H:%M:%S",&tmUtc);
    strcat(pInternetTime,pTimezone);
    free(pTimezone);
  }
  return pInternetTime;
} /* getInternetTime */

/*--------------------------------------------------------------------*/
BOOL FileTime::list()
{
  BOOL bListed = false;
  SYSTEMTIME st;
  bListed = FileTimeToSystemTime(&m_ft,&st);
  if (bListed)
  {
    printf("%02hu.%02hu.%04hu %02hu:%02hu:%02hu:%03hu\n",
           st.wDay,st.wMonth,st.wYear,st.wHour,st.wMinute,st.wMilliseconds);
  }
  return bListed;
} /* list */

/*==End of File=======================================================*/
