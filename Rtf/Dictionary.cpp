/*-- $Workfile: Dictionary.cpp $ ---------------------------------------
Dictionary for decompressing compressed RTF Streams.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Dictionary class for decompressing compressed RTF Streams.
             See MS-OXRTFCP.pdf (http://msdn.microsoft.com/en-us/library/cc463890.aspx)
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
#define STRICT
#include <windows.h>
#include "../Utils/Heap.h"
#include "Dictionary.h"

#define BUFFER_SIZE 4096

/*--------------------------------------------------------------------*/
int Dictionary::inc(int i)
{
  i++;
  if (i == BUFFER_SIZE)
    i = 0;
  return i;
} /* inc */

/*--------------------------------------------------------------------*/
void Dictionary::initialize()
{
  m_pDictionary = (char*)malloc(BUFFER_SIZE);
  strcpy(m_pDictionary,
    "{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}");
  strcat(m_pDictionary,
    "{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New RomanCourier");
  strcat(m_pDictionary,
    "{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx");
  m_iWritePos = strlen(m_pDictionary); /* must be 207 */
} /* initialize */

/*--------------------------------------------------------------------*/
Dictionary::Dictionary()
{
  initialize();
} /* Dictionary */

/*--------------------------------------------------------------------*/
Dictionary::~Dictionary()
{
  if (m_pDictionary != NULL)
    free(m_pDictionary);
} /* ~Dictionary */

/*--------------------------------------------------------------------*/
BOOL Dictionary::put(char c)
{
  m_pDictionary[m_iWritePos] = c;
  m_iWritePos = inc(m_iWritePos);
  return true;
} /* put */

/*--------------------------------------------------------------------*/
BOOL Dictionary::get(int iOffset, int iLength, char* pBuffer)
{
  BOOL bSucceeded = false;
  if (iOffset != m_iWritePos)
  {
    bSucceeded = true;
    for (int i = 0; bSucceeded && (i < iLength); i++)
    {
      pBuffer[i] = m_pDictionary[iOffset];
      iOffset = inc(iOffset);
      put(pBuffer[i]);
    }
  }
  return bSucceeded;
} /* get */

/*==End of File=======================================================*/
