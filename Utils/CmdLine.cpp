/*-- $Workfile: CmdLine.cpp $ ------------------------------------------
Command-line parsing
Version    : $Id$
Application: Utilities
Platform   : Windows
Description: This object parses the command-line into named and unnamed
             arguments. Named arguments start with "/" or "-" followed
             by the name of the argument, either followed by a colon or 
             an equal, and the value of the argument or followed by the 
             value in the next argument.
             Unnamed arguments are all the other arguments on the command
             line. (The argument 0 is skipped!)
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
#include "CmdLine.h"

/*--------------------------------------------------------------------*/
CmdLine::CmdLine(int argc, char** args) 
{
#ifdef _DEBUG
  for (int i = 0; i < argc; i++)
    printf("arg[%d] : %s\n",i,args[i]);
#endif
  m_pError = NULL;
  m_uArguments = argc-1;
  m_psName = NULL;
  m_psValue = NULL;
  if (m_uArguments > 0)
  {
    m_psName = (char**)malloc(m_uArguments*sizeof(char*));
    m_psValue = (char**)malloc(m_uArguments*sizeof(char*));
    unsigned uArg;
    unsigned uPos;
    char c;
    /* skip argument 0, which under Windows is the executable */
    for (uArg = 0; (m_pError == NULL) && (uArg < m_uArguments); uArg++)
    {
      m_psName[uArg] = NULL;
      m_psValue[uArg] = NULL;
      if ((*args[uArg+1] == '-') || (*args[uArg+1] == '/'))
      {
        /* named argument */
        for (uPos = 1; (isalpha(args[uArg+1][uPos])) || (isdigit(args[uArg+1][uPos])) && (uPos < strlen(args[uArg+1])); uPos++) {}
        if (uPos > 1)
        {
          c = args[uArg+1][uPos];
          args[uArg+1][uPos] = '\0';
          if ((c == ':') || (c == '=') || (c == '\0'))
          {
            m_psName[uArg] = _strdup(&args[uArg+1][1]);
            if (c != '\0')
              m_psValue[uArg] = _strdup(&args[uArg+1][uPos+1]);
          }
          else
          {
            char sError[256];
            sprintf(sError, "Colon or Equals expected after option %s!",&args[uArg+1][1]);
            m_pError = _strdup(sError);
          }
          args[uArg+1][uPos] = c;
        }
        else
          m_pError = _strdup("Empty option!");
      }
      else
      {
        /* unnamed argument */
        m_psValue[uArg] = _strdup(args[uArg+1]);
      }
    }
  }
} /* CmdLine */

/*--------------------------------------------------------------------*/
CmdLine::~CmdLine()
{
  unsigned uArg;
  if (m_pError != NULL)
    free(m_pError);
  if (m_psName != NULL)
  {
    for (uArg = 0; uArg < m_uArguments; uArg++)
    {
      if (m_psName[uArg] != NULL)
        free(m_psName[uArg]);
    }
    free(m_psName);
  }
  if (m_psValue != NULL)
  {
    for (uArg = 0; uArg < m_uArguments; uArg++)
    {
      if (m_psValue[uArg] != NULL)
        free(m_psValue[uArg]);
    }
    free(m_psValue);
  }
} /* ~CmdLine */

/*--------------------------------------------------------------------*/
char* CmdLine::error()
{
  return m_pError;
} /* error */

/*--------------------------------------------------------------------*/
unsigned CmdLine::arguments()
{
  return m_uArguments;
} /* arguments */

/*--------------------------------------------------------------------*/
char* CmdLine::name(unsigned uArg)
{
  char* pName = NULL;
  if (uArg < m_uArguments)
    pName = m_psName[uArg];
  return pName;
} /* name */

/*--------------------------------------------------------------------*/
char* CmdLine::value(unsigned uArg)
{
  char* pValue = NULL;
  if (uArg < m_uArguments)
    pValue = m_psValue[uArg];
  return pValue;
} /* value */

/*==End of File=======================================================*/
