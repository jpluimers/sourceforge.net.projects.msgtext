/*-- $Workfile: Container.h $ ------------------------------------------
Command-line parsing
Version    : $Id$
Application: Utilities
Platform   : Windows
Description: This object parses the command-line into named and unnamed
             arguments. Named arguments start with "/" or "-" followed
             by the name (consisting of A-Z, a-z or 0-9) of the argument, 
             followed by a colon or an equal, and the value of the argument.
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
#ifndef CMDLINE_H
#define CMDLINE_H

#define STRICT
#include <windows.h>

/*======================================================================
CmdLine contains digested args
======================================================================*/
class CmdLine
{
public:
  /*--------------------------------------------------------------------
  constructor 
    argc : number of arguments
    args : array of char* representing the arguments
  heap size not invariant: establishes arrays of names and values.
  --------------------------------------------------------------------*/
  CmdLine(int argc, char** args);

  /*--------------------------------------------------------------------
  destructor
  heap size not invariant: cleans up arrays of names and values.
  --------------------------------------------------------------------*/
  ~CmdLine();

  /*--------------------------------------------------------------------
  error
  return an error string, if parsing encountered an error, NULL if no
  error was encountered.
  --------------------------------------------------------------------*/
  char* error();

  /*--------------------------------------------------------------------
  size
  return size of arguments array
  --------------------------------------------------------------------*/
  unsigned arguments();

  /*--------------------------------------------------------------------
  name
  return name of ith argument, NULL, if unnamed uArg >= size()
  --------------------------------------------------------------------*/
  char* name(unsigned uArg);

  /*--------------------------------------------------------------------
  value
  return value of ith argument, NULL if value is true or uArg >= size()
  --------------------------------------------------------------------*/
  char* value(unsigned uArg);

private:
  /* table of names */
  char** m_psName;
  /* table of values */
  char** m_psValue;
  /* number of arguments */
  unsigned m_uArguments;
  /* error message if parsing failed */
  char* m_pError;
}; /* class CmdLine */

#endif // CMDLINE_H not defined

/*==End of File=======================================================*/
