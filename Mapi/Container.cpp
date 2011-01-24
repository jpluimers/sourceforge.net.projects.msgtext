/*-- $Workfile: Container.cpp $ ----------------------------------------
Wrapper for an IMAPIContainer.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Container class wraps an IMAPIContainer.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#include <stdio.h>
#define STRICT
#include <mapix.h>
#include "Container.h"

/*--------------------------------------------------------------------*/
Container::Container(IMAPIContainer* pMapiContainer) 
  : Properties((IMAPIProp*)pMapiContainer)
{
} /* Container */

/*--------------------------------------------------------------------*/
Container::~Container()
{
} /* ~Container */

/*==End of File=======================================================*/
