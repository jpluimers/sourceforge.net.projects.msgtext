/*-- $Workfile: Container.h $ ------------------------------------------
Wrapper for an IMAPIContainer.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Container class wraps an IMAPIContainer.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#ifndef CONTAINER_H
#define CONTAINER_H

#define STRICT
#include <mapix.h>
#include "Properties.h"

/*======================================================================
Container is the IMAPIContainer wrapper
======================================================================*/
class Container : public Properties
{
public:
  /*--------------------------------------------------------------------
  constructor 
    pMapiContainer: pointer to IMAPIContainer interface.
  heap size not invariant: establishes list of PropValues.
  --------------------------------------------------------------------*/
  Container(IMAPIContainer* pMapiContainer);

  /*--------------------------------------------------------------------
  destructor releases IMAPIContainer interface.
  heap size not invariant: cleans up list of PropValues.
  --------------------------------------------------------------------*/
  ~Container();

}; /* class Container */

#endif // CONTAINER_H not defined
/*==End of File=======================================================*/
