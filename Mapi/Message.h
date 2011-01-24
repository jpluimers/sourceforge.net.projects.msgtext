/*-- $Workfile: Message.h $ --------------------------------------------
Wrapper for an IMessage.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Message class wraps an IMessage.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#ifndef MESSAGE_H
#define MESSAGE_H

#define STRICT
#include <mapix.h>
#include <imessage.h>
#include "../Utils/FileTime.h"
#include "Container.h"
#include "Storage.h"

/*======================================================================
Message is the IMessage wrapper
======================================================================*/
class Message : public Container
{
public:
  /*--------------------------------------------------------------------
  constructor 
    pMessage: pointer to IMessage interface or NULL, if to be opened
              later.
  heap size invariant.
  --------------------------------------------------------------------*/
  Message(IMessage* pMessage);

  /*--------------------------------------------------------------------
  destructor releases IMessage interface and closes open message session.
  heap size invariant.
  --------------------------------------------------------------------*/
  ~Message();

  /*--------------------------------------------------------------------
  isIpmNote
  ¨ returns: true, if IMessage is of type "IPM.Note" (i.e. real mail)
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL isIpmNote();

  /*--------------------------------------------------------------------
  isFromMe
  ¨ returns: true, if IMessage originated from current user.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL isFromMe();

  /*--------------------------------------------------------------------
  getReceivedByEmailAddress
  ¨ returns: email address of receiver or NULL if an error occurred.
  heap size not invariant: caller must delete returned string.
  --------------------------------------------------------------------*/
  char* getReceivedByEmailAddress();

  /*--------------------------------------------------------------------
  getSenderEmailAddress
  ¨ returns: email address of sender or NULL if an error occurred.
  heap size not invariant: caller must delete returned string.
  --------------------------------------------------------------------*/
  char* getSenderEmailAddress();

  /*--------------------------------------------------------------------
  getSubject
  ¨ returns: subject or NULL if an error occurred.
  heap size not invariant: caller must delete returned string.
  --------------------------------------------------------------------*/
  char* getSubject();

  /*--------------------------------------------------------------------
  getLastModifiedTime
  ¨ returns: last modified time or NULL, if an error occurred.
  heap size not invariant: caller must delete returned object.
  --------------------------------------------------------------------*/
  FileTime* getLastModificationTime();

  /*--------------------------------------------------------------------
  getClientSubmitTime
  ¨ returns: client submit time or NULL, if an error occurred.
  heap size not invariant: caller must delete returned object.
  --------------------------------------------------------------------*/
  FileTime* getClientSubmitTime();

  /*--------------------------------------------------------------------
  getMessageDeliveryTime
  ¨ returns: message delivery time or NULL, if an error occurred.
  heap size not invariant: caller must delete returned object.
  --------------------------------------------------------------------*/
  FileTime* getMessageDeliveryTime();

  /*--------------------------------------------------------------------
  open opens a message on a storage object.
  ¨ returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT open(Storage* pStorage);

  /*--------------------------------------------------------------------
  save saves the changes of the message to the storage object.
  ¨ returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT save();

  /*--------------------------------------------------------------------
  list displays receiver/sender and subject of message.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT list();

  /*--------------------------------------------------------------------
  listFile lists file name under which this message would be stored.
    The file name starts with the receiver's or sender's email address
    followed by the subject. It is then normalized and truncated to the 
    maximum filename length as well as disabiguated.
    ulMaxFilenameLength: maximum folder/file name length.
    pFolderPath: directory in which message is to be stored.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT listFile(ULONG ulMaxFilenameLength, char* pFolderPath);

  /*--------------------------------------------------------------------
  archive stores the message as a .msg file in given directory.
    The file name starts with the receiver's or sender's email address
    followed by the subject. It is then normalized and truncated to the 
    maximum filename length as well as disabiguated.
    ulMaxFilenameLength: maximum folder/file name length.
    pFolderPath: directory in which message is to be stored.
  ¨ returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT archive(ULONG ulMaxFilenameLength, char* pFolderPath);

private:
  /* pointer to IMessage interface or NULL, if not yet opened */
  IMessage* m_pMessage;
  /* pointer to MSGSESS interface if open on a storage object. */
  LPMSGSESS m_pMsgSession;
}; /* class Message */

#endif // MESSAGE_H not defined
/*==End of File=======================================================*/
