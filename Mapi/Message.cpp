/*-- $Workfile: Message.cpp $ ------------------------------------------
Wrapper for an IMessage.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Message class wraps an IMessage.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#define _CRT_SECURE_NO_WARNINGS
#include <stdio.h>
#define STRICT
#include <mapix.h>
#include <mapiutil.h>
#include "../Utils/Heap.h"
#include "../Utils/Fs.h"
#include "Message.h"
#include "Storage.h"

/*--------------------------------------------------------------------*/
Message::Message(IMessage* pMessage)
  : Container((IMAPIContainer*)pMessage)
{
  m_pMessage = pMessage;
  m_pMsgSession = NULL;
} /* Message */

/*--------------------------------------------------------------------*/
Message::~Message()
{
  if (m_pMsgSession != NULL)
    CloseIMsgSession(m_pMsgSession);
} /* ~Message */

/*--------------------------------------------------------------------*/
BOOL Message::isIpmNote()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  BOOL bIpmNote = false;
  PropValue* pPv = this->getPropValue(PR_OBJECT_TYPE);
  ULONG ulObjectType = pPv->getLongValue();
  delete pPv;
  if (ulObjectType == MAPI_MESSAGE)
  {
    pPv = this->getPropValue(PR_MESSAGE_CLASS_A);
    if (pPv != NULL)
    {
      char* pMessageClass = pPv->getString8Value();
      if (strcmp(pMessageClass,"IPM.Note") == 0)
        bIpmNote = true;
      delete pMessageClass;
      delete pPv;
    }
  }
#ifdef _DEBUG
  if (!(pHeap->isValid() && (pHeap->hasInitialSize())))
    bIpmNote = false;
  delete pHeap;
#endif
  return bIpmNote;
} /* isIpmNote */

/*--------------------------------------------------------------------*/
BOOL Message::isFromMe()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  BOOL bFromMe = false;
  PropValue* pPv = this->getPropValue(PR_MESSAGE_FLAGS);
  if (pPv != NULL)
  {
    ULONG ulFlags = pPv->getLongValue();
    if ((ulFlags & MSGFLAG_FROMME) != 0)
      bFromMe = true;
    delete pPv;
  }
#ifdef _DEBUG
  if (!(pHeap->isValid() && (pHeap->hasInitialSize())))
    bFromMe = false;
  delete pHeap;
#endif
  return bFromMe;
} /* isFromMe */

/*--------------------------------------------------------------------*/
char* Message::getSenderEmailAddress()
{
  char* pSenderEmailAddress = NULL;
  PropValue* pPv = this->getPropValue(PR_SENDER_EMAIL_ADDRESS_A);
  if (pPv != NULL)
  {
    pSenderEmailAddress = pPv->getString8Value();
    delete pPv;
  }
  if (pSenderEmailAddress == NULL)
  {
    pSenderEmailAddress = new char[1];
    pSenderEmailAddress[0] = '\0';
  }
  return pSenderEmailAddress;
} /* getSenderEmailAddress */

/*--------------------------------------------------------------------*/
char* Message::getReceivedByEmailAddress()
{
  char* pReceivedByEmailAddress = NULL;
  PropValue* pPv = this->getPropValue(PR_RECEIVED_BY_EMAIL_ADDRESS_A);
  if (pPv != NULL)
  {
    pReceivedByEmailAddress = pPv->getString8Value();
    delete pPv;
  }
  if (pReceivedByEmailAddress == NULL)
  {
    pReceivedByEmailAddress = new char[1];
    pReceivedByEmailAddress[0] = '\0';
  }
  return pReceivedByEmailAddress;
} /* getReceivedByEmailAddress */

/*--------------------------------------------------------------------*/
char* Message::getSubject()
{
  char* pSubject = NULL;
  PropValue* pPv = this->getPropValue(PR_SUBJECT_A);
  if (pPv != NULL)
  {
    pSubject = pPv->getString8Value();
    delete pPv;
  }
  if (pSubject == NULL)
  {
    pSubject = new char[1];
    pSubject[0] = '\0';
  }
  return pSubject;
} /* getSubject */

/*--------------------------------------------------------------------*/
FileTime* Message::getLastModificationTime()
{
  FileTime* pLastModificationTime = NULL;
  PropValue* pPv = this->getPropValue(PR_LAST_MODIFICATION_TIME);
  if (pPv != NULL)
  {
    pLastModificationTime = new FileTime(pPv->getSysTimeValue());
    delete pPv;
  }
  return pLastModificationTime;
} /* getLastModificationTime */

/*--------------------------------------------------------------------*/
FileTime* Message::getClientSubmitTime()
{
  FileTime* pClientSubmitTime = NULL;
  PropValue* pPv = this->getPropValue(PR_CLIENT_SUBMIT_TIME);
  if (pPv != NULL)
  {
    pClientSubmitTime = new FileTime(pPv->getSysTimeValue());
    delete pPv;
  }
  return pClientSubmitTime;
} /* getClientSubmitTime */

/*--------------------------------------------------------------------*/
FileTime* Message::getMessageDeliveryTime()
{
  FileTime* pMessageDeliveryTime = NULL;
  PropValue* pPv = this->getPropValue(PR_MESSAGE_DELIVERY_TIME);
  if (pPv != NULL)
  {
    pMessageDeliveryTime = new FileTime(pPv->getSysTimeValue());
    delete pPv;
  }
  return pMessageDeliveryTime;
} /* getMessageDeliveryTime */

/*--------------------------------------------------------------------*/
HRESULT Message::open(Storage* pStorage)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = E_FAIL;
  if ((m_pMessage == NULL) && (m_pMsgSession == NULL))
  {
    // get memory allocation function
    LPMALLOC pMalloc = MAPIGetDefaultMalloc();
    hr = OpenIMsgSession(pMalloc, 0, &this->m_pMsgSession);
    if (SUCCEEDED(hr))
    {
      hr = OpenIMsgOnIStg(m_pMsgSession,
        MAPIAllocateBuffer,MAPIAllocateMore,MAPIFreeBuffer,
        pMalloc,NULL,pStorage->getStg(),NULL,0,0,&m_pMessage);
      if (SUCCEEDED(hr))
        ;// this->setMapiProp(m_pMessage)
      else
      {
        printf("Message could not be opened on storage (0x%08x)!\n",hr);
        CloseIMsgSession(m_pMsgSession);
        m_pMsgSession = NULL;
      }
    }
    else
    {
      printf("MessageSession could not be opened (0x%08x)!\n",hr);
      if (m_pMsgSession != NULL)
        m_pMsgSession = NULL;
    }
  }
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
  {
    hr = E_FAIL;
    printf("Heap inconsistent after Message::open()!\n");
  }
  delete pHeap;
#endif
  return hr;
} /* open */

/*--------------------------------------------------------------------*/
HRESULT Message::save()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = E_FAIL;
  if ((m_pMessage != NULL) && (m_pMsgSession != NULL))
    hr = m_pMessage->SaveChanges(0);
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
    hr = E_FAIL;
  delete pHeap;
#endif
  return hr;
} /* save */

/*--------------------------------------------------------------------*/
HRESULT Message::list()
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = NOERROR;
  char* pEmailAddress;
  if (this->isFromMe())
    pEmailAddress = this->getReceivedByEmailAddress();
  else
    pEmailAddress = this->getSenderEmailAddress();
  char* pSubject = this->getSubject();
  printf("%s: %s\n",pEmailAddress,pSubject);
  hr = Properties::list();
  free(pEmailAddress);
  free(pSubject);
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
    hr = E_FAIL;
  delete pHeap;
#endif
  return hr;
} /* list */

/*--------------------------------------------------------------------*/
HRESULT Message::listFile(ULONG ulMaxFilenameLength, char* pFolderPath)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = NOERROR;
  /* create the file name */
  char* pEmailAddress;
  if (this->isFromMe())
    pEmailAddress = this->getReceivedByEmailAddress();
  else
    pEmailAddress = this->getSenderEmailAddress();
  char* pSubject = this->getSubject();
  char* pAddressSubject = new char[strlen(pEmailAddress)+strlen(pSubject)+2];
  strcpy(pAddressSubject,pEmailAddress);
  pAddressSubject[strlen(pEmailAddress)] = '_';
  strcpy(&pAddressSubject[strlen(pEmailAddress)+1],pSubject);
  Fs* pFsFolder = new Fs(pFolderPath);
  char* pMsgFile = pFsFolder->normalizeFile(pAddressSubject,ulMaxFilenameLength,".msg");
  /* list pMsgFile */
  printf("%s\n",pMsgFile);
  delete pMsgFile;
  delete pFsFolder;
  delete pAddressSubject;
  delete pSubject;
  delete pEmailAddress;
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
  {
    hr = E_FAIL;
    printf("Message::listFile: Heap inconsistent!\n");
  }
  delete pHeap;
#endif
  return hr;
} /* listFile */

/*--------------------------------------------------------------------*/
HRESULT Message::archive(ULONG ulMaxFilenameLength, char* pFolderPath)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  HRESULT hr = E_FAIL;
  /* create the file name */
  char* pEmailAddress;
  if (this->isFromMe())
    pEmailAddress = this->getReceivedByEmailAddress();
  else
    pEmailAddress = this->getSenderEmailAddress();
  char* pSubject = this->getSubject();
  char* pAddressSubject = new char[strlen(pEmailAddress)+strlen(pSubject)+2];
  strcpy(pAddressSubject,pEmailAddress);
  pAddressSubject[strlen(pEmailAddress)] = '_';
  strcpy(&pAddressSubject[strlen(pEmailAddress)+1],pSubject);
  delete pEmailAddress;
  delete pSubject;
  Fs* pFsFolder = new Fs(pFolderPath);
  char* pMsgFile = pFsFolder->normalizeFile(pAddressSubject,ulMaxFilenameLength,".msg");
  delete pFsFolder;
  delete pAddressSubject;
  /* use pMsgFile to create storage */
  Storage* pStorage = new Storage(pMsgFile);
  hr = pStorage->create(CLSID_MailMessage);
  if (!FAILED(hr))
  {
    /* open a msg session and a msg interface on the storage file */
    Message* pMessageStorage = new Message(NULL);
    hr = pMessageStorage->open(pStorage);
    if (!FAILED(hr))
    {
      /* copy this message to the one on storage */
      hr = this->m_pMessage->CopyTo(0,NULL,NULL,NULL,NULL,(LPIID)&IID_IMessage,
        pMessageStorage->m_pMessage,0,NULL);
      if (FAILED(hr))
      {
        printf("Message could not be copied to Storage (0x%08x)!\n",hr);
        LPMAPIERROR pme;
        m_pMessage->GetLastError(hr,0,&pme);
        if (pme != NULL)
        {
          printf("Component: %s, Error: %s\n",
            pme->lpszComponent,pme->lpszError);
          MAPIFreeBuffer(pme);
        }
      }
      /* save */
      pMessageStorage->save();
      /* commit */
      pStorage->commit();
    }
    /* close msg and message session */
    delete pMessageStorage;
    delete pStorage;
    /* get message time to be stored as file time */
    FileTime* pft = this->getMessageDeliveryTime();
    if (pft == NULL)
    {
      pft = this->getClientSubmitTime();
      if (pft == NULL)
        pft = this->getLastModificationTime();
    }
    if (pft != NULL)
    {
      // pft->list();
      Fs* pfsFullFile = new Fs(pMsgFile);
      pfsFullFile->setTimes(pft,pft,pft);
      delete pfsFullFile;
      delete pft;
    }
  }
  delete pMsgFile;
#ifdef _DEBUG
  if (!(SUCCEEDED(hr) && pHeap->isValid() && (pHeap->hasInitialSize())))
  {
    hr = E_FAIL;
    printf("Message::archive: Heap inconsistent!\n");
  }
  delete pHeap;
#endif
  if (FAILED(hr))
    printf("Message could not be archived (0x%08x)!\n",hr);
  return hr;
} /* archive */

/*==End of File=======================================================*/
