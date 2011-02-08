/*-- $Workfile: Msg.cpp $ ----------------------------------------------
Class representiong a Message object in a Storage file.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Msg class extends a Storage object.
             See MS-OXMSG.pdf (http://msdn.microsoft.com/en-us/library/cc463912.aspx)
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
#include "../Utils/Fs.h"
#include "../Utils/Heap.h"
#include "../Rtf/CompressedStream.h"
#include "../Rtf/HtmlWriter.h"
#include "Msg.h"

/*--------------------------------------------------------------------*/
void Msg::clear()
{
  m_apwcsRecipient = NULL;
  m_iRecipientCount = 0;
  m_apwcsAttachment = NULL;
  m_iAttachmentCount = 0;
  m_bEmbedded = false;
} /* clear */

ULONG ulEMBEDDED_HEADER_SIZE = 24;
ULONG ulROOT_HEADER_SIZE = 32;
/*--------------------------------------------------------------------*/
LPBYTE Msg::header(LPBYTE pHeader)
{
  /* 8 4-byte values:
    dwReserved1,
    dwReserved2,
    dwNextRecipientId, (to use for naming next recipient)
    dwNextAttachmentId, (to use for naming next attachment)
    dwRecipientCount, (could be checked against count determined by initialize)
    dwAttachmentCount, ( " )
    dwReserved3,
    dwReserved4.
    Normally dwNextRecipientId = dwRecipientCount and
    dwNextAttachmentId = dwAttachmentCount. */
  if (m_bEmbedded)
    pHeader += ulEMBEDDED_HEADER_SIZE;
  else
    pHeader += ulROOT_HEADER_SIZE;
  return pHeader;
} /* header */

/*--------------------------------------------------------------------*/
HRESULT Msg::initialize()
{
  LPOLESTR awcNAMEID = L"__nameid_version1.0";
  LPOLESTR awcATTACH = L"__attach_version1.0";
  LPOLESTR awcRECIP = L"__recip_version1.0";
  ULONG ulSTORE_UNICODE_OK = 0x00040000;
  int iStorage;
  LPOLESTR pwcsName;
  LPOLESTR pwcsPrefix;
  Property* pPropStoreSupportMask;
  ULONG ulPropStoreSupportMask;
  HRESULT hr = PropStorage::initialize();
  if (SUCCEEDED(hr))
  {
    /* lookup PR_STORE_SUPPORT_MASK for UNICODE flag */
    pPropStoreSupportMask = getProperty(PR_STORE_SUPPORT_MASK);
    if (pPropStoreSupportMask != NULL)
    {
      ulPropStoreSupportMask = pPropStoreSupportMask->getLongValue();
      if ((ulPropStoreSupportMask & ulSTORE_UNICODE_OK) == ulSTORE_UNICODE_OK)
        m_bUnicode = true;
    }
    /* loop over contained storages */
    for (iStorage = 0; SUCCEEDED(hr) && (iStorage < getStorageCount()); iStorage++)
    {
      pwcsName = getStorageName(iStorage);
      if (wcscmp(pwcsName,awcNAMEID) == 0)
      {
        /* named properties are not really needed anywhere ... */
        ;
      }
      else
      {
        /* if it starts with __attach_version1.0_#HHHHHHHH */
        pwcsPrefix = _wcsdup(pwcsName);
        if (wcslen(pwcsPrefix) > wcslen(awcATTACH))
          pwcsPrefix[wcslen(awcATTACH)] = L'\0';
        if (wcscmp(pwcsPrefix,awcATTACH) == 0)
          m_iAttachmentCount = append((void***)&m_apwcsAttachment,m_iAttachmentCount,pwcsName);
        else
        {
          /* if it starts with __recip_version1.0_#HHHHHHHH */
          pwcsPrefix[wcslen(awcRECIP)] = L'\0';
          if (wcscmp(pwcsPrefix,awcRECIP) == 0)
            m_iRecipientCount = append((void***)&m_apwcsRecipient,m_iRecipientCount,pwcsName);
          else
          {
            printf("invalid entry found:%S\n",pwcsName);
            hr = E_FAIL;
          }
        }
        free(pwcsPrefix);
      }
    }
  }
  return hr;
} /* initialize */

/*--------------------------------------------------------------------*/
Msg::Msg(char* pName): PropStorage(pName)
{
  clear();
} /* Msg */

/*--------------------------------------------------------------------*/
Msg::Msg(IStorage* pStorage, BOOL bUnicode): PropStorage(pStorage, bUnicode)
{
  clear();
  m_bEmbedded = true;
} /* Msg */

/*--------------------------------------------------------------------*/
Msg::~Msg()
{
  if (m_apwcsRecipient != NULL)
    free(m_apwcsRecipient);
  if (m_apwcsAttachment != NULL)
    free(m_apwcsAttachment);
} /* ~Msg */

/*--------------------------------------------------------------------*/
HRESULT Msg::open()
{
  HRESULT hr = PropStorage::open();
  // has called initialize() already
  if (SUCCEEDED(hr))
  {
    if (!matches(CLSID_MailMessage) || (type() != STGTY_STORAGE))
    {
      LPOLESTR pwcsName = name();
      printf("Storage %S is not a mail message!",pwcsName);
      CoTaskMemFree(pwcsName);
      hr = E_FAIL;
    }
  }
  return hr;
} /* open */

/*--------------------------------------------------------------------*/
FileTime* Msg::getLastModificationTime()
{
  return getSysTimeValue(PR_LAST_MODIFICATION_TIME);
} /* getLastModificationTime */

/*--------------------------------------------------------------------*/
FileTime* Msg::getClientSubmitTime()
{
  return getSysTimeValue(PR_CLIENT_SUBMIT_TIME);
} /* getClientSubmitTime */

/*--------------------------------------------------------------------*/
FileTime* Msg::getMessageDeliveryTime()
{
  return getSysTimeValue(PR_MESSAGE_DELIVERY_TIME);
} /* getMessageDeliveryTime */

/*--------------------------------------------------------------------*/
char* Msg::getSenderName()
{
  return getStringValue(PR_SENDER_NAME);
} /* getSenderName */

/*--------------------------------------------------------------------*/
char* Msg::getSenderEmailAddress()
{
  return getStringValue(PR_SENDER_EMAIL_ADDRESS);
} /* getSenderEmailAddress */

/*--------------------------------------------------------------------*/
char* Msg::getSubject()
{
  return getStringValue(PR_SUBJECT);
} /* getSubject */

/*--------------------------------------------------------------------*/
char* Msg::getBody()
{
  return getStringValue(PR_BODY);
} /* getBody */

/*--------------------------------------------------------------------*/
BOOL Msg::hasRtf()
{
  return (getProperty(PR_RTF_COMPRESSED) != NULL);
} /* hasRtf */

/*--------------------------------------------------------------------*/
char* getMailAddress(char* pDisplayName, char* pEmailAddress)
{
  char* pMailAddress = NULL;
  if (pEmailAddress == NULL)
    pMailAddress = _strdup(pDisplayName);
  else
  {
    int iLength = strlen(pEmailAddress)+2; /* '<address>' */
    if (pDisplayName != NULL)
      iLength += strlen(pDisplayName) + 3; /* '"name" ' */
    iLength += 1; /* terminating '\0' */
    pMailAddress = (char*)malloc(iLength);
    if (pDisplayName == NULL)
      sprintf(pMailAddress,"<%s>",pEmailAddress);
    else
      sprintf(pMailAddress,"\"%s\" <%s>",pDisplayName,pEmailAddress);
  }
  return pMailAddress;
} /* getMailAddress */

/*--------------------------------------------------------------------*/
int Msg::getRecipientCount()
{
  return m_iRecipientCount;
} /* getRecipientCount */

/*--------------------------------------------------------------------*/
LPOLESTR Msg::getRecipientName(int iRecipient)
{
  LPOLESTR pwcsName = NULL;
  if ((iRecipient >= 0) && (iRecipient < m_iRecipientCount))
    pwcsName = m_apwcsRecipient[iRecipient];
  return pwcsName;
} /* getRecipientName */

/*--------------------------------------------------------------------*/
Recipient* Msg::getRecipient(LPOLESTR pwcsName)
{
  Recipient* pRecipient = new Recipient(getIStorage(pwcsName),m_bUnicode);
  HRESULT hr = pRecipient->initialize();
  if (FAILED(hr))
  {
    delete pRecipient;
    pRecipient = NULL;
  }
  return pRecipient;
} /* getRecipient */

/*--------------------------------------------------------------------*/
int Msg::getAttachmentCount()
{
  return m_iAttachmentCount;
} /* getAttachmentCount */

/*--------------------------------------------------------------------*/
LPOLESTR Msg::getAttachmentName(int iAttachment)
{
  LPOLESTR pwcsName = NULL;
  if ((iAttachment >= 0) && (iAttachment < m_iAttachmentCount))
    pwcsName = m_apwcsAttachment[iAttachment];
  return pwcsName;
} /* getAttachmentName */

/*--------------------------------------------------------------------*/
Attachment* Msg::getAttachment(LPOLESTR pwcsName)
{
  Attachment* pAttachment = new Attachment(getIStorage(pwcsName),m_bUnicode);
  HRESULT hr = pAttachment->initialize();
  if (FAILED(hr))
  {
    delete pAttachment;
    pAttachment = NULL;
  }
  return pAttachment;
} /* getAttachment */

/*--------------------------------------------------------------------*/
HRESULT Msg::copyTo(Storage* pStorage)
{
  HRESULT hr = S_OK;
  LPOLESTR awcPROPERTIES = L"__properties_version1.0";
  // copy all storages to pStorage
  for (int iStorage = 0; SUCCEEDED(hr) && (iStorage < getStorageCount()); iStorage++)
    hr = copyStorage(getStorageName(iStorage),pStorage);
  // copy all streams except __properties1.0 to pStorage
  for (int iStream = 0; SUCCEEDED(hr) && (iStream < getStreamCount()); iStream++)
  {
    if (wcscmp(awcPROPERTIES,getStreamName(iStream)) != 0)
      hr = copyStream(getStreamName(iStream),pStorage);
  }
  if (SUCCEEDED(hr))
  {
    hr = E_FAIL;
    /* copy __properties1.0 to pStorage, with new header */
    Stream* pSourcePropertiesStream = getStream(awcPROPERTIES);
    if (pSourcePropertiesStream != NULL)
    {
      Stream* pTargetPropertiesStream = pStorage->createStream(awcPROPERTIES);
      if (pTargetPropertiesStream != NULL)
      {
        ULONG ulSourceSize = pSourcePropertiesStream->size();
        ULONG ulHeaderDifference = ulROOT_HEADER_SIZE - ulEMBEDDED_HEADER_SIZE;
        // hr = pTargetPropertiesStream->setSize(ulSourceSize+ulHeaderDifference);
        ULONG ulSize = BUFSIZ;
        char* pBuffer = (char*)malloc(ulSize);
        // header
        if (pSourcePropertiesStream->read(ulEMBEDDED_HEADER_SIZE,pBuffer) == ulEMBEDDED_HEADER_SIZE)
        {
          ulSourceSize -= ulEMBEDDED_HEADER_SIZE;
          // root header has two reserved DWORDS equal to 0
          memset(&pBuffer[ulEMBEDDED_HEADER_SIZE],0,ulROOT_HEADER_SIZE - ulEMBEDDED_HEADER_SIZE);
          if (pTargetPropertiesStream->write(ulROOT_HEADER_SIZE,pBuffer))
          {
            // rest
            ULONG ulSize = BUFSIZ;
            for (hr = S_OK; (SUCCEEDED(hr)) && (ulSourceSize > 0); ulSourceSize -= ulSize)
            {
              if (ulSize > ulSourceSize)
                ulSize = ulSourceSize;
              hr = E_FAIL;
              if (pSourcePropertiesStream->read(ulSize,pBuffer))
              {
                if (pTargetPropertiesStream->write(ulSize,pBuffer))
                  hr = S_OK;
              }
            }
          }
        }
        free(pBuffer);
        pTargetPropertiesStream->commit();
        delete pTargetPropertiesStream;
      }
      delete pSourcePropertiesStream;
    }
  }
  return hr;
} /* copyTo */

/*--------------------------------------------------------------------*/
BOOL Msg::saveEmbeddedMsg(Attachment* pAttachment,char* pFilename)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  BOOL bSucceeded = false;
  LPOLESTR awcNAMEID = L"__nameid_version1.0";
  IStorage* pStorage = pAttachment->getMsgStorage();
  if (pStorage != NULL)
  {
    Msg* pEmbeddedMsg = new Msg(pStorage,m_bUnicode);
    HRESULT hr = pEmbeddedMsg->initialize();
    if (SUCCEEDED(hr))
    {
      Storage* pSaveStorage = new Storage(pFilename);
      hr = pSaveStorage->create(CLSID_MailMessage);
      if (SUCCEEDED(hr))
      {
        hr = copyStorage(awcNAMEID,pSaveStorage);
        if (SUCCEEDED(hr))
          hr = pEmbeddedMsg->copyTo(pSaveStorage);
      }
      if (SUCCEEDED(hr))
        hr = pSaveStorage->commit();
      delete pSaveStorage;
      /* set this message's time as file time of embedded message */
      if (SUCCEEDED(hr))
      {
        FileTime* pft = getMessageDeliveryTime();
        if (pft == NULL)
        {
          pft = getClientSubmitTime();
          if (pft == NULL)
            pft = getLastModificationTime();
        }
        if (pft != NULL)
        {
          Fs* pfsSaveMsg = new Fs(pFilename);
          if (pfsSaveMsg->setTimes(pft,pft,pft) == 0)
            bSucceeded = true;
          delete pfsSaveMsg;
          delete pft;
        }
      }
    }
    delete pEmbeddedMsg;
  }
#ifdef _DEBUG
  if (!(bSucceeded && pHeap->isValid() && pHeap->hasInitialSize()))
  {
    printf("Msg::saveEmbeddedMsg: Heap inconsistent after saveEmbeddedMsg()!\n");
    bSucceeded = false;
  }
  delete pHeap;
#endif
  return bSucceeded;
} /* saveEmbeddedMsg */

/*--------------------------------------------------------------------*/
BOOL Msg::saveFrom(FILE* pfile)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  BOOL bSucceeded = true;
  char* pSenderEmailAddress = getSenderEmailAddress();
  if (pSenderEmailAddress != NULL)
  {
    char* pSenderName = getSenderName();
    char* pFromAddress = getMailAddress(pSenderName, pSenderEmailAddress);
    fprintf(pfile,"From: %s\r\n",pFromAddress);
    free(pFromAddress);
    if (pSenderName != NULL)
      free(pSenderName);
    free(pSenderEmailAddress);
  }
#ifdef _DEBUG
  if (!(bSucceeded && pHeap->isValid() && pHeap->hasInitialSize()))
  {
    printf("Msg::saveFrom: Heap inconsistent after saveDate()!\n");
    bSucceeded = false;
  }
  delete pHeap;
#endif
  return bSucceeded;
} /* saveFrom */

/*--------------------------------------------------------------------*/
BOOL Msg::saveToCcBcc(FILE* pfile)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  BOOL bSucceeded = false;
  /* To, Cc, Bcc */
  char* pTo = NULL;
  char* pCc = NULL;
  char* pBcc = NULL;
  for (int iRecipient = 0; iRecipient < getRecipientCount(); iRecipient++)
  {
    Recipient* pRecipient = getRecipient(getRecipientName(iRecipient));
    char* pEmailAddress = pRecipient->getEmailAddress();
    char* pDisplayName = pRecipient->getDisplayName();
    char* pAddress = getMailAddress(pDisplayName,pEmailAddress);
    if (pEmailAddress != NULL)
      free(pEmailAddress);
    if (pDisplayName != NULL)
      free(pDisplayName);
    ULONG ulType = pRecipient->getType();
    switch(ulType)
    {
      case MAPI_TO:
        if (pTo == NULL)
        {
          pTo = (char*)malloc(strlen(pAddress)+5);
          sprintf(pTo,"To: %s",pAddress);
        }
        else
        {
          pTo = (char*)realloc(pTo,strlen(pTo)+strlen(pAddress)+3);
          sprintf(pTo,"%s, %s",pTo,pAddress);
        }
        break;
      case MAPI_CC:
        if (pCc == NULL)
        {
          pCc = (char*)malloc(strlen(pAddress)+5);
          sprintf(pCc,"Cc: %s",pAddress);
        }
        else
        {
          pCc = (char*)realloc(pCc,strlen(pCc)+strlen(pAddress)+3);
          sprintf(pCc,"%s, %s",pCc,pAddress);
        }
        break;
      case MAPI_BCC:
        if (pBcc == NULL)
        {
          pBcc = (char*)malloc(strlen(pAddress)+6);
          sprintf(pBcc,"Bcc: %s",pAddress);
        }
        else
        {
          pBcc = (char*)realloc(pBcc,strlen(pBcc)+strlen(pAddress)+3);
          sprintf(pBcc,"%s, %s",pBcc,pAddress);
        }
        break;
      default:
        break;
    }
    free(pAddress);
    delete pRecipient;
  }
  if (pTo != NULL)
  {
    fprintf(pfile,"%s\r\n",pTo);
    free(pTo);
  }
  if (pCc != NULL)
  {
    fprintf(pfile,"%s\r\n",pCc);
    free(pCc);
  }
  if (pBcc != NULL)
  {
    fprintf(pfile,"%s\r\n",pBcc);
    free(pBcc);
  }
  bSucceeded = true;
#ifdef _DEBUG
  if (!(bSucceeded && pHeap->isValid() && pHeap->hasInitialSize()))
  {
    printf("Msg::saveToCcBcc: Heap inconsistent after saveToCcBcc()!\n");
    bSucceeded = false;
  }
  delete pHeap;
#endif
  return bSucceeded;
} /* saveToCcBcc */

/*--------------------------------------------------------------------*/
BOOL Msg::saveDate(FILE* pfile)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  BOOL bSucceeded = false;
  ULONG ulFlags = MSGFLAG_UNSENT;
  Property* pPropFlags = getProperty(PR_MESSAGE_FLAGS);
  if (pPropFlags != NULL)
    ulFlags = pPropFlags->getLongValue();
  /* no "Date:" header if unsent */
  if ((ulFlags & MSGFLAG_UNSENT) == 0) 
  {
    FileTime* pTimeDelivery = getMessageDeliveryTime();
    if (pTimeDelivery != NULL)
    {
      char* pInternetTime = pTimeDelivery->getInternetTime();
      delete pTimeDelivery;
      fprintf(pfile,"Date: %s\r\n",pInternetTime);
      free(pInternetTime);
    }
  }
  bSucceeded = true;
#ifdef _DEBUG
  if (!(bSucceeded && pHeap->isValid() && pHeap->hasInitialSize()))
  {
    printf("Msg::saveDate: Heap inconsistent after saveDate()!\n");
    bSucceeded = false;
  }
  delete pHeap;
#endif
  return bSucceeded;
} /* saveDate */

/*--------------------------------------------------------------------*/
BOOL Msg::saveSubject(FILE* pfile)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  BOOL bSucceeded = false;
  char* pSubject = getSubject();
  if (pSubject != NULL)
  {
    fprintf(pfile,"Subject: %s\r\n",pSubject);
    free(pSubject);
  }
  bSucceeded = true;
#ifdef _DEBUG
  if (!(bSucceeded && pHeap->isValid() && pHeap->hasInitialSize()))
  {
    printf("Msg::saveSubject: Heap inconsistent after saveSubject()!\n");
    bSucceeded = false;
  }
  delete pHeap;
#endif
  return bSucceeded;
} /* saveSubject */

/*--------------------------------------------------------------------*/
BOOL Msg::saveBody(FILE* pfile)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  BOOL bSucceeded = false;
  char* pBody = getBody();
  if (pBody != NULL)
  {
    if ((strlen(pBody) >= 2) && (strcmp(&pBody[strlen(pBody)-2],"\r\n") == 0))
      fprintf(pfile,"%s",pBody);
    else
      fprintf(pfile,"%s\r\n",pBody);
    free(pBody);
  }
  bSucceeded = true;
#ifdef _DEBUG
  if (!(bSucceeded && pHeap->isValid() && pHeap->hasInitialSize()))
  {
    printf("Msg::saveBody: Heap inconsistent after saveBody()!\n");
    bSucceeded = false;
  }
  delete pHeap;
#endif
  return bSucceeded;
} /* saveBody */

/*--------------------------------------------------------------------*/
BOOL Msg::saveHeader(FILE* pfile)
{
#ifdef _DEBUG
  Heap* pHeap = new Heap();
#endif
  BOOL bSucceeded = false;
  char* pHeader = getStringValue(PR_TRANSPORT_MESSAGE_HEADERS);
  if (pHeader != NULL)
  {
    if ((strlen(pHeader) >= 4) && (strcmp(&pHeader[strlen(pHeader)-4],"\r\n\r\n") == 0))
      fprintf(pfile,"%s",pHeader);
    else if ((strlen(pHeader) >= 2) && (strcmp(&pHeader[strlen(pHeader)-2],"\r\n") == 0))
      fprintf(pfile,"%s\r\n",pHeader);
    else
      fprintf(pfile,"%s\r\n\r\n",pHeader);
    free(pHeader);
    bSucceeded = true;
  }
  else
  {
    if (saveFrom(pfile))
    {
      if (saveToCcBcc(pfile))
      {
        if (saveDate(pfile))
        {
          if (saveSubject(pfile))
          {
            /* empty line separating header from body */
            fprintf(pfile,"\r\n");
            bSucceeded = true;
          }
        }
      }
    }
  }
#ifdef _DEBUG
  if (!(bSucceeded && pHeap->isValid() && pHeap->hasInitialSize()))
  {
    printf("Msg::saveHeader: Heap inconsistent after saveHeader()!\n");
    bSucceeded = false;
  }
  delete pHeap;
#endif
  return bSucceeded;
} /* saveHeader */

/*--------------------------------------------------------------------*/
BOOL Msg::saveText(char* pFilename)
{
  BOOL bSucceeded = false;
  FILE* pfile = fopen(pFilename,"wb");
  if (pfile != NULL)
  {
    char* pBuf = (char*)malloc(BUFSIZ);
    setvbuf(pfile,pBuf,_IOFBF,BUFSIZ);
#ifdef _DEBUG
    Heap* pHeap = new Heap();
#endif
    if (saveHeader(pfile))
    {
      if (saveBody(pfile))
        bSucceeded = true;
    }
#ifdef _DEBUG
    if (!(bSucceeded && pHeap->isValid() && pHeap->hasInitialSize()))
    {
      printf("Msg::saveText: Heap inconsistent in saveText()!\n");
      bSucceeded = false;
    }
    delete pHeap;
#endif
    fclose(pfile);
    free(pBuf);
  }
  return bSucceeded;
} /* saveText */

/*--------------------------------------------------------------------*/
BOOL Msg::saveRtf(BOOL bListFiles, char* pFilename, char* pAttachmentDir)
{
  BOOL bSucceeded = false;
  Property* pPropCompressedRtf = getProperty(PR_RTF_COMPRESSED);
  if (pPropCompressedRtf != NULL)
  {
    IStream* pStream = pPropCompressedRtf->getBinaryIStream();
    if (pStream != NULL)
    {
      CompressedStream* pCompRtf = new CompressedStream(pStream);
      HtmlWriter* pHtmlWriter = new HtmlWriter(bListFiles,pAttachmentDir,pFilename);
      bSucceeded = true;
      ULONG ulRead = BUFSIZ;
      char* pBuffer = (char*)malloc(ulRead);
      for (ULONG ulSize = pCompRtf->size(); bSucceeded && (ulRead > 0); ulSize -= ulRead)
      {
        bSucceeded = false;
        if (ulRead > ulSize)
          ulRead = ulSize;
        ulRead = pCompRtf->read(ulRead,pBuffer);
        if (pHtmlWriter->write(ulRead,pBuffer))
          bSucceeded = true;
      }
      if (bSucceeded)
        bSucceeded = !pCompRtf->isCorrupt();
      if (!pHtmlWriter->close())
        bSucceeded = false;
      delete pHtmlWriter;
      delete pCompRtf;
    }
  }
  return bSucceeded;
} /* saveRtf */

/*--------------------------------------------------------------------*/
BOOL Msg::save(BOOL bListFiles, char* pFilename, char* pAttachmentDir)
{
  printf("%s\n",pFilename);
  BOOL bSucceeded = true;
  if (!bListFiles)
    bSucceeded = saveText(pFilename);
  if (bSucceeded)
  {
    if (hasRtf())
    {
      Fs* pfsText = new Fs(pFilename);
      char* pRtfName = pfsText->newFilename();
      delete pfsText;
      bSucceeded = saveRtf(bListFiles, pRtfName, pAttachmentDir);
    }
    /* save all attachments */
    if (getAttachmentCount() > 0)
    {
      /* if the attachment dir does not exist, create it */
      Fs* pfsAttachFolder = new Fs(pAttachmentDir);
      if (!(bListFiles || pfsAttachFolder->exists()))
        pfsAttachFolder->createFolder();
      for (int iAttachment = 0; bSucceeded && (iAttachment < getAttachmentCount()); iAttachment++)
      {
        Attachment* pAttachment = getAttachment(getAttachmentName(iAttachment));
        char* pAttachmentFilename = pAttachment->getFilename();
        char* pAttachmentName = pfsAttachFolder->newFile(pAttachmentFilename);
        Fs* pfsAttachment = new Fs(pAttachmentName);
        if (pfsAttachment->exists())
        {
          char* pExtension = pfsAttachment->extension();
          pAttachmentFilename[strlen(pAttachmentFilename)-strlen(pExtension)] = '\0';
          /* disambîguate */
          pAttachmentName = pfsAttachFolder->newFile(pAttachmentFilename,MAX_PATH,pExtension);
          free(pExtension);
          pfsAttachment = new Fs(pAttachmentName);
        }
        printf("%s\n",pAttachmentName);
        free(pAttachmentFilename);
        if (pAttachment->isEmbeddedMsg())
        {
          if (!bListFiles)
          {
            HRESULT hr = saveEmbeddedMsg(pAttachment,pAttachmentName);
            if (FAILED(hr))
              bSucceeded = false;
          }
        }
        else
        {
          if (!bListFiles)
            bSucceeded = pAttachment->save(pAttachmentName);
        }
        free(pAttachmentName);
        delete pfsAttachment;
        delete pAttachment;
      }
      delete pfsAttachFolder;
    }
  }
  return bSucceeded;
} /* save */

/*==End of File=======================================================*/
