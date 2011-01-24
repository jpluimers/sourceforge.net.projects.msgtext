/*-- $Workfile: Msg.h $ ------------------------------------------------
Class representiong a Message object in a Storage file.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: Msg class extends a PropStorage object.
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
#ifndef MSG_H
#define MSG_H

#define STRICT
#include <windows.h>
#include "PropStorage.h"
#include "Recipient.h"
#include "Attachment.h"

/*======================================================================
Msg extends PropStorage
======================================================================*/
class Msg : public PropStorage
{
public:
  /*--------------------------------------------------------------------
  constructor 
    pFilename: name of storage file.
  heap size invariant.
  --------------------------------------------------------------------*/
  Msg(char* pFilename);

  /*--------------------------------------------------------------------
  destructor releases IStorage interface.
  heap size invariant.
  --------------------------------------------------------------------*/
  ~Msg();

  /*--------------------------------------------------------------------
  open opens an existing storage file.
    returns: 0, if no error occurred.
  not heap size invariant: arrays will only be released in destructor.
  --------------------------------------------------------------------*/
  HRESULT open();

  /*--------------------------------------------------------------------
  getLastModificationTime
    returns: the time stamp of last modification of message.
  heap size not invariant: Caller must delete returned object.
  --------------------------------------------------------------------*/
  FileTime* getLastModificationTime();

  /*--------------------------------------------------------------------
  getClientSubmitTime
    returns: the time stamp of message submission (sending).
  heap size not invariant: Caller must delete returned object.
  --------------------------------------------------------------------*/
  FileTime* getClientSubmitTime();

  /*--------------------------------------------------------------------
  getMessageDeliveryTime
    returns: the time stamp of message delivery (receiving).
  heap size not invariant: Caller must delete returned object.
  --------------------------------------------------------------------*/
  FileTime* getMessageDeliveryTime();

  /*--------------------------------------------------------------------
  getSenderName
    returns: the name of the sender.
  heap size not invariant: Caller must free returned string.
  --------------------------------------------------------------------*/
  char* getSenderName();

  /*--------------------------------------------------------------------
  getSenderEmailAddress
    returns: the email address of the sender.
  heap size not invariant: Caller must free returned string.
  --------------------------------------------------------------------*/
  char* getSenderEmailAddress();

  /*--------------------------------------------------------------------
  getSubject
    returns: the subject of the mail message.
  heap size not invariant: Caller must free returned string.
  --------------------------------------------------------------------*/
  char* getSubject();

  /*--------------------------------------------------------------------
  getBody
    returns: the body of the mail message.
  heap size not invariant: Caller must free returned string.
  --------------------------------------------------------------------*/
  char* getBody();

  /*--------------------------------------------------------------------
  hasRtf
    returns: true, if an compressed RTF (representing an HTML message) 
             stream is present.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL Msg::hasRtf();

  /*--------------------------------------------------------------------
  getRecipientCount 
    returns: the number of contained Recipient Storage objects
  heap size invariant.
  --------------------------------------------------------------------*/
  int getRecipientCount();

  /*--------------------------------------------------------------------
  getRecipientName
    iRecipient: index of contained Recipient Storage object
    returns name of one contained Recipient Storage object.
  heap size invariant.
  --------------------------------------------------------------------*/
  LPOLESTR getRecipientName(int iRecipient);

  /*--------------------------------------------------------------------
  getRecipient
    pwcsRecipient: name of contained Recipient Storage object
    returns: open Storage object of that name or NULL.
  not heap size invariant: Caller must free returned object.
  --------------------------------------------------------------------*/
  Recipient* getRecipient(LPOLESTR pwcsRecipient);

  /*--------------------------------------------------------------------
  getAttachmentCount 
    returns: the number of contained Attachment Storage objects
  heap size invariant.
  --------------------------------------------------------------------*/
  int getAttachmentCount();

  /*--------------------------------------------------------------------
  getAttachmentName
    iAttachment: index of contained Attachment Storage object
    returns name of one contained Attachment Storage object.
  heap size invariant.
  --------------------------------------------------------------------*/
  LPOLESTR getAttachmentName(int iAttachment);

  /*--------------------------------------------------------------------
  getAttachment
    pwcsAttachment: name of contained Attachment Storage object
    returns: open Storage object of that name or NULL.
  not heap size invariant: Caller must free returned object.
  --------------------------------------------------------------------*/
  Attachment* getAttachment(LPOLESTR pwcsAttachment);

  /*--------------------------------------------------------------------
  save saves the content of the message as a .txt file in 
    "RFC 822 format" in ISO-8859-1.
    If an HTML version is present, it will be stored too,
    as well as all attachments and embedded messages. (The latter
    will not be converted recursively.)
    bListFiles: if true, filenames that would be created are listed.
                 message content is not stored.
    pFilename: full name of .txt file to be created.
    pAttachmentDir: name of folder where attachments are to be stored.
    returns: true, unless an error occurred.
  heap size not invariant: System allocates irrecoverable space on heap
                           for stream.
  --------------------------------------------------------------------*/
  BOOL save(BOOL bListFiles, char* pFilename, char* pAttachmentDir);

/*====================================================================*/
private:
  /*--------------------------------------------------------------------
  saveFrom writes a "From:" header to the text file.
    returns: true, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL saveFrom(FILE* pfile);

  /*--------------------------------------------------------------------
  saveToCcBcc writes the "To:", "Cc:", "Bcc:" headers to the text file.
    returns: true, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL saveToCcBcc(FILE* pfile);

  /*--------------------------------------------------------------------
  saveDate writes the "Date:" header to the text file.
    returns: true, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL saveDate(FILE* pfile);

  /*--------------------------------------------------------------------
  saveSubject writes the "Subject:" header to the text file.
    returns: true, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL saveSubject(FILE* pfile);

  /*--------------------------------------------------------------------
  saveBody writes the message body to the text file.
    returns: true, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL saveBody(FILE* pfile);

  /*--------------------------------------------------------------------
  saveHeader writes the message header and the separator line to the text file.
    returns: true, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL saveHeader(FILE* pfile);

  /*--------------------------------------------------------------------
  saveText writes the message to the text file in RFC-882 format
  (ISO-8859-1 / header lines + empty line + body).
    returns: true, if no error occurred.
    pFilename: full name of .txt file to be saved.
  heap size not invariant: System allocates irrecoverable space on heap
                           for stream.
  --------------------------------------------------------------------*/
  BOOL saveText(char* pFilename);

  /*--------------------------------------------------------------------
  saveRtf decompresses the RTF stream, converts it to HTML if
          possible and stored result in AttachmentDir with .rtf or .html
          extension.
    bListFiles: if true, filenames that would be created are listed.
                 message content is not stored.
    pFilename: file name of .txt file to which this RTF/HTML file belongs.
    pAttachmentDir: name of folder where RTF/HTML is to be stored.
    returns: true, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL saveRtf(BOOL bListFiles, char* pFilename, char* pAttachmentDir);

  /*--------------------------------------------------------------------
  saveEmbeddedMsg saves the attachment representing an embedded Msg
    as a .msg file.
    pAttachment: attachment object holding Msg to be saved.
    pFilename: full name of file to be created.
    returns: true, unless an error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL saveEmbeddedMsg(Attachment* pAttachment,char* pFilename);

  /*--------------------------------------------------------------------
  copyTo copies the contents of this embedded Msg to the open Storage
    object.
    pStorage: destination storage for copy of this embedded Msg.
    returns: S_OK, unless an error occurred.
  not heap size invariant: Caller must free returned object.
  --------------------------------------------------------------------*/
  HRESULT copyTo(Storage* pStorage);

  /*--------------------------------------------------------------------
  constructor 
    pStorage: open IStorage object.
  N.B.: resulting Msg must be initialized after construction!
  heap size invariant.
  --------------------------------------------------------------------*/
  Msg(IStorage* pStorage, BOOL bUnicode);

  /*--------------------------------------------------------------------
  initialize is called by open() and recursively by itself.
  It reads all direct children and fills the name arrays.
    returns: 0, if no error occurred.
  not heap size invariant: arrays are only released in destructor.
  --------------------------------------------------------------------*/
  virtual HRESULT initialize();

  /*--------------------------------------------------------------------
  header analyzes header at pHeader and returns pointer to first property.
  heap size invariant.
  --------------------------------------------------------------------*/
  virtual LPBYTE header(LPBYTE pHeader);

  /*--------------------------------------------------------------------
  clear clears the contained elements.
  heap size invariant.
  --------------------------------------------------------------------*/
  void clear();

  /* embedded msg object */
  BOOL m_bEmbedded;
  /* list of recipient child storage objects */
  LPOLESTR* m_apwcsRecipient;
  /* size of array of recipient child storage objects */
  int m_iRecipientCount;
  /* list of attachment child storage objects */
  LPOLESTR* m_apwcsAttachment;
  /* size of array of attachmane child storage objects */
  int m_iAttachmentCount;

}; /* class Msg */

#endif // MSG_H not defined
/*==End of File=======================================================*/
