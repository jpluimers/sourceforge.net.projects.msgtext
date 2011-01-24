/*-- $Workfile: Folder.h $ ---------------------------------------------
Wrapper for an IMAPIFolder.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: Folder class wraps an IMAPIFolder.
------------------------------------------------------------------------
Copyright  : Enter AG, Zurich, 2009
----------------------------------------------------------------------*/
#ifndef FOLDER_H
#define FOLDER_H

#define STRICT
#include <mapix.h>
#include "Container.h"
#include "Message.h"
#include "Table.h"

/*======================================================================
Folder is the IMAPIFolder wrapper
======================================================================*/
class Folder : public Container
{
public:
  /*--------------------------------------------------------------------
  constructor 
    pMapiFolder: pointer to IMAPIFolder interface.
  heap size not invariant: establishes list of PropValues, list of
                           subfolders and list of messages.
  --------------------------------------------------------------------*/
  Folder(IMAPIFolder* pMapiFolder);

  /*--------------------------------------------------------------------
  destructor releases IMAPIFolder interface.
  heap size not invariant: cleans up list of subfolders, list of messages
                           and list of PropValues.
  --------------------------------------------------------------------*/
  ~Folder();

  /*--------------------------------------------------------------------
  getDisplayName returns the display name of the MAPI folder.
    returns: display name of the MAPI folder.
  heap size not invariant: caller must delete display name.
  --------------------------------------------------------------------*/
  char* getDisplayName();

  /*--------------------------------------------------------------------
  folders
    returns: number of MAPI subfolders of this MAPI folder.
  heap size invariant.
  --------------------------------------------------------------------*/
  ULONG folders();

  /*--------------------------------------------------------------------
  getFolder return MAPI subfolder of given index.
    ulIndex: index of MAPI subfolder to return.
    returns: the MAPI subfolder with given index or NULL, if an error occurred.
  heap size not invariant: caller must delete returned object.
  --------------------------------------------------------------------*/
  Folder* getFolder(ULONG ulIndex);

  /*--------------------------------------------------------------------
  messages
    returns: number of MAPI message objects (not only of type "IPM.Note"!)
             of this MAPI folder.
  heap size invariant.
  --------------------------------------------------------------------*/
  ULONG messages();

  /*--------------------------------------------------------------------
  getMessage return MAPI message object of given index.
    ulIndex: index of MAPI message object to return.
    returns: the MAPI message object with given index or NULL, if an error occurred.
  heap size not invariant: caller must delete returned object.
  --------------------------------------------------------------------*/
  Message* getMessage(ULONG ulIndex);

  /*--------------------------------------------------------------------
  isEmpty returns TRUE, if folder does not contain any message objects
    of type "IPM.Note" or any non-empty subfolders.
  heap size invariant.
  --------------------------------------------------------------------*/
  BOOL isEmpty();

  /*--------------------------------------------------------------------
  list write a line with the folder's name and statistics to stdout.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT list();

  /*--------------------------------------------------------------------
  listMapi lists all MAPI folders unter the given one.
    bEmpty      : true, if empty folders are to be listed.
    pMapiFolder : only folders under this folder are to be listed.
    pMapiPath   : prefix to be used for listing and identifying MAPI subfolders.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT listMapi(BOOL bEmpty, char* pMapiFolder, char* pMapiPath);

  /*--------------------------------------------------------------------
  listFiles lists all MSG files that would be created in archive.
    bEmpty        : true, if empty folders are to be listed.
    pMapiFolder   : Only mails in this folder are to be listed.
    pMapiPath     : prefix to be used for identifying MAPI folder.
    ulMaxFilenameLength: maximum folder/file name length
    pFolder       : directory, where folder contents would be archived.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT listFiles(BOOL bEmpty, char* pMapiFolder, char* pMapiPath,
    ULONG ulMaxFilenameLength, char* pFolderPath);

  /*--------------------------------------------------------------------
  archive write the folder's messages and subfolders to the given
    directory.
    The file name is based on the display name of the folder.
    This is normalized and truncated to the maximum file name length
    as well as disambiguated.
    bEmpty        : true, if empty folders are to be listed.
    pMapiFolder   : Only mails in this folder are to be listed.
    pMapiPath     : prefix to be used for identifying MAPI folder.
    ulMaxFilenameLength: maximum folder/file name length
    pFolder       : directory, where folder contents would be archived.
    returns: 0, if no error occurred.
  heap size invariant.
  --------------------------------------------------------------------*/
  HRESULT archive(BOOL bEmpty, char* pMapiFolder, char* pMapiPath,
    ULONG ulMaxFilenameLength, char* pFolderPath);

private:
  /* table of MAPI subfolders of this MAPI folder */
  Table* m_pTableFolders;
  /* table of MAPI message objects of this MAPI folder */
  Table* m_pTableMessages;
  /* maximum file name length for archive */
  ULONG m_ulMaxFilenameLength;
}; /* class Folder */

#endif // FOLDER_H not defined
/*==End of File=======================================================*/
