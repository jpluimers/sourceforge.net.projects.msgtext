/*-- $Workfile: MsgText.cpp $ ------------------------------------------
This program converts a MSG file to a text file and a folder with attachments.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows, Extended MAPI
Description: converts a MSG file to a text file and a folder with attachments.
             This makes use of the Windows platform and the OLE interfaces
             but needs no Outlook installed.
             http://msdn.microsoft.com/en-us/library/cc463912.aspx
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
#define STRICT
#include <stdio.h>
#define USES_CLSID_MailMessage
#define USES_PS_MAPI
#define USES_PS_PUBLIC_STRINGS
#define INITGUID
#include <initguid.h>
#include <shlobj.h>
#include "Utils/Heap.h"
#include "Utils/CmdLine.h"
#include "Utils/Fs.h"
#include "Msg/Msg.h"

static int iRESULT_OK = 0;
static int iRESULT_WARN = 4;
static int iRESULT_ERROR = 8;
static int iRESULT_SEVERE = 12;

#define PROGRAM "MsgText 2.03" /* 1.x was the C# version */ 
#define BANNER "Convert .msg files to .txt files and attachments"
#define COPYRIGHT "(c) 2009, 2010 Enter AG, Zurich, Switzerland\n\
This program comes with ABSOLUTELY NO WARRANTY; It is free software,\n\
and you are welcome to redistribute it under certain conditions;\n\
Use option /c for details.\n"
  

/*----------------------------------------------------------------------
getUsersDocumentsFolder
----------------------------------------------------------------------*/
int getUsersDocumentsFolder(char** psUsersDocumentsFolder)
{
  int iResult = iRESULT_ERROR;
  *psUsersDocumentsFolder = NULL;
  wchar_t* wsUsersDocumentsFolder = new wchar_t[MAX_PATH];
  DWORD dwError = SHGetFolderPath(NULL,CSIDL_PERSONAL,NULL,SHGFP_TYPE_CURRENT,wsUsersDocumentsFolder);
  if (dwError == S_OK)
  {
    iResult = iRESULT_OK;
    *psUsersDocumentsFolder = newMultiByte(wsUsersDocumentsFolder);
    delete wsUsersDocumentsFolder;
  }
  return iResult;
} /* getUsersDocumentsFolder */

/*----------------------------------------------------------------------
displayCopyright
----------------------------------------------------------------------*/
void displayCopyright()
{
  Fs* pfsCopyright = new Fs("COPYING.txt");
  if (!pfsCopyright->exists())
  {
    delete pfsCopyright;
    pfsCopyright = new Fs("../COPYING.txt");
  }
  if (pfsCopyright->exists())
  {
    char* pCopyright = pfsCopyright->name();
    FILE* pfileCopy = fopen(pCopyright,"rt");
    free(pCopyright);
    char* pBuffer = (char*)malloc(BUFSIZ);
    if (pfileCopy != NULL)
    {
      while (fgets(pBuffer,BUFSIZ-1,pfileCopy) != NULL)
        printf("%s\n",pBuffer);
      fclose(pfileCopy);
    }
  }
  else
  {
    printf("Unfortunately someone removed the file COPYING.txt from the distribution\n");
    printf("Please find the GPL notice under http://www.gnu.org/licenses/\n");
  }
  delete pfsCopyright;
} /* displayCopyright */

/*----------------------------------------------------------------------
displayHelp
----------------------------------------------------------------------*/
void displayHelp()
{
  printf("Usage:\n");
  printf("  MsgText [/h] | [/c] | [/l] [/d] <msg file> [<text file> [<attachment dir>]]\n");
  printf("with\n");
  printf("/h               displays this usage note\n");
  printf("/c               displays the copyright note\n");
  printf("/l               list all files that would be created\n");
  printf("                 (msg file is not converted)\n");
  printf("/d               overwrite <textfile> and\n");
  printf("                 delete previous contents of <attachmentdir>\n");
  printf("<msg file>       .msg file to be converted\n");
  printf("<text file>      .txt file with message to be created\n");
  printf("                 (default: same as .msg file with .txt extension)\n");
  printf("<attachment dir> folder for attachments to be created, if needed\n");
  printf("                 (default: same as .txt file without extension)\n");
  printf("\n");
} /* displayHelp */

/*----------------------------------------------------------------------
displayParameters
----------------------------------------------------------------------*/
void displayParameters(BOOL bListFiles, BOOL bDelete,
  char* sMsgFile, char* sTextFile, char* sAttachmentDir)
{
  if (bListFiles)
    printf("files that would be created will be listed\n");
  printf(".msg file to be converted : %s\n",sMsgFile);
  printf(".txt file to be created   : %s\n",sTextFile);
  printf("attachment directory      : %s\n",sAttachmentDir);
  if (!bListFiles && bDelete)
    printf(".txt file and previous contents of attachment directory will be deleted\n");
  printf("\n");
} /* displayParameters */

/*----------------------------------------------------------------------
validateParameters
----------------------------------------------------------------------*/
int validateParameters(BOOL bListFiles, BOOL bDelete,
  char* sMsgFile, char* sTextFile, char* sAttachmentDir)
{
  int iResult = iRESULT_ERROR;
  DWORD dwError = 0;
  /* msg file must exist */
  Fs* pFsMsg = new Fs(sMsgFile);
  if (pFsMsg->exists())
  {
    if (!bListFiles)
    {
      Fs* pFsText = new Fs(sTextFile);
      if (bDelete)
        dwError = pFsText->deleteFile();
      if (!pFsText->exists())
      {
        Fs* pFsAttachment = new Fs(sAttachmentDir);
        if (bDelete)
          dwError = pFsAttachment->deleteFolder(TRUE);
        if (!pFsAttachment->exists())
          iResult = iRESULT_OK;
        else
          printf("Attachment directory \"%s\" exists already!\n",sAttachmentDir);
        delete pFsAttachment;
      }
      else
        printf("Text file \"%s\" exists already!\n",sTextFile);
      delete pFsText;
    }
    else
      iResult = iRESULT_OK;
  }
  else
    printf("Message file \"%s\" does not exist!\n",sMsgFile);
  delete pFsMsg;
  printf("\n");
  return iResult;
} /* validateParameters */

/*----------------------------------------------------------------------
getParameters
----------------------------------------------------------------------*/
int getParameters(
  int argc, char** args, BOOL* pbListFiles,
  char** psMsgFile, char** psTextFile, char** psAttachmentDir)
{
  int iResult = iRESULT_ERROR;
  unsigned uArg;
  unsigned uUnnamed = 0;
  *pbListFiles = FALSE;
  BOOL bDelete = FALSE;
  *psTextFile = NULL;
  *psAttachmentDir = NULL;
  CmdLine* pcl = new CmdLine(argc,args);
  if (pcl->error() == NULL)
  {
    iResult = iRESULT_OK;
    for (uArg = 0; uArg < pcl->arguments(); uArg++)
    {
      if (pcl->name(uArg) == NULL)
      {
        switch(uUnnamed)
        {
        case 0:
          *psMsgFile = _strdup(pcl->value(uArg));
          break;
        case 1:
          *psTextFile = _strdup(pcl->value(uArg));
          break;
        case 2:
          *psAttachmentDir = (char*)malloc(strlen(pcl->value(uArg))+2);
          strcpy(*psAttachmentDir,pcl->value(uArg));
          if ((*psAttachmentDir)[strlen(*psAttachmentDir)-1] != '\\')
            strcat(*psAttachmentDir,"\\");
          break;
        default:
          printf("Invalid extra argument \"%s\" on command line!\n",pcl->value(uArg));
          displayHelp();
          iResult = iRESULT_ERROR;
          break;
        }
        uUnnamed++;
      }
      else if (strcmp(pcl->name(uArg),"h") == 0)
      {
        displayHelp();
        iResult = iRESULT_WARN;
      }
      else if (strcmp(pcl->name(uArg),"c") == 0)
      {
        displayCopyright();
        iResult = iRESULT_WARN;
      }
      else if (strcmp(pcl->name(uArg),"l") == 0)
        *pbListFiles = TRUE;
      else if (strcmp(pcl->name(uArg),"d") == 0)
        bDelete = TRUE;
      else
      {
        printf("Invalid option \"%s\" on command line!\n",pcl->name(uArg));
        displayHelp();
        iResult = iRESULT_ERROR;
      }
    }
    if (iResult == iRESULT_OK)
    {
      if (*psMsgFile == NULL)
      {
        displayHelp();
        printf("No .msg file given!\n");
        iResult = iRESULT_ERROR;
      }
    }
    if (iResult == iRESULT_OK)
    {
      /* default text file */
      if (*psTextFile == NULL)
      {
        Fs* pFsMsg = new Fs(*psMsgFile);
        *psTextFile = pFsMsg->deriveFile(".txt");
        delete pFsMsg;
      }
      /* default attachment dir */
      if (*psAttachmentDir == NULL)
      {
        Fs* pFsText = new Fs(*psTextFile);
        *psAttachmentDir = pFsText->deriveFile("\\");
        delete pFsText;
      }
      iResult = validateParameters(*pbListFiles, bDelete, 
        *psMsgFile, *psTextFile, *psAttachmentDir);
    }
    if (iResult == iRESULT_OK)
      displayParameters(*pbListFiles, bDelete, 
        *psMsgFile, *psTextFile, *psAttachmentDir);
  }
  else
    printf("Command line parsing error: %s\n",pcl->error());
  delete pcl;
  return iResult;
} /* getParameters */

/*----------------------------------------------------------------------
main
----------------------------------------------------------------------*/
int main(int argc, char** args)
{
  printf("%s; %s\n%s\n\n",PROGRAM,BANNER,COPYRIGHT);
  /* default arguments */
  BOOL bListFiles = FALSE;
  char* sMsgFile = NULL;
  char* sTextFile = NULL;
  char* sAttachmentDir = NULL;
  int iResult = getParameters(argc, args, &bListFiles, &sMsgFile, &sTextFile, &sAttachmentDir);
  if (iResult == iRESULT_OK)
  {
    iResult = iRESULT_ERROR;
    Msg* pMsg = new Msg(sMsgFile);
    HRESULT hr = pMsg->open();
    if (SUCCEEDED(hr))
      if (pMsg->save(bListFiles,sTextFile,sAttachmentDir))
        iResult = iRESULT_OK;
    delete pMsg;
  }
  if ((iResult == iRESULT_OK) || (iResult == iRESULT_WARN))
    printf("\n%s terminates successfully.\n",PROGRAM);
  else
    printf("\n%s terminates with errors!\n",PROGRAM);
  return iResult;
} /* main */
