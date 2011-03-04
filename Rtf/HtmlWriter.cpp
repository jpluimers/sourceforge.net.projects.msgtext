/*-- $Workfile: HtmlWriter.cpp $ ---------------------------------------
HtmlWriter parses RTF and ignores, copies or translates it.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: HtmlWriter parses RTF and ignores it, if it was generated 
             from text, translates it to HTML, if it was generated from
             HTML, or writes it unchanged.
             See MS-OXRTFEX.pdf (http://msdn.microsoft.com/en-us/library/cc425505.aspx)
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
#define STRICT
#include <windows.h>
#include "../Utils/Heap.h"
#include "../Utils/Fs.h"
#include "HtmlWriter.h"

#define PARSESTATE_GROUP_START  1
#define PARSESTATE_GROUP_END  2
#define PARSESTATE_CONTROL_SYMBOL  3
#define PARSESTATE_CONTROL_WORD  4
#define PARSESTATE_TEXT  5
#define PARSESTATE_UNFINISHED  6

/*--------------------------------------------------------------------*/
HtmlWriter::HtmlWriter(BOOL bListFiles, char* pAttachmentDir, char*pFilename)
{
  m_bListFiles = bListFiles;
  m_pAttachmentDir = _strdup(pAttachmentDir);
  m_pFilename = _strdup(pFilename);
  m_pUntranslated = NULL;
  m_ulUntranslated = 0;
  m_pUnwritten = NULL;
  m_ulUnwritten = 0;
  m_pfile = NULL;
  m_bTranslate = false;
  m_bSuppress = false;
  m_ulps = PARSESTATE_TEXT;
  m_pKeyword = NULL;
  m_pText = NULL;
  m_pParm = NULL;
  m_ulParm = 0;
  m_pRtfName = NULL;
  m_ulGroups = 0;
  m_bHtml = false;
  m_bHtmlTag = false;
  m_cSymbol = ' ';
  m_bInBody = false;
  m_bInHead = false;
  m_bInHtml = false;
  m_bHasHead = false;
  m_bInComment = false;
} /* HtmlWriter */

/*--------------------------------------------------------------------*/
HtmlWriter::~HtmlWriter()
{
  if (m_pAttachmentDir)
    free(m_pAttachmentDir);
  if (m_pFilename)
    free(m_pFilename);
  if (m_pUntranslated != NULL)
    free(m_pUntranslated);
  if (m_pUnwritten != NULL)
    free(m_pUnwritten);
  if (m_pKeyword != NULL)
    free(m_pKeyword);
  if (m_pText != NULL)
    free(m_pText);
  if (m_pParm != NULL)
    free(m_pParm);
  if (m_pRtfName != NULL)
    free(m_pRtfName);
} /* ~Crc */

/*--------------------------------------------------------------------*/
BYTE HtmlWriter::hexNybble(char* p)
{
  BYTE b = 0;
  if ((*p >= '0') && (*p <= '9'))
    b = (BYTE)(*p) - (BYTE)'0';
  else if ((*p >= 'A') && (*p <= 'F'))
    b = 10 + (BYTE)(*p) - (BYTE)'A';
  else if ((*p >= 'a') && (*p <= 'f'))
    b = 10 + (BYTE)(*p) - (BYTE)'a';
  return b;
} /* hexNybble */

/*--------------------------------------------------------------------*/
char HtmlWriter::hexChar(char* p)
{
  BYTE b = (hexNybble(p) << 4) + hexNybble(p+1);
  return (char)b;
} /* hexChar */

/*--------------------------------------------------------------------*/
int HtmlWriter::index(ULONG ulOffset, char c)
{
  int iIndex;
  for (iIndex = (int)ulOffset;
       (iIndex < (int)m_ulUntranslated) && (m_pUntranslated[iIndex] != c);
       iIndex++)
  {}
  if (iIndex == m_ulUntranslated)
    iIndex = -1;
  return iIndex;
} /* index */

/*--------------------------------------------------------------------*/
char* HtmlWriter::append(char* p, char c)
{
  ULONG ulLength = 0;
  if (p != NULL)
    ulLength = strlen(p);
  ulLength++; /* new string length */
  p = (char*)realloc(p,ulLength+1); /* space for '\0' */
  p[ulLength-1] = c;
  p[ulLength] = '\0';
  return p;
} /* append */

/*--------------------------------------------------------------------*/
ULONG HtmlWriter::parseControl(ULONG ulOffset)
{
  if (m_pKeyword != NULL)
  {
    free(m_pKeyword);
    m_pKeyword = NULL;
  }
  if (m_pParm != NULL)
  {
    free(m_pParm);
    m_pParm = NULL;
  }
  m_ulps = PARSESTATE_UNFINISHED;
  ULONG ul = ulOffset + 1; // skip '\'
  if (ul < m_ulUntranslated)
  {
    if (islower(m_pUntranslated[ul]))
    {
      /* parseControlWord */
      for (; (ul < m_ulUntranslated) && islower(m_pUntranslated[ul]); ul++)
        m_pKeyword = append(m_pKeyword,m_pUntranslated[ul]);
      if (ul < m_ulUntranslated)
      {
        m_ulps = PARSESTATE_CONTROL_WORD;
        if (m_pUntranslated[ul] == ' ')
          ul++;
        else if (isdigit(m_pUntranslated[ul]) || (m_pUntranslated[ul] == '-'))
        {
          m_ulps = PARSESTATE_UNFINISHED;
          m_pParm = append(m_pParm,m_pUntranslated[ul]);
          ul++;
          for (; (ul < m_ulUntranslated) && (isdigit(m_pUntranslated[ul])); ul++)
            m_pParm = append(m_pParm,m_pUntranslated[ul]);
          if (ul < m_ulUntranslated)
          {
            m_ulps = PARSESTATE_CONTROL_WORD;
            if (m_pUntranslated[ul] == ' ')
              ul++;
          }
        }
      }
    }
    else
    {
      m_ulps = PARSESTATE_CONTROL_SYMBOL;
      m_cSymbol = m_pUntranslated[ul];
      ul++;
    }
  }
  if (m_ulps != PARSESTATE_UNFINISHED)
    ulOffset = ul;
  return ulOffset;
} /* parseControl */

/*--------------------------------------------------------------------*/
ULONG HtmlWriter::parseText(ULONG ulOffset)
{
  if (m_pText != NULL)
  {
    free(m_pText);
    m_pText = NULL;
  }
  int igs = index(ulOffset,'{');
  int ige = index(ulOffset,'}');
  int ictl = index(ulOffset,'\\');
  int icr = index(ulOffset,'\r');
  int ilf = index(ulOffset,'\n');
  int iEnd = (int)m_ulUntranslated;
  if ((igs >= 0) && (igs < iEnd))
    iEnd = igs;
  if ((ige >= 0) && (ige < iEnd))
    iEnd = ige;
  if ((ictl >= 0) && (ictl < iEnd))
    iEnd = ictl;
  if ((icr >= 0) && (icr < iEnd))
    iEnd = icr;
  if ((ilf >= 0) && (ilf < iEnd))
    iEnd = ilf;
  m_ulps = PARSESTATE_TEXT;
  m_pText = (char*)malloc(iEnd - ulOffset + 1);
  memcpy(m_pText,&m_pUntranslated[ulOffset],iEnd - ulOffset);
  m_pText[iEnd - ulOffset] = '\0';
  ulOffset = iEnd;
  return ulOffset;
} /* parseText */

/*--------------------------------------------------------------------*/
ULONG HtmlWriter::parseGroupStart(ULONG ulOffset)
{
  ulOffset++;
  m_ulps = PARSESTATE_GROUP_START;
  m_ulGroups++;
  return ulOffset;
} /* parseGroupStart */

/*--------------------------------------------------------------------*/
ULONG HtmlWriter::parseGroupEnd(ULONG ulOffset)
{
  ulOffset++;
  m_ulps = PARSESTATE_GROUP_END;
  m_ulGroups--;
  return ulOffset;
} /* parseGroupEnd */

/*--------------------------------------------------------------------*/
ULONG HtmlWriter::parseIgnore(ULONG ulOffset)
{
  m_ulps = PARSESTATE_UNFINISHED;
  while ((ulOffset < m_ulUntranslated) &&
       ((m_pUntranslated[ulOffset] == '\r') || (m_pUntranslated[ulOffset] == '\n')))
    ulOffset++;
  return ulOffset;
} /* parseIgnore */

/*--------------------------------------------------------------------*/
ULONG HtmlWriter::parse(ULONG ulOffset)
{
  if (m_ulps != PARSESTATE_UNFINISHED)
  {
    ulOffset = parseIgnore(ulOffset);
    if (ulOffset < m_ulUntranslated)
    {
      switch (m_pUntranslated[ulOffset])
      {
        case '{': ulOffset = parseGroupStart(ulOffset); break;
        case '}': ulOffset = parseGroupEnd(ulOffset); break;
        case '\\': ulOffset = parseControl(ulOffset); break;
        default: ulOffset = parseText(ulOffset); break;
      }
    }
  }
  return ulOffset;
} /* parse */

/*--------------------------------------------------------------------*/
char* HtmlWriter::replaceUrl(ULONG* pulOffset, char* pText)
{
  ULONG ul = *pulOffset;
  ULONG ulLength = strlen(pText);
  char* pMatch = strstr(&pText[ul],"\"cid:");
  if (pMatch != NULL)
  {
    pMatch++; /* keep quote */
    /* search for "cid:<filename>@<pr_content_id>" ... */
    char* pAt = strstr(pMatch,"@");
    char* pEnd = strstr(pMatch,"\"");
    if (pEnd != NULL)
    {
      if ((pAt != NULL) && (pAt < pEnd))
      {
        /* ... and replace it by <filename> */
        *pAt = '\0';
        char* pFilename = &pMatch[strlen("cid:")];
        ULONG ulFilenameLength = strlen(pFilename);
        memcpy(pMatch,pFilename,ulFilenameLength);
        pMatch += ulFilenameLength;
        strcpy(pMatch,pEnd);
        ul = pMatch - pText;
      }
      else
        ul = pEnd - pText;
    }
    else
      ul = strlen(pText);
  }
  else
    ul = strlen(pText);
  *pulOffset = ul;
  return pText;
} /* replaceUrl */

/*--------------------------------------------------------------------*/
char* HtmlWriter::replaceUrls(char* p)
{
  ULONG ulLength = strlen(p);
  for (ULONG ulOffset = 0; ulOffset < ulLength; )
  {
    p = replaceUrl(&ulOffset,p);
    ulLength = strlen(p);
  }
  return p;
} /* replaceUrls */

/*--------------------------------------------------------------------*/
void HtmlWriter::appendUnwritten()
{
  if (m_pText != NULL)
  {
    m_pUnwritten = (char*)realloc(m_pUnwritten, m_ulUnwritten+strlen(m_pText));
    memcpy(&m_pUnwritten[m_ulUnwritten],m_pText,strlen(m_pText));
    m_ulUnwritten += strlen(m_pText);
    free(m_pText);
    m_pText = NULL;
  }
} /* appendUnwritten */

/*--------------------------------------------------------------------*/
ULONG HtmlWriter::translateControlSymbol(ULONG ulOffset)
{
  if (m_bHtml || m_bHtmlTag)
  {
    if (m_cSymbol == '\'')
    {
      if (ulOffset < m_ulUntranslated - 1)
      {
        m_pText = append(m_pText,hexChar(&m_pUntranslated[ulOffset]));
        ulOffset +=  2;
      }
      else
      {
        ulOffset -= 2;
        m_ulps = PARSESTATE_UNFINISHED;
      }
    }
    else if ((m_cSymbol == '\\') || (m_cSymbol == '{') || (m_cSymbol == '}'))
      m_pText = append(m_pText,m_cSymbol);
    else if (m_cSymbol == '~')
      m_pText = _strdup("&nbsp;");
    else if (m_cSymbol == '_')
      m_pText = _strdup("&shy;");
    appendUnwritten();
  }
  return ulOffset;
} /* translateControlSymbol */

/*--------------------------------------------------------------------*/
void HtmlWriter::translateHtmlTag()
{
  if (m_pParm != NULL)
  {
    ULONG ulParm = atol(m_pParm);
    ULONG ulDestination = ulParm & 0x0003;  /* 0: INBODY, 1: INHEAD, 2: INHTML, 3: OUTHTML */
    ULONG ulTagType = ulParm & 0x00F0;
    ULONG ulInparagraph = ulParm & 0x0004;
    ULONG ulClose = ulParm & 0x0008;
    ULONG ulMhtml = ulParm & 0x0100; // rewritable URL parameters
    if (ulClose == 0)
    {
      if (ulDestination < 3)
      {
        if (!m_bInHtml)
        {
          m_pText = _strdup("<html>\r\n");
          appendUnwritten();
          m_bInHtml = true;
        }
        if (ulDestination < 2)
        {
          if (ulDestination == 1)
          {
            if (!m_bInHead)
            {
              m_pText = _strdup("<head>\r\n");
              appendUnwritten();
              m_bInHead = true;
              m_bHasHead = true;
            }
          }
          else if (ulDestination == 0)
          {
            if (!m_bHasHead)
            {
              m_pText = _strdup("<head>\r\n");
              appendUnwritten();
              m_bInHead = true;
              m_bHasHead = true;
            }
            if (m_bInHead)
            {
              m_pText = _strdup("</head>\r\n");
              appendUnwritten();
              m_bInHead = false;
            }
            if (!m_bInBody)
            {
              m_pText = _strdup("<body>\r\n");
              appendUnwritten();
              m_bInBody = true;
            }
          }
        }
        else
        {
          if (m_bInHead)
          {
            m_pText = _strdup("</head>\r\n");
            appendUnwritten();
            m_bInHead = false;
          }
          if (m_bInBody)
          {
            m_pText = _strdup("</body>\r\n");
            appendUnwritten();
            m_bInBody = false;
          }
        }
      }
      else
      {
        if (m_bInHead)
        {
          m_pText = _strdup("</head>\r\n");
          appendUnwritten();
          m_bInHead = false;
          m_pText = _strdup("<body>\r\n");
          appendUnwritten();
          m_bInBody = true;
        }
        if (m_bInBody)
        {
          m_pText = _strdup("</body>\r\n");
          appendUnwritten();
          m_bInBody = false;
        }
        if (m_bInHtml)
        {
          m_pText = _strdup("</html>\r\n");
          appendUnwritten();
          m_bInHtml = false;
        }
      }
    }
    if (ulTagType == 0) /* non-printing HTML between tags */
    {
      m_bHtmlTag = false;
      m_bHtml = true;
    }
    else
    {
      m_bHtmlTag = true;
      m_bHtml = false;
    }
  }
} /* translateHtmlTag */

/*--------------------------------------------------------------------*/
void HtmlWriter::translateControlWord()
{
  if (m_bHtml || m_bHtmlTag)
  {
    if (strcmp(m_pKeyword, "par") == 0)
      m_pText = _strdup("\r\n");
    else if (strcmp(m_pKeyword, "tab") == 0)
      m_pText = _strdup("\t");
    else if (strcmp(m_pKeyword, "lquote") == 0)
      m_pText = _strdup("&lsquo;");
    else if (strcmp(m_pKeyword, "rquote") == 0)
      m_pText = _strdup("&rsquo;");
    else if (strcmp(m_pKeyword, "ldblquote") == 0)
      m_pText = _strdup("&ldquo;");
    else if (strcmp(m_pKeyword, "rdblquote") == 0)
      m_pText = _strdup("&rdquo;");
    else if (strcmp(m_pKeyword, "bullet") == 0)
      m_pText = _strdup("&bull;");
    else if (strcmp(m_pKeyword, "endash") == 0)
      m_pText = _strdup("&ndash;");
    else if (strcmp(m_pKeyword, "emdash") == 0)
      m_pText = _strdup("&mdash;");
    else if (m_pKeyword[0] == 'u') // Unicode
    {
      ULONG ulCode = atol(&m_pKeyword[1]);
      m_pText = _strdup("&#xHHHH;");
      sprintf(m_pText,"&#x%04x;",ulCode);
    }
    appendUnwritten();
  }
} /* translateControlWord */

/*--------------------------------------------------------------------*/
void HtmlWriter::translate()
{
  m_ulps = PARSESTATE_TEXT;
  ULONG ulOffset = 0;
  for (ulOffset = parse(ulOffset); m_ulps != PARSESTATE_UNFINISHED; ulOffset = parse(ulOffset))
  {
    switch (m_ulps)
    {
      case PARSESTATE_GROUP_END:
        m_bHtmlTag = false;
        break;
      case PARSESTATE_GROUP_START:
        break;
      case PARSESTATE_CONTROL_SYMBOL:
        if (m_pText != NULL)
        {
          free(m_pText);
          m_pText = NULL;
        }
        ulOffset = translateControlSymbol(ulOffset);
        break;
      case PARSESTATE_CONTROL_WORD:
        if (m_pText != NULL)
        {
          free(m_pText);
          m_pText = NULL;
        }
        if (strcmp(m_pKeyword, "htmltag") == 0)
          translateHtmlTag();
        else if (strcmp(m_pKeyword, "mhtmltag") == 0)
          m_bHtml = false; // is to be ignored!
        else if (strcmp(m_pKeyword, "htmlrtf") == 0)
        {
          if ((m_pParm != NULL) && (strcmp(m_pParm, "0") == 0))
            m_bHtml = true;
          else
            m_bHtml = false;
        }
        else translateControlWord();
        break;
      case PARSESTATE_TEXT:
        if (m_bHtml)
          appendUnwritten();
        else if (m_bHtmlTag)
        {
          if (strlen(m_pText) >= 6)
          {
            if ((memcmp(m_pText,"<html",5) == 0) || (memcmp(m_pText,"<HTML",5) == 0))
              m_bInHtml = true;
            else if ((memcmp(m_pText,"<head",5) == 0) || (memcmp(m_pText,"<HEAD",5) == 0))
            {
              m_bInHead = true;
              m_bHasHead = true;
            }
            else if ((memcmp(m_pText,"<body",5) == 0) || (memcmp(m_pText,"<BODY",5) == 0))
              m_bInBody = true;
            else if ((memcmp(m_pText,"</html",6) == 0) || (memcmp(m_pText,"</HTML",6) == 0))
              m_bInHtml = false;
            else if ((memcmp(m_pText,"</head",6) == 0) || (memcmp(m_pText,"</HEAD",6) == 0))
              m_bInHead = false;
            else if ((memcmp(m_pText,"</body",6) == 0) || (memcmp(m_pText,"</BODY",6) == 0))
              m_bInBody = false;
          }
          if ((strlen(m_pText) >= 4) && (memcmp(m_pText,"<!--",4) == 0))
            m_bInComment = true;
          if ((strlen(m_pText) >= 3) && (strcmp(&m_pText[strlen(m_pText)-3],"-->") == 0))
            m_bInComment = false;
          if (!m_bInComment)
            m_pText = replaceUrls(m_pText);
          appendUnwritten();
        }
        break;
      default:
        break;
    }
  }
  /* update m_pUntranslated */
  m_ulUntranslated = m_ulUntranslated - ulOffset;
  char* p = (char*)malloc(m_ulUntranslated);
  memcpy(p,&m_pUntranslated[ulOffset],m_ulUntranslated);
  free(m_pUntranslated);
  m_pUntranslated = p;
} /* translate */

/*--------------------------------------------------------------------*/
BOOL HtmlWriter::open()
{
  BOOL bSucceeded = false;
  /* parse until \fromhtml or the first group is encountered */
  ULONG ulOffset = parse(0);
  if (m_ulps == PARSESTATE_GROUP_START)
  {
    for (ulOffset = parse(ulOffset);
         (!m_bTranslate) &&
         (!m_bSuppress) &&
         (ulOffset < m_ulUntranslated) &&
         (m_ulps == PARSESTATE_CONTROL_WORD);
         ulOffset = parse(ulOffset))
    {
      if (strcmp(m_pKeyword,"fromhtml") == 0)
        m_bTranslate = true;
      if (strcmp(m_pKeyword,"fromtext") == 0)
        m_bSuppress = true;
    }
    bSucceeded = true;
  }
  else
    printf("invalid RTF stream!!!\n");
  if (bSucceeded)
  {
    Fs* pfsText = new Fs(m_pFilename);
    char* pExtension = _strdup(".rtf");
    if (m_bTranslate)
    {
      free(pExtension);
      pExtension = _strdup(".html");
    }
    if (m_bSuppress)
    {
      free(pExtension);
      pExtension = _strdup(".txt");
    }
    char* pRtfFilename = pfsText->deriveFile(pExtension);
    free(pExtension);
    delete pfsText;
    Fs* pfsAttachFolder = new Fs(m_pAttachmentDir);
    m_pRtfName = pfsAttachFolder->newFile(pRtfFilename);
    free(pRtfFilename);
    if (!m_bSuppress)
    {
      if (!(m_bListFiles || pfsAttachFolder->exists()))
        pfsAttachFolder->createFolder();
      printf("%s\n",m_pRtfName);
      if (!m_bListFiles)
        m_pfile = fopen(m_pRtfName,"wb");
    }
    delete pfsAttachFolder;
    m_ulGroups = 0; // in any case restart groups count
  }
  return bSucceeded;
} /* open */

/*--------------------------------------------------------------------*/
BOOL HtmlWriter::write()
{
  BOOL bSucceeded = true;
  if (m_bTranslate)
  {
    translate();
    if (fwrite(m_pUnwritten,1,m_ulUnwritten,m_pfile) == m_ulUnwritten)
      m_ulUnwritten = 0;
    else
      bSucceeded = false;
  }
  else
  {
    if (fwrite(m_pUntranslated,1,m_ulUntranslated,m_pfile) == m_ulUntranslated)
      m_ulUntranslated = 0;
    else
      bSucceeded = false;
  }
  return bSucceeded;
} /* write */

/*--------------------------------------------------------------------*/
BOOL HtmlWriter::write(ULONG ulSize, char* pBuffer)
{
  BOOL bSucceeded = true;
  /* append buffer to untranslated RTF */
  m_pUntranslated = (char*)realloc(m_pUntranslated, m_ulUntranslated+ulSize);
  memcpy(&m_pUntranslated[m_ulUntranslated],pBuffer,ulSize);
  m_ulUntranslated += ulSize;
  /* we assume, that the first write holds enough header material to make decisions */
  if (m_pRtfName == NULL)
    bSucceeded = open();
  /* translate a bit and write a bit */
  if (bSucceeded)
  {
    if (!(m_bListFiles || m_bSuppress))
      bSucceeded = write();
  }
  return bSucceeded;
} /* write */

/*--------------------------------------------------------------------*/
BOOL HtmlWriter::close()
{
  BOOL bSucceeded = true;
  if (!(m_bListFiles || m_bSuppress))
  {
    if (m_bTranslate)
    {
      if (m_bInBody)
        fprintf(m_pfile, "</body>\r\n");
      if (m_bInHtml)
        fprintf(m_pfile, "</html>\r\n");
    }
    fclose(m_pfile);
    if (m_ulGroups == 0)
      bSucceeded = true;
    else
      printf("RTF file could not be parsed!\n");
  }
  else
    bSucceeded = true;
  return bSucceeded;
} /* close */

/*==End of File=======================================================*/
