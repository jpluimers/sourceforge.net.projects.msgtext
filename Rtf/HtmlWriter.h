/*-- $Workfile: HtmlWriter.h $ -----------------------------------------
HtmlWriter parses RTF and ignores, copies or translates it.
Version    : $Id$
Application: MSG file archiving
Platform   : Windows
Description: HtmlWriter parses RTF and ignores it, if it was generated 
             from text, translates it to HTML, if it was generated from
             HTML, or writes it unchanged.
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
#ifndef HTMLWRITER_H
#define HTMLWRITER_H

/*======================================================================
HtmlWriter parses RTF and ignores, copies or translates it.
======================================================================*/
class HtmlWriter
{
public:
  /*--------------------------------------------------------------------
  constructor saves copies of pAttachmentDir and pFilename in case
              output is needed later.
  heap size not invariant: Buffers will be freed by destructor.
  --------------------------------------------------------------------*/
  HtmlWriter(BOOL bListFiles, char* pAttachmentDir, char* pFilename);

  /*--------------------------------------------------------------------
  destructor
  heap size invariant.
  --------------------------------------------------------------------*/
  ~HtmlWriter();

  /*--------------------------------------------------------------------
  write writes the buffer.
  --------------------------------------------------------------------*/
  BOOL write(ULONG ulSize, char* pBuffer);

  /*--------------------------------------------------------------------
  close writes the remaining buffer and closes the output file.
  --------------------------------------------------------------------*/
  BOOL close();

/*====================================================================*/
private:

  /*--------------------------------------------------------------------
  open parses RTF, until it finds fromtext or fromhtml and opens
       the file.
  --------------------------------------------------------------------*/
  BOOL open();

  /*--------------------------------------------------------------------
  write translates if necessary and writes Unwritten to file.
  --------------------------------------------------------------------*/
  BOOL write();

  /*--------------------------------------------------------------------
  translate translates from Untranslated to Unwritten.
  --------------------------------------------------------------------*/
  void translate();

  /*--------------------------------------------------------------------
  translateControlSymbol translates control symbols within HTML text.
  --------------------------------------------------------------------*/
  ULONG translateControlSymbol(ULONG ulOffset);

  /*--------------------------------------------------------------------
  translateHtmlTag translates control word "htmltag".
  --------------------------------------------------------------------*/
  void translateHtmlTag();

  /*--------------------------------------------------------------------
  translateControlWord translates control words within HTML text.
  --------------------------------------------------------------------*/
  void translateControlWord();

  /*--------------------------------------------------------------------
  appendUnwritten appends m_pText to m_pUnwritten and clears m_pText.
  --------------------------------------------------------------------*/
  void appendUnwritten();

  /*--------------------------------------------------------------------
  replaceUrls returns pText with URLs replaced.
  --------------------------------------------------------------------*/
  char* replaceUrls(char* p);

  /*--------------------------------------------------------------------
  replaceUrl returns pText with next URL after *pulOffset replaced
  and updates the offset for subsequent URL replacements.
  Searches for "cid:<filename>@<pr_content_id>" and replaces it by
  <filename>.
  --------------------------------------------------------------------*/
  char* replaceUrl(ULONG* pulOffset, char* pText);

  /*--------------------------------------------------------------------
  parse parses RTF from m_pUntranslated[ulOffset] and returns the
    resulting offset.
  --------------------------------------------------------------------*/
  ULONG parse(ULONG ulOffset);

  /*--------------------------------------------------------------------
  parseControl parses a control starting with '\'
  --------------------------------------------------------------------*/
  ULONG parseControl(ULONG ulOffset);

  /*--------------------------------------------------------------------
  parseText parses text
  --------------------------------------------------------------------*/
  ULONG parseText(ULONG ulOffset);

  /*--------------------------------------------------------------------
  parseGroupStart digests a '{'
  --------------------------------------------------------------------*/
  ULONG parseGroupStart(ULONG ulOffset);

  /*--------------------------------------------------------------------
  parseGroupEnd digests a '}'
  --------------------------------------------------------------------*/
  ULONG parseGroupEnd(ULONG ulOffset);

  /*--------------------------------------------------------------------
  parseIgnore skips CR/LF
  --------------------------------------------------------------------*/
  ULONG parseIgnore(ULONG ulOffset);

  /*--------------------------------------------------------------------
  append appends the character to the string
  --------------------------------------------------------------------*/
  char* append(char* p, char c);

  /*--------------------------------------------------------------------
  index returns position of character in m_pUntranslated or -1
        starting at ulOffset.
  --------------------------------------------------------------------*/
  int index(ULONG ulOffset, char c);

  /*--------------------------------------------------------------------
  hexChar returns a single byte computed from the two hex characters at p.
  --------------------------------------------------------------------*/
  char hexChar(char* p);

  /*--------------------------------------------------------------------
  hexChar returns a byte corresponding to the hex digit at p.
  --------------------------------------------------------------------*/
  BYTE hexNybble(char* p);

  char* m_pAttachmentDir;
  char* m_pFilename;
  BOOL m_bListFiles;
  char* m_pUntranslated;
  ULONG m_ulUntranslated;
  char* m_pUnwritten;
  ULONG m_ulUnwritten;
  FILE* m_pfile;
  BOOL m_bTranslate;
  BOOL m_bSuppress;
  ULONG m_ulps;
  char* m_pKeyword;
  char* m_pText;
  ULONG m_ulText;
  char* m_pParm;
  ULONG m_ulParm;
  char* m_pRtfName;
  ULONG m_ulGroups;
  BOOL m_bHtml;
  BOOL m_bHtmlTag;
  char m_cSymbol;
  BOOL m_bInBody;
  BOOL m_bInHead;
  BOOL m_bInHtml;
  BOOL m_bHasHead;
  BOOL m_bInComment;

}; /* class HtmlWriter */

#endif // HTMLWRITER_H not defined
/*==End of File=======================================================*/
