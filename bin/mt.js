/*===$Workfile: mt.js $=================================================
mt is the user interface frontend for the command-line tool MsgText.
Version    : $Id$
Application: MSG file archving
Platform   : JavaScript, Windows, MAPI, Internet Explorer
Description: This user interface frontend is used for converting .msg
             files to .txt files and extracting their attachments.
------------------------------------------------------------------------
Copyright  : 2009, Enter AG, Zurich, Switzerland
Developed  : Hartwig Thomas, 11. June 2009
=======================================================================*/
var g_sBanner = "mt 1.01: convert .msg files to .txt files\r\n" +
  "(c) 2009, Enter AG, Zurich, Switzerland";

var g_fso = WScript.CreateObject("Scripting.FileSystemObject");
var g_wsh = WScript.CreateObject("WScript.Shell");
var g_ie;
var g_sScriptDir = g_fso.getFile(WScript.ScriptFullName).ParentFolder.Path + "\\";

/*-----------------------------------------------------------------------
constants
-----------------------------------------------------------------------*/
var iRETURN_OK = 0;
var iRETURN_WARNING = 4;
var iRETURN_ERROR = 8;
var iRETURN_FATAL = 12;

/*-----------------------------------------------------------------------
global parameters of MsgText
-----------------------------------------------------------------------*/
var g_sMsgFile = null;
var g_sTextFile = null;
var g_sAttachmentDir = null;
var g_bFinished = false;
var g_bLaunch = false;
var g_bAttachmentFromText = false;

/*-----------------------------------------------------------------------
initialize global parameters of MsgText
-----------------------------------------------------------------------*/
function initParms()
{
  g_sMsgFile = "";
  g_sTextFile = "";
  g_sAttachmentDir = "";
};  /* initParms */

/*-----------------------------------------------------------------------
display the form
-----------------------------------------------------------------------*/
function displayForm(ie)
{
  g_ie.navigate("about:blank");
  g_ie.menubar = false;
  g_ie.toolbar = false;
  g_ie.statusbar = false;
  g_ie.resizable = false;
  g_ie.document.open("text/html","replace");
  g_ie.document.writeln("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\">");
  g_ie.document.writeln("<html>");
  g_ie.document.writeln("  <head>");
  g_ie.document.writeln("    <meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">");
  g_ie.document.writeln("    <title>MsgText</title>");
  g_ie.document.writeln("    <style type=\"text/css\">");
  g_ie.document.writeln("      body { font-family: DejaVu Sans, sans-serif; font-size: 10pt; }");
  g_ie.document.writeln("      table { font-family: DejaVu Sans, sans-serif; font-size: 10pt; }");
  g_ie.document.writeln("      p { margin: 0; border: 0; padding: 0; }");
  g_ie.document.writeln("      form { margin: 0; border: 0; padding: 0; }");
  g_ie.document.writeln("    </style>");
  g_ie.document.writeln("  </head>");
  g_ie.document.writeln("  <body bgcolor=\"#C0C0C0\" scroll=\"no\">");
  g_ie.document.writeln("    <h1>MsgText</h1>");
  g_ie.document.writeln("    <p>"+g_sBanner+"</p>");
  g_ie.document.writeln("    <hr>");
  g_ie.document.writeln("    <p>To convert a .msg file into a .txt file and extract its attachments,");
  g_ie.document.writeln("      you have to choose the .msg file to be converted.</p>");
  g_ie.document.writeln("    <p>You may then change the name of the output .txt file and the attachment directory proposed.</p>");
  g_ie.document.writeln("    <hr>");
  g_ie.document.writeln("    <form name=\"mtform\">");
  g_ie.document.writeln("      <p><input type=\"file\" name=\"selector\" id=\"selector\" style=\"display:none\" /></p>");
  g_ie.document.writeln("      <table>");
  g_ie.document.writeln("        <tr>");
  g_ie.document.writeln("          <td align=\"right\">.msg file</td>");
  g_ie.document.writeln("          <td><input type=\"text\" name=\"msg\" id=\"msg\" size=\"70\" value=\"" + g_sMsgFile + "\" />");
  g_ie.document.writeln("               <input type=\"button\" name=\"msgbutton\" id=\"msgbutton\" value=\"   Select   \" /></td>");
  g_ie.document.writeln("        </tr>");
  g_ie.document.writeln("        <tr>");
  g_ie.document.writeln("          <td align=\"right\">.txt file</td>");
  g_ie.document.writeln("          <td><input type=\"text\" name=\"txt\" id=\"txt\" size=\"70\" value=\"" + g_sTextFile + "\" />");
  g_ie.document.writeln("               <input type=\"button\" name=\"msgbutton\" id=\"txtbutton\" value=\"   Select   \" /></td>");
  g_ie.document.writeln("        </tr>");
  g_ie.document.writeln("        <tr>");
  g_ie.document.writeln("          <td align=\"right\">attachment folder</td>");
  g_ie.document.writeln("          <td><input type=\"text\" name=\"att\" id=\"att\" size=\"70\" value=\"" + g_sAttachmentDir + "\" />");
  g_ie.document.writeln("               <input type=\"button\" name=\"attbutton\" id=\"attbutton\" value=\"   Select   \" /></td>");
  g_ie.document.writeln("        </tr>");
  g_ie.document.writeln("      </table>");
  g_ie.document.writeln("    </form>");
  g_ie.document.writeln("    <hr>");
  g_ie.document.writeln("      <input type=\"text\" name=\"message\" id=\"message\" value=\"\" size=\"110\" readonly=\"1\" style=\"color:yellow;background-color:lightgrey;\" />");
  g_ie.document.writeln("    <hr>");
  g_ie.document.writeln("    <p align=\"center\"><input type=\"button\" name=\"okbutton\" id=\"okbutton\" value=\"  Convert  \">");
  g_ie.document.writeln("                        <input type=\"button\" name=\"cancelbutton\" id=\"cancelbutton\" value=\"    Exit    \"></p>");
  g_ie.document.writeln("    <hr/>");
  g_ie.document.writeln("    <p>MsgText is free software: you can redistribute it and/or modify");
  g_ie.document.writeln("    it under the terms of the GNU General Public License as published by");
  g_ie.document.writeln("    the Free Software Foundation, either version 3 of the License, or");
  g_ie.document.writeln("    any later version.</p>");
  g_ie.document.writeln("    <p>MsgText is distributed in the hope that it will be useful,");
  g_ie.document.writeln("    but WITHOUT ANY WARRANTY; without even the implied warranty of");
  g_ie.document.writeln("    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the");
  g_ie.document.writeln("    GNU General Public License for more details.</p>");
  g_ie.document.writeln("    <p>You should have received a copy of the GNU General Public License");
  g_ie.document.writeln("    along with this program.  If not, see <a href=\"http://www.gnu.org/licenses/\">http://www.gnu.org/licenses/</a>.</p>");
  g_ie.document.writeln("  </body>");
  g_ie.document.writeln("</html>");
  g_ie.Width = 750;
  g_ie.Height = 500;
  g_ie.document.getElementById("msg").focus();
};  /* displayForm */

/*----------------------------------------------------------------------
onCancelPressed is the event handler for the cancel button
----------------------------------------------------------------------*/
function onCancelPressed()
{
  g_bLaunch = false;
  g_bFinished = true;
};  /* onCancelPressed */

/*----------------------------------------------------------------------
onOkPressed is the event handler for the OK button
----------------------------------------------------------------------*/
function onOkPressed()
{
  g_sMsgFile = g_ie.document.getElementById("msg").value;
  g_sTextFile = g_ie.document.getElementById("txt").value;
  g_sAttachmentDir = g_ie.document.getElementById("att").value;
  if (g_sMsgFile)
  {
    sMessage = "Message file " + g_sMsgFile + " will be converted to " + g_sTextFile + ".\nPlease be patient ...";
    g_bLaunch = true;
    g_bFinished = true;
  }
  else
    sMessage = "Please enter .msg file!";
  g_ie.document.getElementById("message").value = sMessage;
};  /* onOkPressed */

/*----------------------------------------------------------------------
onMsgPressed is the event handler for the msg button
----------------------------------------------------------------------*/
function onMsgPressed()
{
  /* display file selector */
  var selector = g_ie.document.getElementById("selector");
  selector.click();
  g_sMsgFile = selector.value;
  g_ie.document.getElementById("msg").value = g_sMsgFile;
  if (!g_sTextFile)
  {
    var sExtension = g_fso.GetExtensionName(g_sMsgFile);
    if (sExtension)
      g_sTextFile = g_sMsgFile.substring(0, g_sMsgFile.length - sExtension.length) + "txt";
    else
      g_sTextFile = g_sMsgFile + ".txt";
    g_ie.document.getElementById("txt").value = g_sTextFile;
    if ((!g_sAttachmentDir) || g_bAttachmentFromText)
    {
      var sExtension = g_fso.GetExtensionName(g_sTextFile);
      if (sExtension)
        g_sAttachmentDir = g_sTextFile.substring(0, g_sTextFile.length - sExtension.length - 1) + "\\";
      else
        g_sAttachmentDir = g_sTextFile + "_\\";
      g_bAttachmentFromText = true;
      g_ie.document.getElementById("att").value = g_sAttachmentDir;
    }
  }
};  /* onMsgPressed */

/*----------------------------------------------------------------------
onTxtPressed is the event handler for the txt button
----------------------------------------------------------------------*/
function onTxtPressed()
{
  /* display file selector */
  /* display file selector */
  var selector = g_ie.document.getElementById("selector");
  selector.click();
  g_sTextFile = selector.value;
  var sExtension = g_fso.GetExtensionName(g_sTextFile);
  if (!sExtension)
    g_sTextFile = g_sTextFile + ".txt";
  g_ie.document.getElementById("txt").value = g_sTextFile;
  if ((!g_sAttachmentDir) || g_bAttachmentFromText)
  {
    var sExtension = g_fso.GetExtensionName(g_sTextFile);
    if (sExtension)
      g_sAttachmentDir = g_sTextFile.substring(0, g_sTextFile.length - sExtension.length - 1) + "\\";
    else
      g_sAttachmentDir = g_sTextFile + "_\\";
    g_bAttachmentFromText = true;
    g_ie.document.getElementById("att").value = g_sAttachmentDir;
  }
};  /* onTxtPressed */

/*----------------------------------------------------------------------
onAttPressed is the event handler for the att button
----------------------------------------------------------------------*/
function onAttPressed()
{
  var shell = WScript.CreateObject("Shell.Application");
  /* display folder selector */
  var folder = shell.BrowseForFolder(g_ie.HWND, "Select an attachment folder:", 1, "");
  if (folder)
  {
    g_sAttachmentDir = folder.Self.Path
    g_bAttachmentFromText = false;
    g_ie.document.getElementById("att").value = g_sAttachmentDir;
    folderAtt = g_fso.GetFolder(g_sAttachmentDir);
    g_sSelectFolder = folderAtt.ParentFolder;
  }
};  /* onAttPressed */

/*-----------------------------------------------------------------------
display the form and validate its entry
-----------------------------------------------------------------------*/
function runForm()
{
  displayForm();
  /* attach event handlers */
  g_ie.document.getElementById("okbutton").onclick = onOkPressed;
  g_ie.document.getElementById("cancelbutton").onclick = onCancelPressed;
  g_ie.document.getElementById("msgbutton").onclick = onMsgPressed;
  g_ie.document.getElementById("txtbutton").onclick = onTxtPressed;
  g_ie.document.getElementById("attbutton").onclick = onAttPressed;
  g_ie.visible = true;
  g_wsh.AppActivate("http:/// - MsgText - Windows Internet Explorer");
  /* wait until ok or cancel button was pressed */
  while (!g_bFinished) { WScript.Sleep(100); }  
  return g_bLaunch;
};  /* runForm */

/*-----------------------------------------------------------------------
msgtext runs the command-line executable MsgText
-----------------------------------------------------------------------*/
function msgtext(sMsgFile,sTextFile,sAttachmentDir)
{
  var iReturn = iRETURN_OK;
  /* display wait cursor */
  var curSaved = g_ie.document.body.style.cursor;
  g_ie.document.body.style.cursor = "wait";
  /* run the command-line app */
  var sCommand = "\"" + g_sScriptDir + "MsgText.exe\" /d \""+sMsgFile+"\"";
  if (sTextFile)
    sCommand = sCommand + " \"" + sTextFile + "\"";
  if (sAttachmentDir)
  {
    if (sAttachmentDir.substr(sAttachmentDir.length - 1) == "\\")
      sCommand = sCommand + " \"" + sAttachmentDir + "\\\"";
    else
      sCommand = sCommand + " \"" + sAttachmentDir + "\"";
  }
  g_wsh.Run("cmd /K \"" + sCommand + "\"", 1, true);
  /* remove the wait cursor */
  g_ie.document.body.style.cursor = curSaved;
  return iReturn;
};  /* msgtext */

/*-----------------------------------------------------------------------
addTrusted adds the site to the trusted sites
-----------------------------------------------------------------------*/
var sREGKEY_ABOUT_BLANK = "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings\\ZoneMap\\Domains\\blank\\about";
var iTRUSTED_SITES = 2;
function addTrusted()
{
  var bAdded = false;
  try
  {
    var iZone = g_wsh.RegRead(sREGKEY_ABOUT_BLANK);
    if (iZone != iTRUSTED_SITES)
      bAdded = true;
  }
  catch (e) { bAdded = true; }
  if (bAdded)
    g_wsh.RegWrite(sREGKEY_ABOUT_BLANK, iTRUSTED_SITES, "REG_DWORD");
  return bAdded;
};  /* addTrusted */

/*-----------------------------------------------------------------------
removeTrusted removes the site from the trusted sites
-----------------------------------------------------------------------*/
function removeTrusted()
{
  var bRemoved = true;
  g_wsh.RegDelete(sREGKEY_ABOUT_BLANK);
  var sKeyBlank = sREGKEY_ABOUT_BLANK.substring(0,sREGKEY_ABOUT_BLANK.lastIndexOf("\\")+1);
  /* if about was the only entry under the blank domain, then remove it too. */
  try { g_wsh.RegDelete(sKeyBlank); }
  catch (e) { }
  return bRemoved;
};  /* removeTrusted */

/*-----------------------------------------------------------------------
main
-----------------------------------------------------------------------*/
function main()
{
  var iReturn = iRETURN_WARNING;
  /* get initial values for parameters of MsgText */
  initParms();
  /* add about:blank to the trusted sites */
  var bAboutBlankAdded = addTrusted();
  /* start IE */
  g_ie = WScript.CreateObject("InternetExplorer.Application", "IE_");
  if (g_ie)
  {
    /* display form and validate entry */
    if (runForm())
      msgtext(g_sMsgFile, g_sTextFile, g_sAttachmentDir);
    g_ie.Quit();
  }
  else
  {
    WScript.Echo("Internet Explorer not found!");
    iReturn = iRETURN_ERROR;
  }
  /* remove about:blank from the trusted sites */
  if (bAboutBlankAdded)
    removeTrusted();
  return iReturn;
};  /* main */

/*---------------------------------------------------------------------*/
try { WScript.quit(main()); }
catch (e)
{
  WScript.Echo("Exception " + e.name + ": " + e.message);
}