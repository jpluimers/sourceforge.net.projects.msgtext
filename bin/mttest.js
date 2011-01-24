/*== mttest.js =========================================================
Script for converting many .msg files and noting the return code.
Version     : $Id$
Application : MsgText
Description : Script for converting many .msg files and noting the return code.
------------------------------------------------------------------------
Created    : 14.10.2009, Hartwig Thomas
Copyright (c) 2009, Enter AG, Zurich, Switzerland
======================================================================*/
/* Global constants */
var g_sProgram = "mttest 1.0";
var g_sBanner = "MsgText test\n";

/* the file system object is needed everywhere:
http://msdn.microsoft.com/en-us/library/6kxy1a51(VS.85).aspx */
var g_fso = WScript.CreateObject("Scripting.FileSystemObject");

/* verbose parameter from command line */
var g_bVerbose = WScript.Arguments.Named.Exists("v");

/*----------------------------------------------------------------------
getScriptDir returns this script's full path name at run time
----------------------------------------------------------------------*/
function getScriptDir()
{
  var fileScript = g_fso.getFile(WScript.ScriptFullName);
  var sScriptDir = fileScript.ParentFolder.Path;
  /* all folders must have a trailing slash */
  if (sScriptDir.substring(sScriptDir.length-1) != "\\")
      sScriptDir = sScriptDir + "\\";
  return sScriptDir;
} /* getScriptDir */

/*----------------------------------------------------------------------
Parameters object
----------------------------------------------------------------------*/
function Parameters()
{
  this.folderSource = null;
  this.folderTarget = null;
} /* Parameters */
        
/*----------------------------------------------------------------------
displayHelp displays the syntax of the script call.
----------------------------------------------------------------------*/
function displayHelp(parms)
{
  WScript.StdOut.WriteLine("Calling syntax:");
  WScript.StdOut.WriteLine("  cscript mttest.js [/h] | <SourceFolder> [<TargetFolder>]");
  WScript.StdOut.WriteLine("where");
  WScript.StdOut.WriteLine("  /h               displays calling syntax");
  WScript.StdOut.WriteLine("  <SourceFolder>   is the source folder of the .msg files to be converted");
  WScript.StdOut.WriteLine("  <TargetFolder>   is the target folder to be produced");
  WScript.StdOut.WriteLine("                   (Default: " + parms.folderTarget.Path + ")");
} /* displayHelp */
        
/*----------------------------------------------------------------------
getParameters fills the parameters object with values from the command line.
----------------------------------------------------------------------*/
function getParameters(parms)
{
  var parmsResult = null;
  var bValid = false;
  var wsn = WScript.Arguments.Named;
  var wsu = WScript.Arguments.Unnamed;
  /* /h for help */
  if (wsn.Exists("h") || wsn.Exists("?") || wsn.Exists("help"))
    displayHelp(parms);
  else
  {
    if (wsu.Count > 0)
    {
      try
      {
        parms.folderSource = g_fso.GetFolder(wsu.Item(0));
        bValid = true;
        if (wsu.Count > 1)
        {
          bValid = false;
          parms.folderTarget = g_fso.GetFolder(wsu.Item(1));
          bValid = true;
        }
      }
      catch (e)
      {
        WScript.StdErr.WriteLine("Source folder " + wsu.Item(0) + " does not exist!");
      }
    }
    else
      WScript.StdErr.WriteLine("No source folder given! Use argument /h to display usage information.");
  }
  if (bValid)
  {
    WScript.StdOut.WriteLine("Source folder: \""+parms.folderSource.Path+"\"");
    WScript.StdOut.WriteLine("Target folder: \""+parms.folderTarget.Path+"\"");
    parmsResult = parms;
  }
  return parmsResult;
} /* getParameters */

/*----------------------------------------------------------------------
(public) run
run uses WScript.shell.Exec, and Wait for running a command
and returns the exit code. If requested, the contents of StdOut and StdErr
are piped through or saved in this.sOutput/this.sError.
Parameters:
iOuputLength if < 0, then child output disappears,
if = 0 then child output is piped through to stdout,
if > 0 then the last iOutputLength bytes of output are
stored in this.sOutput.
iErrorLength ditto.
----------------------------------------------------------------------*/
function run(sCommand, iOutputLength, iErrorLength)
{
  var shell = WScript.CreateObject("WScript.Shell");
  var sOutput = "";
  var sError = "";
  var iMsSleep = 100; /* check stdout and stderr all 100 ms */
  var iReturn = 0;
  /* execute the command */
  var wse = shell.Exec(sCommand);
  /* do a not terribly busy loop, waiting for completion */
  while (wse.Status == 0)
  {
    /* sleep 100 ms */
    WScript.Sleep(iMsSleep);
    /* if there is output then read it */
    if (!wse.StdOut.AtEndOfStream)
    {
      for (var sChar = wse.StdOut.Read(1);
           (sChar != null) && (sChar.length > 0);
           sChar = wse.StdOut.Read(1))
      {
        if (iOutputLength == 0)
          WScript.StdOut.Write(sChar);
        else if (iOutputLength > 0)
        {
          if (sOutput.length < iOutputLength)
            sOutput = sOutput + sChar;
          else
            sOutput = sOutput.substring(1) + sChar;
        }
      }
    }
    /* if there is error output then read it */
    if (!wse.StdErr.AtEndOfStream)
    {
      for (var sChar = wse.StdErr.Read(1);
           (sChar != null) && (sChar.length > 0);
           sChar = wse.StdErr.Read(1))
      {
        if (iErrorLength == 0)
          WScript.StdErr.Write(sChar);
        else if (iErrorLength > 0)
        {
          if (sError.length < iErrorLength)
            sError = sError + sChar;
          else
            sError = sError.substring(1) + sChar;
        }
      }
    }
  } /* while wse.Status == 0 */
  iReturn = wse.ExitCode;
  return iReturn;
} /* run */

/*----------------------------------------------------------------------
testFile
calls MsgText for conversion of .msg file to .txt file and writes
return code and .msg file to StdOut.
----------------------------------------------------------------------*/
function testFile(sSourceFile, sTargetFile)
{
  WScript.StdOut.Write(sSourceFile + "\t" + sTargetFile + "\t");
  var iResult = run("MsgText.exe \"" + sSourceFile + "\" \"" + sTargetFile, -1, -1);
  WScript.StdOut.WriteLine(iResult);
} /* testFile */

/*----------------------------------------------------------------------
testFolder
calls testFolder recursively for each contained folder
calls testFile for each conatined .msg file
----------------------------------------------------------------------*/
function testFolder(folderSource, folderTarget)
{
  /* sub folders */
  for (var enumFolders = new Enumerator(folderSource.SubFolders);
       !enumFolders.atEnd();
       enumFolders.moveNext())
  {
    var folder1 = enumFolders.item();
    var folder2 = g_fso.CreateFolder(folderTarget.Path + "\\" + folder1.Name);
    testFolder(folder1, folder2);
  }
  /* files */
  for (var enumFiles = new Enumerator(folderSource.Files);
       !enumFiles.atEnd();
       enumFiles.moveNext())
  {
    var file1 = enumFiles.item();
    var sExtension = g_fso.getExtensionName(file1.Path);
    if (sExtension == "msg")
    {
      var sFile2 = folderTarget.Path + "\\" + file1.Name;
      sFile2 = sFile2.substr(0,sFile2.length-sExtension.length)+"txt";
      testFile(file1.Path, sFile2);
    }
  }
} /* testFolder */

/*----------------------------------------------------------------------
main
----------------------------------------------------------------------*/
function main()
{
  WScript.StdOut.WriteLine(g_sProgram+": "+g_sBanner);
  var iResult = -1;
  var parms = new Parameters();
  g_fso.DeleteFolder(g_fso.GetSpecialFolder(2).Path + "\\MapiArchive",true);
  parms.folderTarget = g_fso.CreateFolder(g_fso.GetSpecialFolder(2).Path + "\\MapiArchive");
  parms = getParameters(parms);
  if (parms != null)
    testFolder(parms.folderSource, parms.folderTarget);
  return iResult;
} /* main */
              
/*----------------------------------------------------------------------
start it and return the result code
----------------------------------------------------------------------*/
try { WScript.quit(main()); }
catch (e)
{
  WScript.Echo("Exception " + e.name + ": " + e.message);
}
