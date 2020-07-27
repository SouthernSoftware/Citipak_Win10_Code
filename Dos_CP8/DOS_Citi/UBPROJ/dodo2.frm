VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2436
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2436
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This utility analyses Wise for windows installer packages,
' copies the files from the "c:\" directory to the ".\source\" directory
' where the msi file exists. It produces an excel spreadsheet of non "C:\" items
' which are usually components.


Option Explicit

Dim databasePath, TargetDir, objFS
Const msiOpenDatabaseModeReadOnly = 0
Const msiOpenDatabaseModeTransact = 1

'Set objFS = CreateObject("Scripting.FileSystemObject")

Dim argNum, argCount: argCount = Wscript.Arguments.Count
If (argCount < 1) Then
    databasePath = InputBox("Please enter the Msi file name including path", , "")
    If databasePath = "" Then
        Fail "Cancel Selected"
    End If
Else
    databasePath = Wscript.Arguments(0)
End If
If (Not objFS.FileExists(databasePath)) Then Fail "File does not exist"
TargetDir = Mid(databasePath, 1, InStrRev(databasePath, "\")) & "Source\"
'wscript.Echo TargetDir

' Dump interesting stuff to an excel spreadsheet
'dim objXL, iNonCDrive: iNonCDrive = 2
'set objXL = CreateObject( "Excel.Application" )
'objXL.Visible = TRUE
'objXL.WorkBooks.Add

' Connect to Windows installer object
Dim openMode: openMode = msiOpenDatabaseModeTransact
On Error Resume Next
Dim installer: Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer"): CheckError

' Open database
Dim database: Set database = installer.OpenDatabase(databasePath, openMode): CheckError

' Process SQL statements
Dim query, view, record, message, rowData
    query = "Select SourcePath from WiseSourcePath"
    Set view = database.OpenView(query): CheckError
    view.Execute: CheckError
    Do
        Set record = view.Fetch
        If record Is Nothing Then Exit Do
        rowData = Empty
        If (UCase(Mid(record.StringData(1), 1, 3)) <> "C:\") Then
'            If it's not an item on the "C:\" drive then display it in a spreadsheet
'             objXL.Cells(iNonCDrive, 1).Value = record.StringData(1)
'            iNonCDrive = iNonCDrive + 1
        Else
            CopyFile record.StringData(1), TargetDir
        End If
    Loop

' Export the sourcepath table to a file
database.export "WiseSourcePath", "C:\", "test.txt"

' Process it to change source
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim InpFile, OutpFile, objRegExp, szLine, szOutputLine, DeleteFile
    Set objRegExp = New RegExp
    objRegExp.Pattern = "C:\\"
    objRegExp.ignorecase = True
    
    Set InpFile = objFS.OpenTextFile("C:\test.txt", ForReading)
    Set OutpFile = objFS.OpenTextFile("C:\test1.txt", ForWriting, True)
    While (InpFile.AtEndOfStream <> True)
        szLine = InpFile.readline
        If (objRegExp.Test(szLine)) Then
            szOutputLine = objRegExp.Replace(szLine, ".\source\")
        Else
            szOutputLine = szLine
        End If
        OutpFile.WriteLine (szOutputLine)
    Wend
    InpFile.Close
    OutpFile.Close
' Reimport the updated sourcepath table into the database
database.import "C:\", "test1.txt"

If openMode = msiOpenDatabaseModeTransact Then database.Commit

'Delete the temporary files created
Set DeleteFile = objFS.GetFile("C:\test.txt")
DeleteFile.Delete
Set DeleteFile = objFS.GetFile("C:\test1.txt")
DeleteFile.Delete

MsgBox ("Operation Complete")

Wscript.Quit 0

Sub CheckError()
    Dim message, errRec
    If Err = 0 Then Exit Sub
    message = Err.Source & " " & Hex(Err) & ": " & Err.Description
    If Not installer Is Nothing Then
        Set errRec = installer.LastErrorRecord
        If Not errRec Is Nothing Then message = message & vbLf & errRec.FormatText
    End If
    Fail message
End Sub

Sub Fail(message)
    Wscript.Echo message
    Wscript.Quit 2
End Sub

Sub CreatePath(szPath)
' This routine parses through a given path and creates it if it does not exist
Dim iOffset, iCurrentOffset, iLastOffset, szPathComponent, icount
icount = 1
iCurrentOffset = 1: iOffset = 1
iLastOffset = InStrRev(szPath, "\")
Do
    iOffset = InStr(iCurrentOffset + 1, szPath, "\")
    szPathComponent = Mid(szPath, 1, iOffset)
    If (Not objFS.FolderExists(szPathComponent)) Then
        objFS.CreateFolder (szPathComponent)
    End If
    iCurrentOffset = iOffset
    icount = icount + 1
Loop While (iOffset < iLastOffset)
End Sub

Sub CopyFile(szSource, szTarget)
Dim szDriveRemoved, szTargetFile
' Remove the drive info when creating the path
szDriveRemoved = Mid(szSource, 4)
szTargetFile = szTarget & szDriveRemoved
CreatePath szTargetFile
If (objFS.FileExists(szSource)) Then
    objFS.CopyFile szSource, szTargetFile
End If
End Sub
