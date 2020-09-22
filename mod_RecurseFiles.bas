Attribute VB_Name = "mod_RecurseFiles"
Option Explicit

'   Copyright © 2001 DonkBuilt Software
'   Written by Allen S. Donker
'   All rights reserved.

'************************************************************
'   Takes a valid path to a folder and, using recursion,
'   fills a listbox with the complete path and filename
'   of each file in the folder passed in
'************************************************************

Public Enum eFileAttribute
    filesAll
    filesNormal
    filesArchive
    filesCompressed
    filesHidden
    filesReadOnly
    filesSystem
End Enum


'************************************************************
'   Folder/File attributes constants
'************************************************************
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_COMPRESSED = &H800
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Const INVALID_HANDLE_VALUE = -1
Const ERROR_NO_MORE_FILES = 18&
Const MAX_PATH = 255


'************************************************************
'   Struct to hold file data
'************************************************************
Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type


'************************************************************
'   Struct to hold data returned from API calls
'************************************************************
Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type


'************************************************************
'   Windows API declarations
'************************************************************
Private Declare Function FindFirstFile& Lib "kernel32" Alias "FindFirstFileA" _
                                            (ByVal lpFileName As String, _
                                            lpFindFileData As WIN32_FIND_DATA)
Private Declare Function FindNextFile& Lib "kernel32" Alias "FindNextFileA" _
                                            (ByVal hFindFile As Long, _
                                            lpFindFileData As WIN32_FIND_DATA)
Private Declare Function FindClose& Lib "kernel32" (ByVal hFindFile As Long)


'************************************************************
'   To keep track of the number of files listed
'************************************************************
Public fileCount As Long
Public bCancelFileListAction As Boolean


Public Sub RecurseFiles(ByVal sFolderPath As String, _
                        Optional objListBox As ListBox, _
                        Optional lblCounter As Label, _
                        Optional bCancel As Boolean = False)

Dim hFindFile As Long
Dim ReturnValue As Long
Dim Filename As String
Dim fileData As WIN32_FIND_DATA


            'Get First Directory Entry. File value returned will end in
            'a null-terminated string.
    hFindFile = FindFirstFile(sFolderPath, fileData)
  
            'Exit if there was an Error Getting First Entry
    If hFindFile = INVALID_HANDLE_VALUE Then
        FindClose (hFindFile)
        Exit Sub
    End If
  
            'Initialize ReturnValue
    ReturnValue = 1
  
    Do While ReturnValue <> 0 And bCancelFileListAction = False
    
                'Remove the null charactor from Returned Filename
    Filename = StripNulls(fileData.cFileName)
  
                ' If it is a Directory but NOT the "." or ".." directories
                ' get all the Files in that directory (starts the recursion)
    If Filename <> "." And Filename <> ".." And _
                    fileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then

        RecurseFiles Mid(sFolderPath, 1, Len(sFolderPath) - 3) & Filename & "\*.*", objListBox, lblCounter


    Else        'If the item is not a directory (folder) and...
        
        
        If Filename <> "." And Filename <> ".." And Filename <> "" And Not _
                    fileData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
            
          
                    '...if the file is a Read Only file...
            If fileData.dwFileAttributes And FILE_ATTRIBUTE_READONLY Then
                
                            '...do something with it,
                            'in this case, add it to a listbox
                PerformFileAction Mid(sFolderPath, 1, Len(sFolderPath) - 3) & Filename, _
                                        objListBox, lblCounter
                                        
                    
                DoEvents
            
            End If  'Read only file
            
        End If  'Not a folder
      
    End If  'Is a folder
      
                ' Get Next Entry
    ReturnValue = FindNextFile(hFindFile, fileData)
    
    If ReturnValue = 0 Then

        Filename = ""
        Exit Do

    End If
            
    DoEvents
      
  Loop

                                ' Close Handle
  FindClose (hFindFile)
  
End Sub


'************************************************************
'   What to do with the files found during the recursion
'************************************************************
Private Function PerformFileAction(sFilePath As String, _
                                    objListBox As ListBox, _
                                    lblCounter As Label)

    objListBox.AddItem sFilePath
                
    fileCount = fileCount + 1
    lblCounter.Caption = fileCount

End Function



'************************************************************
'   Remove the Read-Only attribute of files, warns before
'   changing System or Hidden files
'************************************************************
Public Function RemoveReadOnlyAttribute(sFilePathAndName As String)
On Error GoTo ErrH

Dim fso As New FileSystemObject
Dim pFile As File
Dim fileData As WIN32_FIND_DATA

            '   Create a File System object and get the File
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set pFile = fso.GetFile(sFilePathAndName)
    
            '   If the file is a Hidden or a System file,
            '   warn the user before changing it.
    If pFile.Attributes And Hidden Or pFile.Attributes And System Then
        
        Dim iAnswer As Integer
        iAnswer = MsgBox(pFile.Name & " is a hidden or system file. " & _
                        "It is highly recommended that you not change this file." & Chr(10) & Chr(10) & _
                        "Are you sure you want to change the status of this file?", _
                        vbYesNo, "Confirm file change")
        
            '   Remove the Read-only attribute if user selects Yes
        If iAnswer = vbYes Then _
                pFile.Attributes = pFile.Attributes - FILE_ATTRIBUTE_READONLY
    
    Else
            '   Otherwise it's not a Hidden or System file, so change it
        pFile.Attributes = pFile.Attributes - FILE_ATTRIBUTE_READONLY
    
    End If
    
    
    
    Set pFile = Nothing
    Set fso = Nothing

Exit Function
ErrH:
    Err.Raise Err.Number
End Function


'************************************************************
'   Make sure the path ends with a \ and
'   add *.* to retreive all file types
'************************************************************
Public Function FormatPath(sPath As String) As String

    If Right$(sPath, 1) <> "\" Then
        FormatPath = sPath & "\*.*"
    Else
        FormatPath = sPath & "*.*"
    End If

End Function


Private Function StripNulls(ByVal FileWithNulls As String) As String

  Dim NullPos As Integer
  
  NullPos = InStr(1, FileWithNulls, vbNullChar, 0)
  
  If NullPos <> 0 Then
    
        StripNulls = Left(FileWithNulls, NullPos - 1)
  
  End If

End Function





'************************************************************
'   Remove the Read-Only attribute of files, warns before
'   changing System or Hidden files
'************************************************************
Public Function SetReadOnlyAttribute(sFilePathAndName As String)
On Error GoTo ErrH

Dim fso As New FileSystemObject
Dim pFile As File
Dim fileData As WIN32_FIND_DATA

            '   Create a File System object and get the File
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set pFile = fso.GetFile(sFilePathAndName)
    
            '   If the file is a Hidden or a System file,
            '   warn the user before changing it.
    If pFile.Attributes And Hidden Or pFile.Attributes And System Then
        
        Dim iAnswer As Integer
        iAnswer = MsgBox(pFile.Name & " is a hidden or system file. " & _
                        "It is highly recommended that you not change this file." & Chr(10) & Chr(10) & _
                        "Are you sure you want to change the status of this file?", _
                        vbYesNo, "Confirm file change")
        
            '   Remove the Read-only attribute if user selects Yes
        If iAnswer = vbYes Then _
                pFile.Attributes = pFile.Attributes + FILE_ATTRIBUTE_READONLY
    
    Else
            '   Otherwise it's not a Hidden or System file, so change it
        pFile.Attributes = pFile.Attributes + FILE_ATTRIBUTE_READONLY
    
    End If
    
    
    
    Set pFile = Nothing
    Set fso = Nothing

Exit Function
ErrH:
    Err.Raise Err.Number
End Function


