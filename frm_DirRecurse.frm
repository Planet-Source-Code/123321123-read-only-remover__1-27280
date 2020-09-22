VERSION 5.00
Begin VB.Form frm_DirRecurse 
   Caption         =   "DonkBuilt Read-Only Remover"
   ClientHeight    =   5565
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   10635
   Icon            =   "frm_DirRecurse.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5565
   ScaleWidth      =   10635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   9495
      TabIndex        =   7
      ToolTipText     =   "Exit"
      Top             =   2025
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   9495
      TabIndex        =   6
      ToolTipText     =   "Cancel"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemoveAttribute 
      Caption         =   "&Remove RO"
      Height          =   375
      Left            =   9495
      TabIndex        =   5
      ToolTipText     =   "Remove Read-Only attribute"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   9495
      TabIndex        =   3
      ToolTipText     =   "Exit"
      Top             =   5085
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear List"
      Height          =   375
      Left            =   9495
      TabIndex        =   2
      ToolTipText     =   "Reset List"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdListRO 
      Caption         =   "&List RO Files"
      Height          =   375
      Left            =   9495
      TabIndex        =   1
      ToolTipText     =   "List Read-Only Files"
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox lstFileList 
      Height          =   5325
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
   Begin VB.Label lblFileCount 
      Height          =   255
      Left            =   9495
      TabIndex        =   4
      Top             =   4725
      Width           =   1095
   End
End
Attribute VB_Name = "frm_DirRecurse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'   Copyright © 2001 DonkBuilt Software
'   Written by Allen S. Donker
'   All rights reserved.

'***************************************************************
'   Demonstrates the use of mod_SelectFolder which
'   uses the windows API's to open a browse folder
'   dialog box, mod_RecurseFolders which uses
'   recursion to browse thru the folder to
'   list the files contained in the folder selected, and
'   how to change read-only files to not read-only.
'
'   WARNING: As is, this routine will also list and change
'   the read-only attribute for hidden and system files.
'   Be carefull which files you elect to remove the read-only
'   attribute on!!
'***************************************************************

Dim colSelectedPaths As New Collection

Private Sub cmdListRO_Click()
On Error GoTo ErrH
Dim sPath As String

    MousePointer = vbHourglass

            '   Open the Browse Folder dialog box and
            '   return the folder selected
    sPath = SelectFolder(Me, "Select folder")
  
            '   If no folder was selected, exit here
    If Len(sPath) = 0 Then Exit Sub
  
            
            '   Make sure the path ends with a \ and
            '   add *.* to retreive all file types
    sPath = FormatPath(sPath)


            '   Call NewPath to either 1)add the path
            '   to the collection of folders selected
            '   and then get the files within that folder
            '   using recursion or 2)return False if the
            '   folder has already been selected and
            '   don't do anything
    If NewPath(sPath) Then
        RecurseFiles sPath, lstFileList, lblFileCount
    End If


    MousePointer = vbDefault

Exit Sub
    
ErrH:
    MousePointer = vbDefault
    MsgBox Err.Number & Chr(10) & Err.Description
End Sub






'************************************************************
'   To prevent the same files being listed more than once,
'   put each path (folder) selected into a collection using
'   the path as both the item and the key. If a folder is
'   selected a second time, trying to add it to the
'   collection will raise an error, so this returns False
'************************************************************
Private Function NewPath(sNewPath As String) As Boolean
On Error GoTo ErrH

    colSelectedPaths.Add sNewPath, sNewPath
    NewPath = True

Exit Function
ErrH:
    NewPath = False
End Function



'************************************************************
'   Go thru each file in the listbox, rmoving
'   the Read Only attribute
'************************************************************
Private Sub cmdRemoveAttribute_Click()
On Error GoTo ErrH
Dim i As Long, iAnswer As Integer

            '   Give the user a chance to cancel before starting
    iAnswer = MsgBox("All files listed will have their read-only property changed. Are you sure you want to continue?", vbYesNo, "Continue?")
    If iAnswer = vbNo Then Exit Sub

    MousePointer = vbHourglass
    
    For i = 0 To lstFileList.ListCount - 1
        
        RemoveReadOnlyAttribute lstFileList.List(i)
        If bCancelFileListAction Then Exit For
        
        DoEvents
        
    Next
    
    MousePointer = vbDefault

Exit Sub
ErrH:
    MousePointer = vbDefault
    MsgBox Err.Number & Chr(10) & Err.Description
    Exit Sub
End Sub



'************************************************************
'   Clear list and collection to start over
'************************************************************
Private Sub cmdClear_Click()
On Error GoTo ErrH

    Set colSelectedPaths = Nothing
    Set colSelectedPaths = New Collection
    
    lstFileList.Clear
    
    fileCount = 0
    lblFileCount.Caption = ""
    
Exit Sub
ErrH:
    Resume Next
End Sub


Private Sub cmdCancel_Click()
    bCancelFileListAction = True
End Sub


Private Sub cmdAbout_Click()
    frm_About.Show
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
