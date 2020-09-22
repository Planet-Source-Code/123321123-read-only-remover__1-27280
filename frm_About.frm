VERSION 5.00
Begin VB.Form frm_About 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About DonkBuilt Read Only Remover"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5085
      TabIndex        =   0
      Top             =   3645
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Ver 1.00"
      Height          =   240
      Left            =   180
      TabIndex        =   6
      Top             =   2115
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "For suggestion, comments or bug reports, e-mail the developer at kc5kqm@vvm.com"
      Height          =   420
      Left            =   180
      TabIndex        =   5
      Top             =   2565
      Width           =   5010
   End
   Begin VB.Label Label2 
      Caption         =   "Visit the DonkBuilt website at http://www.vvm.com/~adonker"
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Top             =   3105
      Width           =   4470
   End
   Begin VB.Label Label1 
      Caption         =   $"frm_About.frx":0000
      Height          =   600
      Left            =   180
      TabIndex        =   3
      Top             =   3420
      Width           =   4605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   195
      X2              =   6135
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      Index           =   0
      X1              =   180
      X2              =   6120
      Y1              =   1875
      Y2              =   1875
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  'Right Justify
      Caption         =   "Copyright © 2001 DonkBuilt Software"
      Height          =   240
      Left            =   3285
      TabIndex        =   2
      Top             =   2115
      Width           =   2805
   End
   Begin VB.Label lblProgram 
      Caption         =   "Read Only Remover"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   960
      Left            =   180
      TabIndex        =   1
      Top             =   810
      Width           =   3705
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   180
      Picture         =   "frm_About.frx":00C7
      Top             =   135
      Width           =   1920
   End
End
Attribute VB_Name = "frm_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub
