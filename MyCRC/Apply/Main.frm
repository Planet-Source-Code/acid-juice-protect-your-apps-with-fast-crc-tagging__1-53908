VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ApplyCRC"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply CRC"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'dim
Dim MyFileName As String
Dim MyRetVal As String

'browse for file
With CommonDialog1
    .Filename = ""
    .DialogTitle = ApplicationName & " - Open eNewsletter Manager v2 message"
    .CancelError = False
    .Filter = "Executable files (*.exe)|*.exe"
    .ShowOpen
    If Len(.Filename) = 0 Then
        Exit Sub
    End If
    MyFileName = .Filename
End With

' perform patch
' the option TRUE enables checking that the bytes where the CRC is written to are empty 00 bytes.

MyRetVal = AppendCRC(MyFileName, True)

'output result message
MsgBox MyRetVal

End Sub


Private Sub Command2_Click()

Unload Main

End Sub


