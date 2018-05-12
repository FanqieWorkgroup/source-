VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8010
   LinkTopic       =   "Form2"
   ScaleHeight     =   4350
   ScaleWidth      =   8010
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "立即处理"
      Height          =   1095
      Left            =   5040
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2760
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "发现有病毒正感染您的电脑"
      Height          =   1260
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3840
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Sub Command1_Click()
Dim i As Integer
For i = 0 To List1.ListCount - 1
DeleteFile List1.List(i)
DoEvents
Next i
MsgBox "finish", vbOKCancel
End Sub

Private Sub Form_Load()
Label1.FontSize = "30"
End Sub

