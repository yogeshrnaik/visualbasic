VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Hello"
      Height          =   1575
      Left            =   1560
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim color As Boolean

Private Sub Command1_Click()
    If (Not color) Then
        Command1.BackColor = &HC0C0C0 'normal color
        color = True
        'MsgBox color
    Else
        Command1.BackColor = &H80FF& 'red color
        color = False
        'MsgBox color
    End If
End Sub
