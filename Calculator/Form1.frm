VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Equals 
      Caption         =   "="
      Height          =   615
      Left            =   3960
      TabIndex        =   18
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Divide 
      Caption         =   "/"
      Height          =   615
      Left            =   4800
      TabIndex        =   17
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton Multilply 
      Caption         =   "*"
      Height          =   615
      Left            =   3960
      TabIndex        =   16
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton Substract 
      Caption         =   "-"
      Height          =   615
      Left            =   4800
      TabIndex        =   15
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Add 
      Caption         =   "+"
      Height          =   615
      Left            =   3960
      TabIndex        =   14
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Inverse 
      Caption         =   "1/X"
      Height          =   615
      Left            =   4800
      TabIndex        =   13
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Pos_Neg 
      Caption         =   "+/-"
      Height          =   615
      Left            =   3960
      TabIndex        =   12
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton ClearBttn 
      Caption         =   "AC"
      Height          =   615
      Left            =   3000
      TabIndex        =   11
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton DotBttn 
      Caption         =   "."
      Height          =   615
      Left            =   2160
      TabIndex        =   10
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "9"
      Height          =   615
      Index           =   9
      Left            =   3000
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "8"
      Height          =   615
      Index           =   8
      Left            =   2160
      TabIndex        =   8
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "7"
      Height          =   615
      Index           =   7
      Left            =   1320
      TabIndex        =   7
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "6"
      Height          =   615
      Index           =   6
      Left            =   3000
      TabIndex        =   6
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "5"
      Height          =   615
      Index           =   5
      Left            =   2160
      TabIndex        =   5
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "4"
      Height          =   615
      Index           =   4
      Left            =   1320
      TabIndex        =   4
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "3"
      Height          =   615
      Index           =   3
      Left            =   3000
      TabIndex        =   3
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "0"
      Height          =   615
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   4800
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "2"
      Height          =   615
      Index           =   2
      Left            =   2160
      TabIndex        =   1
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "1"
      Height          =   615
      Index           =   1
      Left            =   1320
      TabIndex        =   0
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Display 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1320
      TabIndex        =   19
      Top             =   1080
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Operand1 As Double, Operand2 As Double
Dim Operator As String
Dim ClearDisplay As Boolean
Private Sub Form_Load()
    ClearDisplay = True
End Sub

Private Sub Digits_Click(Index As Integer)
    If ClearDisplay Then
        Display.Caption = ""
        ClearDisplay = False
    End If
    Display.Caption = Display.Caption + Digits(Index).Caption
End Sub
Private Sub Digits_KeyPress(Index As Integer, KeyAscii As Integer)
    'any number pressed
    Handle_KeyPressed (KeyAscii)
End Sub

Private Sub Add_Click()
    Operand1 = Val(Display.Caption)
    Operator = "+"
    ClearDisplay = True
End Sub
Private Sub Add_KeyPress(KeyAscii As Integer)
    Handle_KeyPressed (KeyAscii)
End Sub

Private Sub Divide_Click()
    Operand1 = Val(Display.Caption)
    Operator = "/"
    ClearDisplay = True
End Sub
Private Sub Divide_KeyPress(KeyAscii As Integer)
    Handle_KeyPressed (KeyAscii)
End Sub

Private Sub Multilply_Click()
    Operand1 = Val(Display.Caption)
    Operator = "*"
    ClearDisplay = True
End Sub
Private Sub Multilply_KeyPress(KeyAscii As Integer)
    Handle_KeyPressed (KeyAscii)
End Sub

Private Sub Substract_Click()
    Operand1 = Val(Display.Caption)
    Operator = "-"
    ClearDisplay = True
End Sub
Private Sub Substract_KeyPress(KeyAscii As Integer)
    Handle_KeyPressed (KeyAscii)
End Sub

Private Sub ClearBttn_Click()
    Display.Caption = "0"
End Sub
Private Sub ClearBttn_KeyPress(KeyAscii As Integer)
    Handle_KeyPressed (KeyAscii)
End Sub

Private Sub DotBttn_Click()
    If InStr(Display.Caption, ".") Then
        Exit Sub
    Else
        Display.Caption = Display.Caption + "."
    End If
End Sub
Private Sub DotBttn_KeyPress(KeyAscii As Integer)
    Handle_KeyPressed (KeyAscii)
End Sub

Private Sub Equals_Click()
    Dim Answer As Double
    Operand2 = Val(Display.Caption)
    If Operator = "+" Then Answer = Operand1 + Operand2
    If Operator = "-" Then Answer = Operand1 - Operand2
    If Operator = "*" Then Answer = Operand1 * Operand2
    If Operator = "/" And Operand2 <> 0 Then
        Answer = Operand1 / Operand2
    End If
    Display.Caption = Answer
    ClearDisplay = True
End Sub
Private Sub Equals_KeyPress(KeyAscii As Integer)
    Handle_KeyPressed (KeyAscii)
End Sub

Private Sub Handle_KeyPressed(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Digits_Click (KeyAscii - 48)
    End If
    If KeyAscii = 13 Then   'enter key pressed
        Equals_Click
    End If
    If KeyAscii = 42 Then   '* pressed
        Multilply_Click
    End If
    If KeyAscii = 43 Then   '+ pressed
        Add_Click
    End If
    If KeyAscii = 45 Then   '- pressed
        Substract_Click
    End If
    If KeyAscii = 46 Then   '. pressed
        DotBttn_Click
    End If
    If KeyAscii = 47 Then   '/ pressed
        Divide_Click
    End If
    If KeyAscii = 27 Then
        ClearBttn_Click
    End If
    'MsgBox KeyAscii
End Sub

Private Sub Inverse_Click()
    If Val(Display.Caption) <> 0 Then
        Display.Caption = 1 / Val(Display.Caption)
        ClearDisplay = True
    End If
End Sub

Private Sub Pos_Neg_Click()
    Display.Caption = -Val(Display.Caption)
    ClearDisplay = True
End Sub
