VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic - Tac - Toe"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Turn 
      BackColor       =   &H0080C0FF&
      Caption         =   "Player 1's Turn - (O)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Width           =   2895
   End
   Begin VB.CommandButton Help 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Tic - Tac - Toe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton Exit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   1935
   End
   Begin VB.CommandButton NewGame 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&New Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim player As Integer

Private Sub Digits_Click(Index As Integer)
    Dim i As Integer, count As Integer
    
    If Digits(Index).Caption = "" Then
        If player = 0 Then
            Digits(Index).Caption = "O"
            player = 1
            Turn.Caption = "Player 2's Turn - X"
        Else
            Digits(Index).Caption = "X"
            player = 0
            Turn.Caption = "Player 1's Turn - O"
        End If
    End If
    
    If isWin(0, 1, 2) <> "NO" Then
        changeColor 0, 1, 2
        MsgBox isWin(0, 1, 2)
        NewGame_Click
    End If
    If isWin(3, 4, 5) <> "NO" Then
        changeColor 3, 4, 5
        MsgBox isWin(3, 4, 5)
        NewGame_Click
    End If
    If isWin(6, 7, 8) <> "NO" Then
        changeColor 6, 7, 8
        MsgBox isWin(6, 7, 8)
        NewGame_Click
    End If
    If isWin(0, 3, 6) <> "NO" Then
        changeColor 0, 3, 6
        MsgBox isWin(0, 3, 6)
        NewGame_Click
    End If
    If isWin(1, 4, 7) <> "NO" Then
        changeColor 1, 4, 7
        MsgBox isWin(1, 4, 7)
        NewGame_Click
    End If
    If isWin(2, 5, 8) <> "NO" Then
        changeColor 2, 5, 8
        MsgBox isWin(2, 5, 8)
        NewGame_Click
    End If
    If isWin(0, 4, 8) <> "NO" Then
        changeColor 0, 4, 8
        MsgBox isWin(0, 4, 8)
        NewGame_Click
    End If
    If isWin(2, 4, 6) <> "NO" Then
        changeColor 2, 4, 6
        MsgBox isWin(2, 4, 6)
        NewGame_Click
    End If
    If isGameOver Then
        MsgBox "DRAW"
        NewGame_Click
        'End Sub
    End If
End Sub
Private Sub Exit_Click()
    End
End Sub

Private Sub Help_Click()
    MsgBox "Welcome to the Game of 'Tic-Tac-Toe'. This is a Two Player Game. The first player can make a 'O' on any empty square by clicking on it. The second player then can make a 'X' on any empty square by clicking on it."
End Sub

Private Sub NewGame_Click()
    Dim i As Integer
    For i = 0 To 8
        Digits(i).Caption = ""
        Digits(i).BackColor = &HC0C000
    Next
    player = 0
    Turn.Caption = "Player 1's Turn - (O)"
End Sub
Function isWin(i As Integer, j As Integer, k As Integer) As String
    If Digits(i).Caption = "O" And Digits(j).Caption = "O" And Digits(k).Caption = "O" Then
        isWin = "Player 1 (O) Wins"
    ElseIf Digits(i).Caption = "X" And Digits(j).Caption = "X" And Digits(k).Caption = "X" Then
        isWin = "Player 2 (X) Wins"
    Else
        isWin = "NO"
    End If
End Function

Function isGameOver() As Boolean
    Dim i As Integer, count As Integer
    For i = 0 To 8
        If Digits(i).Caption <> "" Then
            count = count + 1
        End If
    Next
    If count = 9 Then
        isGameOver = True
    Else
        isGameOver = False
    End If
    'MsgBox "HI"
End Function
Private Sub changeColor(i As Integer, j As Integer, k As Integer)
    Digits(i).BackColor = &H8080FF
    Digits(j).BackColor = &H8080FF
    Digits(k).BackColor = &H8080FF
End Sub
