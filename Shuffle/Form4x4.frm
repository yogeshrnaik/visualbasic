VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Game Of Shuffle"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   15
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3960
      Width           =   600
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   14
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3960
      Width           =   600
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   13
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3960
      Width           =   600
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   12
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3960
      Width           =   600
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   11
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   600
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   10
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3120
      Width           =   600
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   9
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3120
      Width           =   600
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1440
      Width           =   600
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   1
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1440
      Width           =   600
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   2
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1440
      Width           =   600
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   3
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      Width           =   600
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   4
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2280
      Width           =   600
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   5
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   600
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   6
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2280
      Width           =   600
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   7
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   600
   End
   Begin VB.CommandButton Digits 
      BackColor       =   &H00C0C000&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   8
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   600
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
      TabIndex        =   4
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton MovesBttn 
      BackColor       =   &H0080C0FF&
      Caption         =   "Moves = 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "The Game Of Shuffle"
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
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   1680
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim emp As Integer, temp(16) As Integer
Dim i As Integer, moves As Integer
Dim done As Boolean, gameCompleted As Boolean
Private Sub Command1_KeyPress(KeyAscii As Integer)
    Handle_KeyPress (KeyAscii)
End Sub
Private Sub Digits_Click(Index As Integer)
  If done Then
    If ((Digits(Index).Top - 840 = Digits(emp).Top) And (Digits(Index).Left = Digits(emp).Left)) Then
        'empty cell is at the top of clicked cell
        Digits(emp).Caption = Digits(Index).Caption
        Digits(Index).Caption = ""
        Digits(emp).Visible = True
        Digits(Index).Visible = False
        emp = Index
        moves = moves + 1
    End If
        
    If ((Digits(Index).Top + 840 = Digits(emp).Top) And (Digits(Index).Left = Digits(emp).Left)) Then
        'empty cell is at the bottom of clicked cell
        Digits(emp).Caption = Digits(Index).Caption
        Digits(Index).Caption = ""
        Digits(emp).Visible = True
        Digits(Index).Visible = False
        emp = Index
        moves = moves + 1
    End If
        
    If ((Digits(Index).Left - 840 = Digits(emp).Left) And (Digits(Index).Top = Digits(emp).Top)) Then
        'empty cell is at the left of clicked cell
        Digits(emp).Caption = Digits(Index).Caption
        Digits(Index).Caption = ""
        Digits(emp).Visible = True
        Digits(Index).Visible = False
        emp = Index
        moves = moves + 1
    End If
        
    If ((Digits(Index).Left + 840 = Digits(emp).Left) And (Digits(Index).Top = Digits(emp).Top)) Then
        'empty cell is at the right of clicked cell
        Digits(emp).Caption = Digits(Index).Caption
        Digits(Index).Caption = ""
        Digits(emp).Visible = True
        Digits(Index).Visible = False
        emp = Index
        moves = moves + 1
    End If
    
    If ((Digits(Index).Top - 840 = Digits(emp).Top) And (Digits(Index).Left - 840 = Digits(emp).Left)) Then
        'empty cell is at the top left of clicked cell
        Digits(emp).Caption = Digits(Index).Caption
        Digits(Index).Caption = ""
        Digits(emp).Visible = True
        Digits(Index).Visible = False
        emp = Index
        moves = moves + 1
    End If
    
    If ((Digits(Index).Top - 840 = Digits(emp).Top) And (Digits(Index).Left + 840 = Digits(emp).Left)) Then
        'empty cell is at the top right of clicked cell
        Digits(emp).Caption = Digits(Index).Caption
        Digits(Index).Caption = ""
        Digits(emp).Visible = True
        Digits(Index).Visible = False
        emp = Index
        moves = moves + 1
    End If
    
    If ((Digits(Index).Top + 840 = Digits(emp).Top) And (Digits(Index).Left - 840 = Digits(emp).Left)) Then
        'empty cell is at the bottom left of clicked cell
        Digits(emp).Caption = Digits(Index).Caption
        Digits(Index).Caption = ""
        Digits(emp).Visible = True
        Digits(Index).Visible = False
        emp = Index
        moves = moves + 1
    End If
    
    If ((Digits(Index).Top + 840 = Digits(emp).Top) And (Digits(Index).Left + 840 = Digits(emp).Left)) Then
        'empty cell is at the bottom right of clicked cell
        Digits(emp).Caption = Digits(Index).Caption
        Digits(Index).Caption = ""
        Digits(emp).Visible = True
        Digits(Index).Visible = False
        emp = Index
        moves = moves + 1
    End If
    
    MovesBttn.Caption = "Moves = " & moves
    
    gameCompleted = True
    'check if completed
    For i = 0 To 14
        If Val(Digits(i).Caption) <> i + 1 Then
            gameCompleted = False
        End If
    Next
    If gameCompleted Then
        MsgBox "Congratulations!!! You have completed the game in " & moves & " moves", , "Shuffle"
        NewGame_Click
    End If
    
  End If 'done

End Sub
Private Sub Digits_KeyPress(Index As Integer, KeyAscii As Integer)
    Handle_KeyPress (KeyAscii)
End Sub
Private Sub Exit_Click()
    End
End Sub
'Depending on the key pressed call the Digits_Click
'function with correct argument
Private Sub Handle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 50 Then 'Down Key Pressed
        If emp <> 0 And emp <> 1 And emp <> 2 And emp <> 3 Then
            Digits_Click (emp - 4)
        End If
    End If
    
    If KeyAscii = 52 Then 'Left Key Pressed
        If emp <> 3 And emp <> 7 And emp <> 11 And emp <> 15 Then
            Digits_Click (emp + 1)
        End If
    End If
    
    If KeyAscii = 54 Then 'Right Key Pressed
        If emp <> 0 And emp <> 4 And emp <> 8 And emp <> 12 Then
            Digits_Click (emp - 1)
        End If
    End If
    
    If KeyAscii = 56 Then 'Up Key Pressed
        If emp <> 12 And emp <> 13 And emp <> 14 And emp <> 15 Then
            Digits_Click (emp + 4)
        End If
    End If
    
    'Up Left Key Pressed (Num pad - 7)
    If KeyAscii = 55 Then
        If emp <> 3 And emp <> 7 And emp <> 11 And emp <> 12 And emp <> 13 And emp <> 14 And emp <> 15 Then
            Digits_Click (emp + 5)
        End If
    End If
    
    'Up Right Key Pressed (Num pad - 9)
    If KeyAscii = 57 Then
        If emp <> 0 And emp <> 4 And emp <> 8 And emp <> 12 And emp <> 13 And emp <> 14 And emp <> 15 Then
            Digits_Click (emp + 3)
        End If
    End If

    'Bottom Left Key Pressed (Num pad - 1)
    If KeyAscii = 49 Then
        If emp <> 0 And emp <> 1 And emp <> 2 And emp <> 3 And emp <> 7 And emp <> 11 And emp <> 15 Then
            Digits_Click (emp - 3)
        End If
    End If

    'Bottom Right Key Pressed (Num pad - 3)
    If KeyAscii = 51 Then
        If emp <> 0 And emp <> 1 And emp <> 2 And emp <> 3 And emp <> 4 And emp <> 8 And emp <> 12 Then
            Digits_Click (emp - 5)
        End If
    End If
End Sub
Private Sub Exit_KeyPress(KeyAscii As Integer)
    Handle_KeyPress (KeyAscii)
End Sub
Private Sub Help_Click()
    MsgBox "Welcome to The Game of Shuffle. Its a Puzzle. It contains 15 numbered squares pieces, which can be moved horizontally or vertically or diagonally by clicking on it or by using Num-Pad Arrow Keys. Your job is to arrange the numbered squares in asceding order as shown at the initial screen of the Game.", , "Shuffle"
End Sub
Private Sub Help_KeyPress(KeyAscii As Integer)
    Handle_KeyPress (KeyAscii)
End Sub
Private Sub MovesBttn_KeyPress(KeyAscii As Integer)
    Handle_KeyPress (KeyAscii)
End Sub
Private Sub NewGame_Click()
    generate
    moves = 0
    MovesBttn.Caption = "Moves = 0"
    done = True
    Digits(emp).Visible = False
End Sub
Private Sub generate()
    Dim found As Boolean
    Dim temp As Integer, j As Integer, size As Integer
    
    size = 0
    For j = 0 To 15
        Digits(j).Caption = -1
        Digits(j).Visible = True
    Next
    Randomize
    While size < 16
        temp = (Rnd() * 100) Mod 16
        'MsgBox temp
        If temp >= 0 And temp <= 15 Then
            For j = 0 To size
                If Val(Digits(j).Caption) = temp Then
                    found = True
                End If
            Next
            If found Then
                found = False
            Else
                Digits(size).Caption = temp
                If temp = 0 Then
                    emp = size
                    Digits(size).Caption = ""
                End If
                size = size + 1
            End If
        End If
    Wend
End Sub
Private Sub NewGame_KeyPress(KeyAscii As Integer)
    Handle_KeyPress (KeyAscii)
End Sub
