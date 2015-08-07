VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect Four"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Exit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton NewGame 
      Caption         =   "&New Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton DropBttn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Drop Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton DropBttn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Drop Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   4560
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton DropBttn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Drop Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   3720
      TabIndex        =   3
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton DropBttn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Drop Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   2880
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton DropBttn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Drop Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton DropBttn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Drop Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Title 
      BackColor       =   &H0080C0FF&
      Caption         =   "        Connect Four"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1035
      TabIndex        =   9
      Top             =   240
      Width           =   5895
   End
   Begin VB.Label WhoseTurn 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "    Yellow     Player's        Turn"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6000
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   41
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   40
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   39
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   38
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   37
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   36
      Left            =   360
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   35
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   34
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   33
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   32
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   31
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   30
      Left            =   360
      Shape           =   3  'Circle
      Top             =   5520
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   29
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   28
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   27
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   26
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   25
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   24
      Left            =   360
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   23
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   22
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   21
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   20
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   19
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   18
      Left            =   360
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   17
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   16
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   15
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   14
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   13
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   12
      Left            =   360
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   11
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   10
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   9
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   8
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   7
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   6
      Left            =   360
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   5
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   4
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   3
      Left            =   2880
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   615
   End
   Begin VB.Shape CirclePos 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   360
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim player As String
Dim i As Integer, j As Integer, currCircle As Integer
Dim currInCol(5) As Integer
Dim gameStatus As String
Dim gamePos(6, 5) As String

Private Sub DropBttn_Click(Index As Integer)
    If player = "Y" Then
        If currInCol(Index) >= 0 Then
            CirclePos(currInCol(Index)).FillColor = &HFFFF&   'yellow
            gamePos((currInCol(Index) - Index) / 6, Index) = "Y"
            WhoseTurn.Caption = "      Red        Player's        Turn"
            WhoseTurn.BackColor = &HFF&      'red
            player = "R"
        End If
    ElseIf player = "R" Then
            If currInCol(Index) >= 0 Then
                CirclePos(currInCol(Index)).FillColor = &HFF&     'red
                gamePos((currInCol(Index) - Index) / 6, Index) = "R"
                WhoseTurn.Caption = "    Yellow     Player's        Turn"
                WhoseTurn.BackColor = &HFFFF&   'yellow
                player = "Y"
            End If
    End If
    currInCol(Index) = currInCol(Index) - 6
    If (currInCol(Index) < -6) Then
        currInCol(Index) = currInCol(Index) + 6
    End If
    gameStatus = getGameStatus
    If gameStatus = "WIN" Then
        If player = "Y" Then
            MsgBox "Red Player Wins!", , "Connect Four"
        Else
            MsgBox "Yellow Player Wins!", , "Connect Four"
        End If
        NewGame_Click
    ElseIf gameStatus = "DRAW" Then
        MsgBox "DRAW", , "Connect Four"
        NewGame_Click
    End If
End Sub

Private Sub Exit_Click()
    End
End Sub

Private Sub Form_Load()
    NewGame_Click
End Sub

Function getGameStatus() As String
    getGameStatus = "GAME IN PROGRESS"
    Dim count As Integer
    Dim countRed As Integer, countYellow As Integer
    Dim row As Integer, col As Integer
    Dim k As Integer
    
    'check if draw
    For row = 0 To 6
        For col = 0 To 5
            If gamePos(row, col) = "Y" Or gamePos(row, col) = "R" Then
                count = count + 1
            End If
        Next
    Next
    If count = 42 Then
        getGameStatus = "DRAW"
    End If
    
    'row checking
    For row = 0 To 6
        For i = 0 To 2
            countRed = 0
            countYellow = 0
            For col = i To i + 3
                If (gamePos(row, col) = "Y") Then
                    countYellow = countYellow + 1
                ElseIf (gamePos(row, col) = "R") Then
                    countRed = countRed + 1
                End If
            Next
            If (countRed = 4) Then
                getGameStatus = "WIN"
            End If
            If (countYellow = 4) Then
                getGameStatus = "WIN"
            End If
            
            'show winning position
            If countYellow = 4 Or countRed = 4 Then
                For col = i To i + 3
                    CirclePos((6 * row) + col).FillColor = &H800000   ' dark blue
                Next
                row = 10
                i = 10
            End If
        Next
    Next
    
    'column checking
    countRed = 0
    countYellow = 0
    For col = 0 To 5
        For i = 0 To 3
            countRed = 0
            countYellow = 0
            For row = i To i + 3
                If (gamePos(row, col) = "Y") Then
                    countYellow = countYellow + 1
                ElseIf (gamePos(row, col) = "R") Then
                    countRed = countRed + 1
                End If
            Next
            If (countRed = 4) Then
                getGameStatus = "WIN"
            End If
            If (countYellow = 4) Then
                getGameStatus = "WIN"
            End If
            
            'show winning position
            If countYellow = 4 Or countRed = 4 Then
                For row = i To i + 3
                    CirclePos((6 * row) + col).FillColor = &H800000   ' dark blue
                Next
                row = 10
                i = 10
            End If
        Next
    Next
    
    'left to right diagonal
    countRed = 0
    countYellow = 0
    For row = 0 To 3
        For col = 0 To 2
            countRed = 0
            countYellow = 0
            For j = 0 To 3
                If (gamePos(row + j, col + j) = "Y") Then
                    countYellow = countYellow + 1
                ElseIf (gamePos(row + j, col + j) = "R") Then
                    countRed = countRed + 1
                End If
            Next
            If (countRed = 4) Then
                getGameStatus = "WIN"
            End If
            If (countYellow = 4) Then
                getGameStatus = "WIN"
            End If
            
            'show winning position
            If countYellow = 4 Or countRed = 4 Then
                For j = 0 To 3
                    CirclePos((6 * (row + j)) + (col + j)).FillColor = &H800000 ' dark blue
                Next
                row = 10
                col = 10
            End If
            
        Next
    Next
    
    
    'right to left diagonal
    countRed = 0
    countYellow = 0
    For col = 3 To 5
        For row = 0 To 3
            countRed = 0
            countYellow = 0
            For j = 0 To 3
                If (gamePos(row + j, col - j) = "Y") Then
                    countYellow = countYellow + 1
                ElseIf (gamePos(row + j, col - j) = "R") Then
                    countRed = countRed + 1
                End If
            Next
            If (countRed = 4) Then
                getGameStatus = "WIN"
            End If
            If (countYellow = 4) Then
                getGameStatus = "WIN"
            End If
            
            'show winning position
            If countYellow = 4 Or countRed = 4 Then
                For j = 0 To 3
                    CirclePos((6 * (row + j)) + (col - j)).FillColor = &H800000 ' dark blue
                Next
                row = 10
                col = 10
            End If
            
        Next
    Next
    
End Function

Private Sub NewGame_Click()
    player = "Y"
    currCircle = 0
    WhoseTurn.Caption = "    Yellow     Player's        Turn"
    WhoseTurn.BackColor = &HFFFF&   'yellow
    For i = 36 To 41
        currInCol(i - 36) = i
    Next
    For i = 0 To 41
        CirclePos(i).FillColor = &H80000005 ' White
    Next
    For i = 0 To 6
        For j = 0 To 5
            gamePos(i, j) = ""
        Next
    Next
    
End Sub
