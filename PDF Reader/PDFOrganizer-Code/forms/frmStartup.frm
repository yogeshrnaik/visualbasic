VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStartup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PDF Organizer"
   ClientHeight    =   3735
   ClientLeft      =   2760
   ClientTop       =   3240
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   120
      Picture         =   "frmStartup.frx":0000
      ScaleHeight     =   1155
      ScaleWidth      =   5715
      TabIndex        =   1
      Top             =   240
      Width           =   5775
   End
   Begin VB.Timer timer 
      Interval        =   1000
      Left            =   5400
      Top             =   2400
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "yogesh131080@rediffmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Maintains the data for PDF files. Search facility to find desired article based on various user defined search criteria."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   1560
      TabIndex        =   5
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Yogesh R Naik"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Developed By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblRemaining 
      Alignment       =   1  'Right Justify
      Caption         =   "Wait: 10 seconds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   0
      Top             =   3480
      Width           =   1815
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#################################################################################
Private m_counter As Integer
Private m_showingAsAbout As Boolean
Const DISPLAY_DURATION = 10
'#################################################################################
'Private Sub Form_KeyPress(KeyAscii As Integer)
'  'CTRL + E is the short cut to hide the form and start application immediately
'  If KeyAscii = 5 And Not m_showingAsAbout Then
'    Me.timer.Enabled = False
'    Me.Hide
'    mdiMain.Show
'  ElseIf m_showingAsAbout Then
'    MsgBox KeyAscii
'  End If
'End Sub
'#################################################################################
Private Sub Form_Load()
  m_counter = 0
  m_showingAsAbout = False
  Me.timer.Enabled = True
  Me.lblRemaining.Caption = "Wait: " & DISPLAY_DURATION & " seconds."
End Sub
'#################################################################################
'Private Sub Form_LostFocus()
'  If (m_showingAsAbout) Then
'    Me.Hide
'  End If
'End Sub
'#################################################################################
Private Sub Picture1_KeyPress(KeyAscii As Integer)
  'CTRL + E is the short cut to hide the form and start application immediately
  If KeyAscii = 5 And Not m_showingAsAbout Then
    Me.timer.Enabled = False
    Me.Hide
    mdiMain.Show
  ElseIf m_showingAsAbout And KeyAscii = 27 Then 'hit escape key
    Me.Hide
  End If
End Sub
'#################################################################################
Private Sub timer_Timer()
  If (m_counter >= 10) Then
    Me.timer.Enabled = False
    Me.Hide
    mdiMain.Show
  Else
    m_counter = m_counter + 1
    Me.ProgressBar1.Value = 10 * m_counter
    Me.lblRemaining.Caption = "Wait: " & DISPLAY_DURATION - m_counter & " seconds."
  End If
End Sub
'#################################################################################
'showing this form when user clicks on About menu item
Public Sub showAsAboutDialog()
  m_showingAsAbout = True
  Me.ProgressBar1.Visible = False
  Me.lblRemaining.Visible = False
  Me.Height = 3405
  Me.Show
End Sub
'#################################################################################
