VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   6015
   ClientLeft      =   2040
   ClientTop       =   1035
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   7980
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnOk 
      BackColor       =   &H006ABCFF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      MaskColor       =   &H006ABCFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9600
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub init_form()
'  Me.WebBrowser1.Navigate App.path & "\help\load_pdf.html"
'  Me.WindowState = 2
'  Me.Refresh
'  Me.Show
End Sub
