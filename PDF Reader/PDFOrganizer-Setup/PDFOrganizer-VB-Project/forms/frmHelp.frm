VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   8730
   ClientLeft      =   -195
   ClientTop       =   975
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   10335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15255
      ExtentX         =   26908
      ExtentY         =   18230
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
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
  Me.WebBrowser1.Navigate (App.Path & "\help\load_pdf.html")
  Me.WindowState = 2
  Me.Refresh
  Me.Show
End Sub
