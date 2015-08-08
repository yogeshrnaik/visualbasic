VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InfoBean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'###################################################################################
Public m_fullpath As String
'###################################################################################
Public Sub init(p_path As String)
  m_fullpath = p_path
End Sub
'###################################################################################
Public Function getName() As String
  getName = ""
  'get file name from full path
  Dim index As Integer
  Dim sRev As String
  sRev = StrReverse(m_fullpath)
  If InStr(1, sRev, "/") = 1 Then
    sRev = Mid(sRev, 2, Len(sRev) - 1)
  End If
  index = InStr(1, sRev, "\") 'position of first / in reversed path
  'extract string upto this /
  If index > 0 Then
    getName = StrReverse(Mid(sRev, 1, index - 1))
  End If
End Function
'###################################################################################
Public Function getParentFolder() As String
  getParentFolder = ""
  'get parent folder path from full path
  Dim index As Integer
  Dim sRev As String
  sRev = StrReverse(m_fullpath)
  If InStr(1, sRev, "/") = 1 Then
    sRev = Mid(sRev, 2, Len(sRev) - 1)
  End If
  index = InStr(1, sRev, "\") 'position of first / in reversed path
  'extract string upto this /
  If index > 0 Then
    getParentFolder = StrReverse(Mid(sRev, index + 1, Len(sRev) - index))
  End If
End Function
'###################################################################################
