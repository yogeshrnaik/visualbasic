VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PDFWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#####################################################################################
Public m_word As String
Public m_X As Single
Public m_Y As Single
Public m_FontSize As Single
'variable used to identify to which database field this word belongs
'i.e. whether this word is part of Article title or Journal title or etc.
Public m_field_name As String
'#####################################################################################
Public Sub init(p_word As String, p_X As Single, _
                p_Y As Single, p_FontSize As Single)
  m_word = p_word
  m_X = p_X
  m_Y = p_Y
  m_FontSize = p_FontSize
End Sub
'#####################################################################################


