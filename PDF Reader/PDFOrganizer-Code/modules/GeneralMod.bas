Attribute VB_Name = "GeneralMod"
'#####################################################################################
Public Const OFN_ALLOWMULTISELECT As Long = &H200
Public Const OFN_CREATEPROMPT As Long = &H2000
Public Const OFN_ENABLEHOOK As Long = &H20
Public Const OFN_ENABLETEMPLATE As Long = &H40
Public Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Public Const OFN_EXPLORER As Long = &H80000
Public Const OFN_EXTENSIONDIFFERENT As Long = &H400
Public Const OFN_FILEMUSTEXIST As Long = &H1000
Public Const OFN_HIDEREADONLY As Long = &H4
Public Const OFN_LONGNAMES As Long = &H200000
Public Const OFN_NOCHANGEDIR As Long = &H8
Public Const OFN_NODEREFERENCELINKS As Long = &H100000
Public Const OFN_NOLONGNAMES As Long = &H40000
Public Const OFN_NONETWORKBUTTON As Long = &H20000
Public Const OFN_NOREADONLYRETURN As Long = &H8000& '*see comments
Public Const OFN_NOTESTFILECREATE As Long = &H10000
Public Const OFN_NOVALIDATE As Long = &H100
Public Const OFN_OVERWRITEPROMPT As Long = &H2
Public Const OFN_PATHMUSTEXIST As Long = &H800
Public Const OFN_READONLY As Long = &H1
Public Const OFN_SHAREAWARE As Long = &H4000
Public Const OFN_SHAREFALLTHROUGH As Long = 2
Public Const OFN_SHAREWARN As Long = 0
Public Const OFN_SHARENOWARN As Long = 1
Public Const OFN_SHOWHELP As Long = &H10
Public Const OFN_ENABLESIZING As Long = &H800000
Public Const OFS_MAXPATHNAME As Long = 260
'#####################################################################################
Public Const OKAY = 0
Public Const ERROR_OCCURED = -1
Public Const NO_ORDER = 0
Public Const DESCENDING = 1
Public Const ASCENDING = 2
'#####################################################################################
Public Function getApplnName() As String
  getApplnName = "PDF Organizer"
End Function
'#####################################################################################
Public Function removeSpecialChars(sp_data As String, Optional sp_allowed_chars As String) As String
  If InStr(1, sp_allowed_chars, "´", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "´", "")
  End If
  If InStr(1, sp_allowed_chars, "`", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "`", "")
  End If
  If InStr(1, sp_allowed_chars, "~", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "~", "")
  End If
  If InStr(1, sp_allowed_chars, "!", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "!", "")
  End If
  If InStr(1, sp_allowed_chars, "@", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "@", "")
  End If
  If InStr(1, sp_allowed_chars, "#", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "#", "")
  End If
  If InStr(1, sp_allowed_chars, "$", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "$", "")
  End If
  If InStr(1, sp_allowed_chars, "%", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "%", "")
  End If
  If InStr(1, sp_allowed_chars, "^", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "^", "")
  End If
  If InStr(1, sp_allowed_chars, "&", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "&", "")
  End If
  If InStr(1, sp_allowed_chars, "*", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "*", "")
  End If
  If InStr(1, sp_allowed_chars, "(", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "(", "")
  End If
  If InStr(1, sp_allowed_chars, ")", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, ")", "")
  End If
  If InStr(1, sp_allowed_chars, "-", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "-", "")
  End If
  If InStr(1, sp_allowed_chars, "_", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "_", "")
  End If
  If InStr(1, sp_allowed_chars, "=", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "=", "")
  End If
  If InStr(1, sp_allowed_chars, "+", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "+", "")
  End If
  If InStr(1, sp_allowed_chars, "\", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "\", "")
  End If
  If InStr(1, sp_allowed_chars, "|", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "|", "")
  End If
  If InStr(1, sp_allowed_chars, "[", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "[", "")
  End If
  If InStr(1, sp_allowed_chars, "]", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "]", "")
  End If
  If InStr(1, sp_allowed_chars, "{", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "{", "")
  End If
  If InStr(1, sp_allowed_chars, "}", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "}", "")
  End If
  If InStr(1, sp_allowed_chars, ";", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, ";", "")
  End If
  If InStr(1, sp_allowed_chars, ":", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, ":", "")
  End If
  If InStr(1, sp_allowed_chars, "'", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "'", "")
  End If
  If InStr(1, sp_allowed_chars, """", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, """", "")
  End If
  If InStr(1, sp_allowed_chars, ",", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, ",", "")
  End If
  If InStr(1, sp_allowed_chars, "<", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "<", "")
  End If
  If InStr(1, sp_allowed_chars, ".", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, ".", "")
  End If
  If InStr(1, sp_allowed_chars, ">", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, ">", "")
  End If
  If InStr(1, sp_allowed_chars, "/", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "/", "")
  End If
  If InStr(1, sp_allowed_chars, "?", vbTextCompare) = 0 Then
    sp_data = Replace(sp_data, "?", "")
  End If
  removeSpecialChars = sp_data
End Function
'#####################################################################################
Public Function isSpecialChar(sp_data As String, Optional sp_allowed_chars As String) As Boolean
  If Len(sp_data) > 1 Then
    isSpecialChar = False
  ElseIf InStr(1, sp_data, "´", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "´", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "`", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "`", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "~", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "~", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "!", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "!", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "@", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "@", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "#", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "#", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "$", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "$", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "%", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "%", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "^", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "^", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "&", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "&", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "*", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "*", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "(", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "(", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, ")", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, ")", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "-", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "-", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "_", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "_", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "=", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "=", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "+", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "+", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "\", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "\", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "|", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "|", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "[", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "[", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "]", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "]", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "{", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "{", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "}", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "}", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, ";", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, ";", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, ":", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, ":", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "'", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "'", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, """", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, """", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, ",", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, ",", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "<", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "<", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, ".", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, ".", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, ">", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, ">", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "/", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "/", vbTextCompare) = 0 Then
    isSpecialChar = True
  ElseIf InStr(1, sp_data, "?", vbTextCompare) <> 0 And InStr(1, sp_allowed_chars, "?", vbTextCompare) = 0 Then
    isSpecialChar = True
  Else
    isSpecialChar = False
  End If
End Function
'#####################################################################################
Public Function doesFileExist(sp_Filepath As String, sp_Filename As String) As Boolean
  Dim sFullPath As String
  Dim fs As New FileSystemObject
  
  If InStr(1, StrReverse(sp_Filepath), "\") = 1 Then
    sFullPath = sp_Filepath & sp_Filename
  Else
    sFullPath = sp_Filepath & "\" & sp_Filename
  End If
  
  doesFileExist = fs.FileExists(sFullPath)
End Function
'#####################################################################################
'Public Sub sortFileDetails(ByRef p_fileDetails As Collection)
'  If p_fileDetails Is Nothing Or IsNull(p_fileDetails) Or p_fileDetails.Count = 0 Then
'    Exit Sub
'  End If
'  Dim i, j As Integer
'  For i = 1 To p_fileDetails.Count - 1
'    For j = (i + 1) To p_fileDetails.Count
'      Dim t_pdfDtl1 As PDFFileDetails
'      Dim t_pdfDtl2 As PDFFileDetails
'      Set t_pdfDtl1 = p_fileDetails.Item(i)
'      Set t_pdfDtl2 = p_fileDetails.Item(j)
'
'    Next
'  Next
'End Sub
'#####################################################################################

