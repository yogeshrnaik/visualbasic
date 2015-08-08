Attribute VB_Name = "GeneralMod"
'#####################################################################################
Public Const OKAY = 0
Public Const ERROR_OCCURED = -1
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

