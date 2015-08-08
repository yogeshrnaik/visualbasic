Attribute VB_Name = "StartupMod"
'###################################################################################
Public Sub Main()
  If Len(Command()) > 0 Then
    If (loadInfo()) Then
      MsgBox "Information added successfully.", vbInformation, getApplnName
    End If
  Else
    'show the main form
    frmSearchInfo.Show
  End If
End Sub
'###################################################################################

