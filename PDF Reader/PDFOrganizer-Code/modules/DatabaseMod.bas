Attribute VB_Name = "DatabaseMod"
'#####################################################################################
Private con As ADODB.Connection
'#####################################################################################
Public Function openConnection() As Boolean
On Error GoTo Hell
  If con Is Nothing Then
    Set con = New ADODB.Connection
  End If
  Dim conString As String
  If con.State <> 1 Then
    con.CursorLocation = ADODB.adUseClient
    conString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                App.Path & "\database\pdfFilesInfo.mdb;"
    con.Open conString
  End If
  If con.State = 1 Then
    openConnection = True
  Else
    openConnection = False
  End If
  Exit Function
Hell:
  MsgBox "Error has occured: Could not Open Connection to Database." & vbNewLine & _
         "Error no: " & Err.Number & vbNewLine & _
         "Error Description: " & Err.Description, vbCritical, getApplnName
  openConnection = False
End Function
'#####################################################################################
Public Function executeQuery(strQuery As String) As ADODB.Recordset
  Dim rs As ADODB.Recordset
  If openConnection Then
    Set rs = New ADODB.Recordset
    'rs.Open strQuery, con, ADODB.adOpenDynamic, ADODB.adLockReadOnly
    ''rs.Open strQuery, con, ADODB.adOpenStatic, ADODB.adLockOptimistic
        
    rs.Open strQuery, con
    Set executeQuery = rs
  Else
    Set executeQuery = Nothing
  End If
End Function
'#####################################################################################
Public Function executeUpdate(ByRef sp_error As String, sp_query As String) As Integer
On Error GoTo Hell
  sp_error = ""
  Dim recordsAffected As Integer
  Dim rs As New ADODB.Recordset
  If openConnection Then
    Set rs = con.Execute(sp_query, recordsAffected)
    executeUpdate = recordsAffected
  Else
    executeUpdate = 0
  End If
  Exit Function
Hell:
  sp_error = ErrorHandlingMod.createErrorMsg("Error while executing the query:" & vbCrLf & _
                                             sp_query & vbCrLf, Err.Number, Err.Description)
  executeUpdate = GeneralMod.ERROR_OCCURED
End Function
'#####################################################################################
Public Function beginTrasaction() As Boolean
    If openConnection Then
      con.BeginTrans
      beginTrasaction = True
    Else
      beginTrasaction = False
    End If
End Function
'#####################################################################################
Public Function commitTrasaction() As Boolean
    If openConnection Then
      con.CommitTrans
      commitTrasaction = True
    Else
      commitTrasaction = False
    End If
End Function
'#####################################################################################
Public Function rollback() As Boolean
    If openConnection Then
      con.RollbackTrans
      rollback = True
    Else
      rollback = False
    End If
End Function
'#####################################################################################
Public Function savePDFInfo(sp_article_title As String, sp_journal_title As String, _
                            sp_journal_subject As String, sp_year As String, _
                            sp_volume_no As String, sp_page_no As String, _
                            sp_first_author As String, sp_filename As String, _
                            sp_filepath As String) As Boolean
  
End Function
'#####################################################################################
Public Sub closeRecordSet(ByRef rs As ADODB.Recordset)
On Error GoTo Hell
    If Not IsNull(rs) Then
      rs.Close
    End If
    Set rs = Null
Hell:
End Sub
'#####################################################################################
Public Sub closeConnection(ByRef con As ADODB.Connection)
On Error GoTo Hell
    If Not IsNull(con) Then
      con.Close
    End If
    Set con = Null
Hell:
End Sub
'#####################################################################################


