Attribute VB_Name = "ErrorHandlingMod"
'############################################################################################
Public Const LOG_FILE_NAME = "PDFOrganizer.log"
'############################################################################################
Public Function createErrorMsg(ByVal sp_errDescPrefix As String, _
                               ByVal lp_sysErrNum As Long, _
                               ByVal sp_sysErrDesc As String) As String
  Dim sMessage As String
  If sp_errDescPrefix <> "" Then
    sMessage = sMessage & sp_errDescPrefix & vbCrLf
  End If
  sMessage = sMessage & "Error Number : " & lp_sysErrNum & vbCrLf
  sMessage = sMessage & "Error Description : " & sp_sysErrDesc & vbCrLf
  createErrorMsg = sMessage
End Function
'############################################################################################
'function to log error message to LOG file and give customized error message to user
Public Function handleError(ByVal sp_errDescPrefix As String, _
                            ByVal lp_sysErrNum As Long, _
                            ByVal sp_sysErrDesc As String) As String
  Dim sError As String
  sError = createErrorMsg(sp_errDescPrefix, lp_sysErrNum, sp_sysErrDesc)
  writeToLogFile (sError)
  handleError = sError
  'Return getCustErrMsg(lp_sysErrNum, sp_sysErrDesc)
End Function
'############################################################################################
'function to give customized error message to user
Public Function getCustErrMsg(ByVal lp_sysErrNum As Long, _
                              ByVal sp_sysErrDesc As String) As String
  Dim sCustErrMsg As String
  Select Case lp_sysErrNum
    Case -2147217865
      'table or view does not exist
      sCustErrMsg = "The required database table not found."
    Case -2147467259
      If InStr(1, sp_sysErrDesc, "Data Source", vbTextCompare) <> 0 Or _
         InStr(1, sp_sysErrDesc, "TNS", vbTextCompare) <> 0 Or _
         InStr(1, sp_sysErrDesc, "Connect internal only", vbTextCompare) <> 0 Then
        'sCustErrMsg = "Data Source name not found and no default driver specified."
        sCustErrMsg = "Unable to connect to database."
      Else
        sCustErrMsg = "The system has encountered an unexpected error. Please contact the Help Desk for further assistance."
      End If
    Case 13
      'sCustErrMsg = "Type Mismatch."
      sCustErrMsg = "Incorrect Data."
    Case 91
      sCustErrMsg = "Object variable or with block variable not set."
    Case 429
      'sCustErrMsg = ActiveX component can't create object
      sCustErrMsg = "Required Object cannot be created."
    Case 94
      'sCustErrMsg = "Invalid use of Null."
      sCustErrMsg = "Object required."
    Case 424
      'sCustErrMsg = "Object required."
      sCustErrMsg = "Required object not found."
    Case 462
      sCustErrMsg = "The remote server machine does not exist or is unavailable."
    Case 3704
      'sCustErrMsg = "Operation is not allowed when the object is closed."
      sCustErrMsg = "Object not intialized properly."
    Case 3709
      'sCustErrMsg = "The Connection cannot be used to perform this operation. It is either closed or invalid in this context."
      sCustErrMsg = "Database connection is closed."
    Case 31001
      sCustErrMsg = "Out of memory."
    Case 70
      'this error occurs when you don't have permission to create object using Server.CreateObject()
      sCustErrMsg = "Permission Denied."
    Case Else
      sCustErrMsg = "The system has encountered an unexpected error. Please contact the Help Desk for further assistance."
  End Select
  getCustErrMsg = sCustErrMsg
End Function
'############################################################################################
'function to log error message to a LOG file (ChargeOff.log).
Public Function writeToLogFile(ByVal sp_message As String, _
                              Optional ByVal sp_log_filename As String = LOG_FILE_NAME, _
                              Optional ByVal bp_add_timestamp As Boolean = True) As Boolean
On Error GoTo WriteLog_Error:
  '-----------------------------------------------------------------------
  Dim sFilePath As String
  Dim fileObj As New FileSystemObject
  Dim ts As TextStream
  '-----------------------------------------------------------------------
  sFilePath = App.Path & "\" & sp_log_filename
  '-----------------------------------------------------------------------
  'adding a carriage return and line feed in case it is not there in the message being logged
'  If sp_message.LastIndexOf(vbLf) <> sp_message.length - 1 Then
'    sp_message = sp_message & vbCrLf
'  End If
  '-----------------------------------------------------------------------
  Set ts = fileObj.OpenTextFile(sFilename, ForAppending, True)
  'ts.Write is used instead of ts.WriteLine
  'because, error message contains vbCrLf at the end
  ts.Write ("************************************************************************" & vbCrLf)
  If bp_add_timestamp Then
    ts.Write (Now() & ": " & sErrorDesc)
  Else
    ts.Write (sErrorDesc)
  End If
  ts.Close
  Set ts = Nothing
  Set fileObj = Nothing
  '-----------------------------------------------------------------------
  'separate each entry in LOG file with a line of asterisks
  SWriter.WriteLine ("***********************************************************************")
  If bp_add_timestamp = True Then
    SWriter.Write (Now() & ": " & sp_message) 'write message to file without line feed
  Else
    SWriter.Write (sp_message)      'write message to file without line feed
  End If
  '-----------------------------------------------------------------------
  writeToLogFile = True
  Exit Function
  '-----------------------------------------------------------------------
WriteLog_Error:
  'ignore errors
  Set ts = Nothing
  Set fileObj = Nothing
  writeToLogFile = False
End Function
'############################################################################################


