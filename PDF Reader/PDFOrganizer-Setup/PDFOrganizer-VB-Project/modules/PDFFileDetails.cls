VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PDFFileDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'#####################################################################################
Public m_filename As String
Public m_filepath As String
Public m_article_title As String
Public m_journal_title As String
Public m_journal_subject As String
Public m_year As String
Public m_volume_no As String
Public m_page_no As String
Public m_first_author As String
Public m_notes As String
'#####################################################################################
Public Sub init(p_filename As String, p_filepath As String, _
                p_article_title As String, p_journal_title As String, _
                p_journal_subject As String, p_year As String, _
                p_volume_no As String, p_page_no As String, p_first_author As String, _
                p_notes As String)
  m_filename = p_filename
  m_filepath = p_filepath
  If Len(Trim(p_article_title)) = 0 Then p_article_title = "NIL"
  If Len(Trim(p_journal_title)) = 0 Then p_journal_title = "NIL"
  If Len(Trim(p_journal_subject)) = 0 Then p_journal_subject = "NIL"
  If Len(Trim(p_year)) = 0 Then p_year = "NIL"
  If Len(Trim(p_volume_no)) = 0 Then p_volume_no = "NIL"
  If Len(Trim(p_page_no)) = 0 Then p_page_no = "NIL"
  If Len(Trim(p_first_author)) = 0 Then p_first_author = "NIL"
  m_article_title = p_article_title
  m_journal_title = p_journal_title
  m_journal_subject = p_journal_subject
  m_year = p_year
  m_volume_no = p_volume_no
  m_page_no = p_page_no
  m_first_author = p_first_author
  m_notes = p_notes
End Sub
'#####################################################################################
Public Function save_in_database(ByRef sp_error As String) As Boolean
On Error GoTo Database_Saving_Error
  '--------------------------------------------------------------------------
  sp_error = ""
  '--------------------------------------------------------------------------
  DoEvents
  If (stopLoad) Then
    save_in_database = False
    sp_error = "File loading stopped by user."
    Exit Function
  End If
  '---------------------------------------------------------------------
  Dim sSQL As String
  sSQL = getInsertSQL()
  If DatabaseMod.executeUpdate(sp_error, sSQL) = GeneralMod.ERROR_OCCURED Then
    If InStr(1, sp_error, "-2147467259", vbTextCompare) <> 0 And _
         InStr(1, UCase(sp_error), _
               UCase("duplicate values in the index, primary key, or relationship"), vbTextCompare) <> 0 Then
      'any error other than primary key violations, is shown
      'if error is primary key violations, then that means the file is already loaded
      'so update the file details
      sSQL = getUpdateSQL()
      If DatabaseMod.executeUpdate(sp_error, sSQL) = GeneralMod.ERROR_OCCURED Then
        save_in_database = False
        Exit Function
      End If
    Else
      save_in_database = False
      Exit Function
    End If
  End If
  '--------------------------------------------------------------------------
  save_in_database = True
  Exit Function
  '--------------------------------------------------------------------------
Database_Saving_Error:
  sp_error = ErrorHandlingMod.createErrorMsg("Error while saving PDF File Details in database: ", _
                                             Err.Number, Err.Description)
 save_in_database = False
End Function
'#####################################################################################
Private Function getInsertSQL() As String
  Dim sSQL As String
  sSQL = "insert into PDFFileDetails ( filename, filepath, article_title, "
  sSQL = sSQL & "journal_title, journal_subject, volume_no, page_no, first_author, [year] ) "
  sSQL = sSQL & "values ( '" & Replace(m_filename, "'", "''") & "', '" & _
                          Replace(m_filepath, "'", "''") & "', '" & _
                          Replace(m_article_title, "'", "''") & "', '"
  sSQL = sSQL & Replace(m_journal_title, "'", "''") & "', '" & _
                Replace(m_journal_subject, "'", "''") & "', '" & _
                Replace(m_volume_no, "'", "''") & "', '"
  sSQL = sSQL & Replace(m_page_no, "'", "''") & "', '" & _
                Replace(m_first_author, "'", "''") & "', '" & _
                Replace(m_year, "'", "''") & "')"
  getInsertSQL = sSQL
End Function
'#####################################################################################
Private Function getUpdateSQL() As String
  Dim sSQL As String
  sSQL = "update PDFFileDetails "
  sSQL = sSQL & "set journal_title = '" & Replace(m_journal_title, "'", "''") & "', "
  sSQL = sSQL & "article_title = '" & Replace(m_article_title, "'", "''") & "', "
  sSQL = sSQL & "journal_subject = '" & Replace(m_journal_subject, "'", "''") & "', "
  sSQL = sSQL & "volume_no = '" & Replace(m_volume_no, "'", "''") & "', "
  sSQL = sSQL & "page_no = '" & Replace(m_page_no, "'", "''") & "', "
  sSQL = sSQL & "first_author = '" & Replace(m_first_author, "'", "''") & "', "
  sSQL = sSQL & "[year] = '" & Replace(m_year, "'", "''") & "', "
  sSQL = sSQL & "filepath = '" & Replace(m_filepath, "'", "''") & "' "
  sSQL = sSQL & "where filename = '" & m_filename & "' "
  'sSQL = sSQL & "and filepath = '" & m_filepath & "' "
  getUpdateSQL = sSQL
End Function
'#####################################################################################
Public Function update_in_database(ByRef sp_error As String) As Boolean
On Error GoTo Database_Saving_Error
  '--------------------------------------------------------------------------
  sp_error = ""
  '--------------------------------------------------------------------------
  Dim sSQL As String
  sSQL = getUpdateSQL()
  If DatabaseMod.executeUpdate(sp_error, sSQL) = GeneralMod.ERROR_OCCURED Then
    update_in_database = False
    Exit Function
  End If
  '--------------------------------------------------------------------------
  update_in_database = True
  Exit Function
  '--------------------------------------------------------------------------
Database_Saving_Error:
  sp_error = ErrorHandlingMod.createErrorMsg("Error while saving PDF File Details in database: ", _
                                             Err.Number, Err.Description)
 update_in_database = False
End Function
'#####################################################################################

