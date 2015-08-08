Attribute VB_Name = "LoadInfoMod"
'###################################################################################
Private m_infoIndex As Collection
'###################################################################################
Public Function loadInfo() As Boolean
On Error GoTo Hell
  Dim cmdArgs As Collection
  Set cmdArgs = getCommandLine()
  'iterate through command line parameters and write to index file
  Dim i As Integer
  Dim fs As New FileSystemObject
  Dim ts  As TextStream
  'write to file
  Set ts = fs.OpenTextFile(App.Path & "\InfoQuest.dat", ForAppending, True)
  For i = 1 To cmdArgs.Count
    If Len(Trim(cmdArgs.item(i))) > 0 Then
      ts.WriteLine cmdArgs.item(i)
    End If
  Next
  ts.Close
  Set ts = Nothing
  Set fs = Nothing
  loadInfo = True
  Exit Function
Hell:
  loadInfo = False
  MsgBox "Error occured while loading info: Error Number = " & Err.Number & vbCrLf & _
         "Error Description: " & Err.Description, vbCritical, getApplnName
End Function
'###################################################################################
Private Function getCommandLine() As Collection
   'Declare variables.
   Dim C, CmdLine, CmdLnLen, InArg, InQuotedArgs, i, currArg
   Dim argsPool As New Collection
   NumArgs = 0: InArg = False: InQuotedArgs = False
   'Get command line arguments.
   CmdLine = Command()
   CmdLnLen = Len(CmdLine)
   'Go thru command line one character at a time.
   For i = 1 To CmdLnLen
      C = Mid(CmdLine, i, 1)
      If (C <> """" And Not InQuotedArgs) Then
        'Test for space or tab.
        If (C = " " Or C = vbTab) Then
         'space or tab found. Test if already in argument.
          If Not InArg Then
            'New argument begins.
            InArg = True
          Else
            'completed args
            InArg = False
            argsPool.Add currArg
            currArg = ""
          End If
        Else 'neither space nor tab.
          If Not InArg Then InArg = True
        End If
        'Concatenate character to current argument.
        If C <> " " And C <> vbTab Then currArg = currArg & C
        If i = CmdLnLen Then
          argsPool.Add currArg
        End If
      Else 'found a double quote or already in quoted args
        'Test if already in quoted argument.
        If Not InQuotedArgs Then
          'New argument begins. Test for too many arguments.
           InQuotedArgs = True
           InArg = False
        ElseIf C = """" Then 'found a quote, it means the end of quoted args
          InQuotedArgs = False
          argsPool.Add currArg
          currArg = ""
        End If
        'Concatenate character to current argument.
        If C <> """" Then currArg = currArg & C
      End If
   Next i
   Set getCommandLine = argsPool
End Function
'###################################################################################
Public Function getIndex(Optional blnRefresh As Boolean = False) As Collection
  
  If Not m_infoIndex Is Nothing And Not blnRefresh Then
    Set getIndex = m_infoIndex
    Exit Function
  End If
  
  'read the index file and return the collection of all paths present in index file
  Dim fs1 As New FileSystemObject
  Dim ts1 As TextStream
  Dim infoIndex As New Collection
  If fs1.FileExists(App.Path & "\InfoQuest.dat") Then
    Set ts1 = fs1.OpenTextFile(App.Path & "\InfoQuest.dat", ForReading, False)
    'read file line by line and create collection
    Dim sPath As String
    While Not ts1.AtEndOfStream
      sPath = ts1.ReadLine
      If Not isPresentIn(sPath, infoIndex) Then
        infoIndex.Add sPath, sPath
      End If
'Hell:
'        If Err.Number = 457 Then
'          'ignore
'        End If
    Wend
  End If
  If Not ts1 Is Nothing Then ts1.Close
  Set ts1 = Nothing
  Set fs = Nothing
  Set m_infoIndex = infoIndex
  Set getIndex = infoIndex
End Function
'###################################################################################
Private Function isPresentIn(sFind As String, ByRef oFindIn As Collection) As Boolean
  If oFindIn Is Nothing Or IsNull(sFind) Then
    isPresentIn = False
  End If
  If oFindIn.Count = 0 Or Len(sFind) = 0 Then
    isPresentIn = False
  End If
  Dim i As Integer
  For i = 1 To oFindIn.Count
    If (oFindIn.item(i) = sFind) Then
      isPresentIn = True
      Exit Function
    End If
  Next
  isPresentIn = False
End Function
'###################################################################################
