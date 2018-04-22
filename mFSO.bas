Attribute VB_Name = "mFSO"
'---------------------------------------------------------------------------------------
' Module    : mFSO
' DateTime  : 12/16/2004 21:12
' Author    : Shane Mulligan
' Purpose   : Contains functions derived from file system object
'---------------------------------------------------------------------------------------

Option Explicit

Public fso As Object
Private SubObject As Object


Sub Init()

   Set fso = CreateObject("Scripting.FileSystemObject")

End Sub


Sub OverwriteTextFile(ByVal FILENAME As String, ByVal Text As String)

   On Error GoTo initfso
   Set SubObject = fso.CreateTextFile(FILENAME, True)
   SubObject.Write Text
   SubObject.Close
   
   Exit Sub
   
initfso:
   Init

End Sub

Function GetFolder(ByVal sFile As String) As String
    Dim lCount As Long
    
    For lCount = Len(sFile) To 1 Step -1
        If Mid$(sFile, lCount, 1) = "\" Then
            GetFolder = Left$(sFile, lCount)
            Exit Function
        End If
    Next
    GetFolder = vbNullString
    
End Function

'Function Drives() As Object
'Dim oDriveArray As Object
'Dim oDrive As Object
'
'   Set oDriveArray = fso.Drives
'
'   For Each oDrive In oDriveArray
'      Call FindSubFolders(oDrive.Path & "\") 'Find Folders in Drive
'   Next oDrive
'
'End Function

'Sub SubFolders(ByVal FolderPath As String)
'Dim oFolder As Object
'Dim oSubfolderArray As Object
'Dim oSubfolder As Object
'
'   Set oFolder = fso.GetFolder(FolderPath)
'   Set oSubfolderArray = oFolder.SubFolders 'Find sub folders
'
'   For Each oSubfolder In oSubfolderArray
'      DoEvents
'      Call Files(oSubfolder.Path) 'find files in dir with case sensitivity
'      Call SubFolders(oSubfolder.Path)  'Recall to get the sub folders :=)
'   Next oSubfolder
'
'End Sub

'Sub Files(ByVal FolderPath As String)
'Dim oFolder As Object
'Dim oFiles As Object
'Dim oFile As Object
'
'   Set oFolder = fso.GetFolder(FolderPath)
'   Set oFiles = oFolder.Files
'
'   For Each oFile In oFiles
'      'FILENAME = fso.getfilename(oFile.Path)
'      'DoEvents ' give the comp a break
'   Next oFile
'
'End Sub

Function ReadStr(ByVal FileNum As Integer) As String
' Read In A String
    Do                           ' Start A Loop
        Line Input #FileNum, ReadStr                   ' Read One Line
    Loop While Left$(ReadStr, 1) = "/" Or Len(Trim$(ReadStr)) = 0 ' See If It Is Worthy Of Processing
End Function

Function Tokenize(ByVal s As String, ByVal delim As String) As Variant
    Dim A() As String
    Dim i As Integer
    i = 0
    s = Trim$(s)
    Do While (InStr(s, delim) > 0)
        ReDim Preserve A(i)
        A(i) = Left$(s, InStr(s, delim) - 1)
        s = Trim$(Mid$(s, InStr(s, delim) + 1))
        i = i + 1
    Loop
    If Len(s) > 0 Then
        ReDim Preserve A(i)
        A(i) = s
    End If
        
    Tokenize = A()
End Function
