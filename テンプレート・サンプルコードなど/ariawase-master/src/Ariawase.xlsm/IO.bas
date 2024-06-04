Attribute VB_Name = "IO"
'''+----                                                                   --+
'''|                             Ariawase 0.9.0                              |
'''|                Ariawase is free library for VBA cowboys.                |
'''|          The Project Page: https://github.com/vbaidiot/Ariawase         |
'''+--                                                                   ----+
Option Explicit
Option Private Module

''' @seealso Scripting.FileSystemObject http://msdn.microsoft.com/ja-jp/library/cc409798.aspx
''' @seealso ADODB.Stream http://msdn.microsoft.com/ja-jp/library/cc364272.aspx

Public Enum TristateEnum
    UseDefault = -2
    True_ = -1
    False_ = 0
End Enum

Public Enum OpenFileEnum
    ForReading = 1
    ForWriting = 2
    ForAppending = 8
End Enum

Public Enum FileAttrEnum
    ReadOnly = 1
    Hidden = 2
    System = 4
    Volume = 8
    Directory = 16
    Archive = 32
    Alias = 64
    Compressed = 128
End Enum

Public Enum DriveTypeEnum
    Removable = 1
    Fixed = 2
    Network = 3
    CDROM = 4
    RAMDisk = 5
End Enum

Public Enum SpFolderEnum
    WindowsFolder = 0
    SystemFolder = 1
    TemporaryFolder = 2
End Enum

Public Enum StreamTypeEnum
    adTypeBinary = 1
    adTypeText = 2
End Enum

Public Enum LineSeparatorsEnum
    adCRLF = -1
    adCR = 13
    adLF = 10
End Enum

Public Enum StreamOpenOptionsEnum
    adOpenStreamUnspecified = -1
    adOpenStreamAsync = 1
    adOpenStreamFromRecord = 4
End Enum

Public Enum ConnectModeEnum
    adModeUnknown = 0
    adModeRead = 1
    adModeWrite = 2
    adModeReadWrite = 3
    adModeShareDenyRead = 4
    adModeShareDenyWrite = 8
    adModeShareExclusive = 12
    adModeShareDenyNone = 16
    adModeRecursive = &H400000
End Enum

Public Enum ObjectStateEnum
    adStateClosed = 0
    adStateOpen = 1
    adStateConnecting = 2
    adStateExecuting = 4
    adStateFetching = 8
End Enum

Public Enum SaveOptionsEnum
    adSaveCreateNotExist = 1
    adSaveCreateOverWrite = 2
End Enum

Public Enum StreamReadEnum
    adReadLine = -2
    adReadAll = -1
End Enum

Public Enum StreamWriteEnum
    adWriteChar = 0
    adWriteLine = 1
End Enum

Private xxFso As Object 'Is Scripting.FileSystemObject
Private xxMimeCharsets As Variant '(Of Array(Of String))

''' @return As Object Is Scripting.FileSystemObject
Public Property Get Fso() As Object
    If xxFso Is Nothing Then Set xxFso = CreateObject("Scripting.FileSystemObject")
    Set Fso = xxFso
End Property

''' @return As String
Public Property Get ExecPath() As String
    Dim app As Object: Set app = Application
    Select Case app.Name
        Case "Microsoft Word":   ExecPath = app.MacroContainer.Path
        Case "Microsoft Excel":  ExecPath = app.ThisWorkbook.Path
        Case "Microsoft Access": ExecPath = app.CurrentProject.Path
        Case Else: Err.Raise 17
    End Select
End Property

''' @return As Variant(Of Array(Of String))
Public Property Get MimeCharsets() As Variant
    If IsEmpty(xxMimeCharsets) Then
        Dim stdRegProv As Object: Set stdRegProv = CreateStdRegProv()
        stdRegProv.EnumKey HKEY_CLASSES_ROOT, "MIME\Database\Charset", xxMimeCharsets
    End If
    MimeCharsets = xxMimeCharsets
End Property

''' @param propType As Integer Is StreamTypeEnum
''' @param propCharset As String In MimeCharsets
''' @param propLineSeparator As Integer Is LineSeparatorsEnum
''' @return As Object Is ADODB.Stream
Public Function CreateAdoDbStream( _
    Optional ByVal propType As StreamTypeEnum = adTypeText, _
    Optional ByVal propCharset As String = "Unicode", _
    Optional ByVal propLineSeparator As LineSeparatorsEnum = adCRLF _
    ) As Object
    
    Set CreateAdoDbStream = CreateObject("ADODB.Stream")
    With CreateAdoDbStream
        .Charset = propCharset
        .LineSeparator = propLineSeparator
        .Type = propType
    End With
End Function

Public Sub CopyAdoDbStream(ByVal srcStream As Object, ByVal dstStream As Object)
    If TypeName(srcStream) <> "Stream" Then Err.Raise 13
    If srcStream.Type <> adTypeText Then Err.Raise 5
    If TypeName(dstStream) <> "Stream" Then Err.Raise 13
    If dstStream.Type <> adTypeText Then Err.Raise 5
    
    If srcStream.State = adStateClosed Then srcStream.Open
    If dstStream.State = adStateClosed Then dstStream.Open
    
    If srcStream.LineSeparator = dstStream.LineSeparator Then
        srcStream.CopyTo dstStream
    Else
        While Not srcStream.EOS: dstStream.WriteText srcStream.ReadText(adReadLine), adWriteLine: Wend
    End If
End Sub

Public Function BomSize(ByVal chrset As String) As Integer
    Select Case LCase(chrset)
        Case "utf-8":             BomSize = 3
        Case "utf-16", "unicode": BomSize = 2
        Case Else:                BomSize = 0
    End Select
End Function

Public Sub SaveToFileWithoutBom( _
    ByVal strm As Object, ByVal fpath As String, ByVal opSave As SaveOptionsEnum _
    )
    
    If TypeName(strm) <> "Stream" Then Err.Raise 13
    If strm.Type <> adTypeText Then Err.Raise 5
    
    Dim strmZ As Object: Set strmZ = CreateAdoDbStream(adTypeBinary)
    strmZ.Open
    
    Dim chrset As String: chrset = strm.Charset
    Dim lnsep As Integer: lnsep = strm.LineSeparator
    strm.Type = adTypeBinary
    strm.Position = BomSize(chrset)
    
    strmZ.Write strm.Read(adReadAll)
    strmZ.Position = 0
    strmZ.SaveToFile fpath, opSave
    strmZ.Close
    
    strm.Position = 0
    strm.Type = adTypeText
    strm.Charset = chrset
    strm.LineSeparator = lnsep
End Sub

Public Function IsPathRooted(ByVal fpath As String) As Boolean
    Dim s As String
    s = Left(fpath, 1)
    If s = "\" Or s = "/" Then
        IsPathRooted = True
        GoTo Escape
    End If
    s = Mid(fpath, 2, 1)
    If s = ":" Then
        IsPathRooted = True
        GoTo Escape
    End If
    IsPathRooted = False
    
Escape:
End Function

Public Function GetSpecialFolder(ByVal spFolder As Variant) As String
    If IsNumeric(spFolder) Then
        GetSpecialFolder = Fso.GetSpecialFolder(spFolder)
    ElseIf VarType(spFolder) = vbString Then
        GetSpecialFolder = Wsh.SpecialFolders(spFolder)
    Else
        Err.Raise 13
    End If
End Function

Public Function GetTempFilePath( _
    Optional ByVal tdir As String, Optional extName As String = ".tmp" _
    ) As String
    
    If StrPtr(tdir) = 0 Then tdir = Fso.GetSpecialFolder(TemporaryFolder)
    Do
        GetTempFilePath = Fso.BuildPath(tdir, Replace(Fso.GetTempName(), ".tmp", extName))
    Loop While Fso.FileExists(GetTempFilePath)
End Function

Public Function GetUniqueFileName( _
    ByVal fpath As String, Optional delim As String = "_" _
    ) As String
    
    Dim d As String: d = Fso.GetParentFolderName(fpath)
    Dim b As String: b = Fso.GetBaseName(fpath) & delim
    Dim x As String: x = "." & Fso.GetExtensionName(fpath)
    
    Dim n As Long: n = 0
    While Fso.FileExists(fpath)
        n = n + 1
        fpath = Fso.BuildPath(d, b & CStr(n) & x)
    Wend
    GetUniqueFileName = fpath
End Function

Public Function GetAllFolders(ByVal folderPath As String) As Variant
    Dim ret As Collection: Set ret = New Collection
    GetAllFoldersImpl folderPath, ret
    GetAllFolders = ClctToArr(ret)
End Function
Private Sub GetAllFoldersImpl(ByVal folderPath As String, ByVal ret As Collection)
    Dim d As Object: Set d = Fso.GetFolder(folderPath)
    
    Dim sd As Object
    For Each sd In d.SubFolders
        ret.Add sd.Path
        GetAllFoldersImpl sd.Path, ret
    Next
End Sub

Public Function GetAllFiles(ByVal folderPath As String) As Variant
    Dim ret As Collection: Set ret = New Collection
    GetAllFilesImpl folderPath, ret
    GetAllFiles = ClctToArr(ret)
End Function
Private Sub GetAllFilesImpl(ByVal folderPath As String, ByVal ret As Collection)
    Dim d As Object: Set d = Fso.GetFolder(folderPath)
    
    Dim fl As Object
    For Each fl In d.Files: ret.Add fl.Path: Next
    
    Dim sd As Object
    For Each sd In d.SubFolders: GetAllFilesImpl sd.Path, ret: Next
End Sub

Public Sub CreateFolderTree(ByVal folderPath As String)
    If Not Fso.DriveExists(Fso.GetDriveName(folderPath)) Then Err.Raise 5
    CreateFolderTreeImpl folderPath
End Sub
Private Sub CreateFolderTreeImpl(ByVal folderPath As String)
    If Fso.FolderExists(folderPath) Then GoTo Escape
    CreateFolderTreeImpl Fso.GetParentFolderName(folderPath)
    Fso.CreateFolder folderPath
    
Escape:
End Sub
