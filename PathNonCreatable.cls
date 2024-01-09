VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PathNonCreatable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BACKSLASH As String = "\"
Private Const SLASH As String = "/"
Private Const IOEXCEPTION As Long = 101

Private Type TConfig
    ErrorGuard As ErrorGuard
    Fso As FileSystemObject
End Type
Private this As TConfig

Public Function ConcatenatePath(ParamArray pathElements() As Variant) As String
    On Error GoTo ErrHandler

    Dim sPath As String

    Dim i As Long: For i = LBound(pathElements) To UBound(pathElements)

        sPath = this.Fso.BuildPath(sPath, CStr(pathElements(i)))

    Next

    ConcatenatePath = sPath

Exit Function
ErrHandler:
    this.ErrorGuard.RaiseGuard "CreateDirectory"
End Function


Public Sub CreateDirectory(ByVal folderPath As String)
    On Error GoTo ErrHandler

    folderPath = Trim$(folderPath)
    If folderPath = vbNullString Then Exit Sub

    If this.Fso.GetDriveName(folderPath) = vbNullString Then
        Debug.Print "Path: " & folderPath & " does not start with drive"
        Err.Raise IOEXCEPTION, , "Path: " & folderPath & " does not start with drive"
        Exit Sub
    End If

    'If Me.HasExctension(folderPath) Then folderPath = Left(folderPath, InStrRev(folderPath, BACKSLASH) - 1)

    Dim directories() As String: directories = Split(folderPath, BACKSLASH)
    Dim BuildPath As String

    Dim i As Long: For i = LBound(directories) To UBound(directories)

        If i > 0 And Not Me.IsFolderNameValid(directories(i)) Then
            Debug.Print directories(i) & " contains illegal chars!"
            Exit Sub
        End If

        BuildPath = this.Fso.BuildPath(BuildPath, directories(i) & Application.PathSeparator)
        If Not this.Fso.FolderExists(BuildPath) Then this.Fso.CreateFolder BuildPath

    Next

Exit Sub
ErrHandler:
    this.ErrorGuard.RaiseGuard "CreateDirectory"
End Sub

Public Function IsFolderNameValid(folderName As String) As Boolean
    On Error GoTo ErrHandler

    Dim invalidCharsPattern As String: invalidCharsPattern = InvalidCharsToRegexPattern

    Dim oRegex As RegExp: Set oRegex = New RegExp
    With oRegex
        .Pattern = "[" & invalidCharsPattern & "]"
        .Global = True
        IsFolderNameValid = Not .Test(folderName)
    End With

Exit Function
ErrHandler:
    this.ErrorGuard.RaiseGuard "IsFolderNameValid"
End Function

Public Function RemoveInvalidCharsFromFileName(ByVal FileName As String, Optional replaceInvalidChars As Boolean = False, Optional replacement As String = "_") As String
    On Error GoTo ErrHandler

    FileName = Trim$(FileName)

    Dim invalidCharsPattern As String: invalidCharsPattern = InvalidCharsToRegexPattern

    Dim oRegex As RegExp: Set oRegex = New RegExp
    With oRegex
        .Global = True
        .IgnoreCase = True
        .Pattern = "[" & invalidCharsPattern & "]+"
        If replaceInvalidChars Then
            RemoveInvalidCharsFromFileName = .Replace(FileName, replacement)
        Else
            RemoveInvalidCharsFromFileName = .Replace(FileName, vbNullString)
        End If
    End With

Exit Function
ErrHandler:
    this.ErrorGuard.RaiseGuard "RemoveInvalidCharsFromFileName"
End Function

Public Function HasExctension(ByVal Path As String) As Boolean
    On Error GoTo ErrHandler

    Path = Trim$(Path)
    If Path = vbNullString Then Exit Function

    HasExctension = Not (this.Fso.GetExtensionName(Path) = vbNullString)

Exit Function
ErrHandler:
    this.ErrorGuard.RaiseGuard "HasExctension"
End Function

Public Function UriPathToLocal(Path As String) As String
    On Error GoTo ErrHandler

    Const LITERAL_DOCUMENTS As String = "/documents"

    Path = LCase$(Trim$(Path))
    If Path = vbNullString Then Exit Function
    If InStr(Path, "https://") = 0 Then
        UriPathToLocal = Path
        Exit Function
    End If

    'Comercial OneDrive has /Documents
    Dim PathOneDriveCommercial As String
    PathOneDriveCommercial = Mid$(Path, InStr(Path, LITERAL_DOCUMENTS) + Len(LITERAL_DOCUMENTS), Len(Path))
    PathOneDriveCommercial = Replace(PathOneDriveCommercial, SLASH, BACKSLASH)

    'TODO: find https path change to drive
    'UriPathToLocal = Replace(Left(Path, InStr(Path, LITERAL_DOCUMENTS) - 1), SLASH, BACKSLASH) & PathOneDriveCommercial


    UriPathToLocal = environ("ONEDRIVECOMMERCIAL") & PathOneDriveCommercial

Exit Function
ErrHandler:
    this.ErrorGuard.RaiseGuard "UriPathToLocal"
End Function

Private Function InvalidCharsToRegexPattern() As String

    Dim invalidChars As String: invalidChars = Join(GetInvalidFileNameChars, vbNullString)
    invalidChars = Replace(invalidChars, SLASH, BACKSLASH & SLASH)
    invalidChars = Replace(invalidChars, BACKSLASH, BACKSLASH & BACKSLASH)

    InvalidCharsToRegexPattern = invalidChars

End Function

Public Property Get GetInvalidFileNameChars() As String()

    Const invalidChars As String = "\,<,>,:,"",/,|,?,*"
    GetInvalidFileNameChars = Split(invalidChars, ",")

End Property

Private Sub Class_Initialize()
    Set this.Fso = New FileSystemObject
    Set this.ErrorGuard = New ErrorGuard
End Sub

Private Sub Class_Terminate()
    Set this.Fso = Nothing
    Set this.ErrorGuard = Nothing
End Sub

Public Property Get ErrorGuard() As ErrorGuard
    Set ErrorGuard = this.ErrorGuard
End Property

Public Property Set ErrorGuard(ByVal vNewValue As ErrorGuard)
    Set this.ErrorGuard = vNewValue
End Property

Public Property Get Fso() As FileSystemObject
    Set Fso = this.Fso
End Property


