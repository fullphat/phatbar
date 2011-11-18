VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TAddOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mName As String         ' // folder name
Dim mPath As String         ' // full path
Dim mScript As Boolean      ' // script-based (always TRUE for now)
Dim mIcon As String         ' // not currently used
Dim mInfo As XAddOnInfo

Dim WithEvents theScript As ScriptControl
Attribute theScript.VB_VarHelpID = -1

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mName

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = mInfo.Hint

End Function

Public Function Parse(ByRef Args As Collection) As Boolean

    On Error Resume Next

    Err.Clear

    If NOTNULL(theScript) Then
        theScript.Run "pba_Parse", Args
        Parse = (Err.Number = 0)

    End If

End Function

Public Function InitAsScript(ByVal Path As String) As Boolean
Dim szScript As String

    On Error Resume Next

    szScript = g_MakePath(Path) & "script.vbs"
    If Not g_Exists(szScript) Then _
        Exit Function

    Set theScript = New ScriptControl
    theScript.Language = "VBScript"

Dim szCode As String
Dim sz As String
Dim i As Integer

    Debug.Print "loading script '" & szScript & "'..."

    i = FreeFile()

    Err.Clear
    Open szScript For Input As #i
    If Err.Number <> 0 Then _
        Exit Function

    Do While Not EOF(i)
        Line Input #i, sz
        szCode = szCode & sz & vbCrLf

    Loop

    Close #i

    If Not uSafeAddCode(szCode) Then
        Debug.Print "Syntax error in script code"
        Exit Function

    End If

    If (Not uHasFunc("pba_Parse")) Or (Not uHasProc("pba_Init")) Then
        Debug.Print "missing pba_Parse() or pba_Init()"
        Exit Function

    End If

    mScript = True

    mName = g_FilenameFromPath(Path)
    mPath = Path
    mIcon = g_MakePath(Path) & "icon.png"

    Set mInfo = New XAddOnInfo

    theScript.AddObject "helpers", New THelpers, True
    theScript.AddObject "AddOn", mInfo

    Err.Clear
    theScript.Run "pba_Init"
    Debug.Print Err.Description

    If Err.Number <> 0 Then
        Debug.Print "TAddOn.InitAsScript(): pba_Init() failed"
        Exit Function

    End If

    If (Not mInfo.SupportsFiles) And (Not mInfo.SupportsFolders) And (Not mInfo.SupportsURLs) And (Not mInfo.SupportsAnything) Then
        Debug.Print "TAddOn.InitAsScript(): must support at least one type"
        Exit Function

    End If

    Debug.Print "ok"
    InitAsScript = True

End Function

Public Function Icon() As String

    Icon = mIcon

End Function

Private Function uSafeAddCode(ByVal code As String) As Boolean

    On Error Resume Next

    Err.Clear
    theScript.AddCode code
    uSafeAddCode = (Err.Number = 0)

End Function

Private Function uHasProc(ByVal Name As String, Optional ByRef Index As Long) As Boolean

    If (theScript Is Nothing) Then
        Debug.Print "TAddOn.uHasProc(): script not loaded"
        Exit Function

    End If

    If (Name = "") Or (theScript.Procedures.Count = 0) Then
        Debug.Print "TAddOn.uHasProc(): bad arg"
        Exit Function

    End If

    Name = LCase$(Name)

Dim n As Long

    With theScript.Procedures
        If .Count > 0 Then
            For n = 1 To .Count
                If LCase$(.Item(n).Name) = Name Then
                    Index = n
                    uHasProc = True
                    Exit Function

                End If
            Next n
        End If
    End With

    Debug.Print "TAddOn.uHasProc(): '" & Name & "' not found"

End Function

Private Function uHasFunc(ByVal Name As String) As Boolean
Dim n As Long

    If uHasProc(Name, n) Then _
        uHasFunc = theScript.Procedures.Item(n).HasReturnValue

End Function

Private Function uCallFunc(ByVal Name As String, ParamArray Args() As Variant) As Variant

    On Error Resume Next

    If Not uHasFunc(Name) Then
        Debug.Print "TAddOn.uCallFunc(): '" & Name & "' not found"
        Exit Function

    End If

    Err.Clear
    If IsMissing(Args) Then
        uCallFunc = theScript.Run(Name)

    Else
        uCallFunc = theScript.Run(Name, Args)

    End If

    If Err.Number <> 0 Then _
        Debug.Print "TAddOn.uCallFunc(): '" & Name & "' failed: " & Err.Description

End Function

Private Function uCallParse(ByRef Args As Collection, ByRef Result As String) As Boolean

    On Error Resume Next

    If Not uHasFunc("pba_Parse") Then
        Debug.Print "TAddOn.uCallParse(): no pba_Parse()"
        Exit Function

    End If

Dim sz As String

    Err.Clear
    sz = theScript.Run("pba_Parse", Args)
    If Err.Number <> 0 Then
        Debug.Print "TAddOn.uCallParse(): failed: " & Err.Description
        Result = Err.Description
        Exit Function

    End If

'    If g_BeginsWith(sz, "!") Then
'        ' /* parsing returned an error */
'        Result = g_SafeRightStr(sz, Len(sz) - 1)
'
'    Else
'        ' /* parsing returned success */
'        Result = sz
'        uCallParse = True
'
'    End If

End Function

Public Function Name() As String

    Name = mName

End Function

Public Function Hint() As String

    Hint = IIf(mInfo.Hint = "", mName, mInfo.Hint)

End Function

Public Function Info() As XAddOnInfo

    Set Info = mInfo

End Function

'Public Function SupportsFiles() As Boolean
'
'    SupportsFiles = mInfo.SupportsFiles
'
'End Function
'
'Public Function SupportsFolders() As Boolean
'
'    SupportsFolders = mInfo.SupportsFolders
'
'End Function
