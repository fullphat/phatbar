VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mHandler As BObject
Dim mName As String
Dim mHint As String
Dim mPath As String
Dim mScript As Boolean
Dim mIcon As String

'Dim mIsExtension As Boolean
'Dim mIsAddOn As Boolean

Dim WithEvents theScript As ScriptControl
Attribute theScript.VB_VarHelpID = -1

Implements BTagItem

Private Function BTagItem_Name() As String

    BTagItem_Name = mName

End Function

Private Function BTagItem_Value() As String

    BTagItem_Value = mHint

End Function

Public Function Parse(ByVal Args As String, ByRef Result As String) As Boolean
Dim hr As B_STATUS

    On Error GoTo er

    If mScript Then
        Parse = uCallParse(g_MakeArgList(Args), Result)

    ElseIf NOTNULL(mHandler) Then
        hr = mHandler.Init(Args)
        Result = mHandler.Name
        Parse = (hr = B_OK)

    End If

    Exit Function

er:
    Result = Err.Description
    Parse = False

End Function

Public Sub Init(ByVal Name As String, ByVal Hint As String, ByRef Handler As BObject)

    mScript = False
    mName = Name
    mHint = Hint
    Set mHandler = Handler

End Sub

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

'    mIsExtension = uHasFunc("pbx_Parse")
'    mIsAddOn = uHasFunc("pba_Parse")

'    If (Not mIsExtension) And (Not mIsAddOn) Then
    If Not uHasFunc("pbx_Parse") Then
        Debug.Print "missing pbx_Parse()"
        Exit Function

    End If

    mScript = True

    mName = g_FilenameFromPath(Path)
    mPath = Path
    mIcon = g_MakePath(Path) & "icon.png"

    theScript.AddObject "helpers", New THelpers, True
    uGetHint

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
        Debug.Print "TScriptExtension.uHasProc(): script not loaded"
        Exit Function

    End If

    If (Name = "") Or (theScript.Procedures.Count = 0) Then
        Debug.Print "TScriptExtension.uHasProc(): bad arg"
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

    Debug.Print "TScriptExtension.uHasProc(): '" & Name & "' not found"

End Function

Private Function uHasFunc(ByVal Name As String) As Boolean
Dim n As Long

    If uHasProc(Name, n) Then _
        uHasFunc = theScript.Procedures.Item(n).HasReturnValue

End Function

Private Function uCallProc(ByVal Name As String, ParamArray Args() As Variant) As Boolean

    On Error Resume Next

    If Not uHasProc(Name) Then
        Debug.Print "TScriptExtension.CallProc(): '" & Name & "' not found"
        Exit Function

    End If

    Err.Clear
    If IsMissing(Args) Then
        theScript.Run Name

    Else
        theScript.Run Name, Args

    End If

    If Err.Number <> 0 Then _
        Debug.Print "TScriptExtension.CallProc(): '" & Name & "' failed: " & Err.Description

    uCallProc = (Err.Number = 0)

End Function

Private Function uCallParse(ByRef Args As Collection, ByRef Result As String) As Boolean

    On Error Resume Next

    If Not uHasFunc("pbx_Parse") Then
        Debug.Print "TExtension.uCallParse(): no pbx_Parse()"
        Exit Function

    End If

Dim sz As String

    Err.Clear
    sz = theScript.Run("pbx_Parse", Args)
    If Err.Number <> 0 Then
        Debug.Print "TExtension.uCallParse(): failed: " & Err.Description
        Result = Err.Description
        Exit Function

    End If

    If g_BeginsWith(sz, "!") Then
        ' /* parsing returned an error */
        Result = g_SafeRightStr(sz, Len(sz) - 1)

    Else
        ' /* parsing returned success */
        Result = sz
        uCallParse = True
    
    End If

End Function

'Public Function Procedures() As Procedures
'
'    If Not (theScript Is Nothing) Then _
'        Set Procedures = theScript.Procedures
'
'End Function

Public Function Name() As String

    Name = mName

End Function

Public Function Hint() As String

    If mHint = "" Then
        Hint = "No hint for '" & mName & "'"

    Else
        Hint = mHint

    End If

End Function

Private Sub uGetHint()
Dim sz As String

    On Error Resume Next

    If NOTNULL(theScript) Then
        If uHasFunc("pbx_Hint") Then
            Err.Clear
            sz = theScript.Run("pbx_Hint")
            If Err.Number <> 0 Then
                Debug.Print "TExtension.uGetHint(): failed: " & Err.Description

            Else
                mHint = sz

            End If
        
        Else
            Debug.Print "TExtension.uGetHint(): script has no pbx_Hint() function"

        End If

    Else

    End If

End Sub
