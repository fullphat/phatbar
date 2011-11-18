Attribute VB_Name = "mMain"
Option Explicit

Public gWindow As TWindow

Public Sub Main()

    Set gWindow = New TWindow

    With New BMsgLooper
        .Run

    End With

    Set gWindow = Nothing

End Sub

Public Function g_MakeArgList(ByVal Command As String) As Collection

    On Error GoTo er

    Set g_MakeArgList = New Collection

    If (Command = "") Or (Command = Chr$(34)) Then _
        Exit Function

Dim fQuoted As Boolean
Dim sz As String
Dim c As Long
Dim i As Long
Dim a() As String

    For i = 1 To Len(Command)
        If Mid$(Command, i, 1) = Chr$(34) Then
            fQuoted = Not fQuoted
'            sz = sz & Chr$(34)

        ElseIf (Mid$(Command, i, 1) = " ") And (fQuoted) Then
            sz = sz & Chr$(255)

        Else
            sz = sz & Mid$(Command, i, 1)

        End If
    Next i

    ' /* now split the string on each space */

    a() = Split(sz, " ", , vbTextCompare)

    ' /* lastly, change all � markers back to spaces and add to collection */

    If UBound(a) > -1 Then
        For i = 0 To UBound(a)
            g_MakeArgList.Add Replace$(a(i), Chr$(255), " ", , , vbTextCompare)

        Next i

    End If

    Exit Function

er:

End Function


