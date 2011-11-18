
Sub pba_Init()

  AddOn.Hint = "Tell me what I typed"
  AddOn.SupportsAnything = True

End Sub

Function pba_Parse(Args)
Dim sz
Dim i

  If Args.Count Then
    For i = 1 To Args.Count
      sz = sz & Args.Item(i)
      If i < Args.Count Then _
	sz = sz & " "

    Next

  End If

  MsgBox "You typed '" & sz & "'"

End Function