
' // |new extension - creates a new file

Function pbx_Parse(Args)
Dim fso
Dim ftf

  If Args.Count Then
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ftf = fso.CreateTextFile(Args.Item(1), False)
    ftf.Close
    pbx_Parse = quote(Args.Item(1))

  Else
    pbx_Parse = "!Arg missing"

  End If

'  pbx_Parse = "You entered " & Args.Count & " args"

End Function

Function pbx_Hint()

  pbx_Hint = "Creates a new file"

End Function
