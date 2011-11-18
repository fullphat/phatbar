
Function pbx_Parse(Args)
Dim sz

  If Args.Count > 0 Then _
    sz = Args.Item(1)

  shell_execute "open", "cmd.exe", , sz

  pbx_Parse = "Ok"

End Function

Function pbx_Hint()

  pbx_Hint = "Opens a command prompt, optionally at a particular path"

End Function

Function pba_Parse(Args)
Dim sz

End Function

Function pba_Hint()

  pba_Hint = "Command prompt here"

End Function
