
Sub pba_Init()

  AddOn.Hint = "Command prompt here"
  AddOn.SupportsFolders = True

End Sub

Function pba_Parse(Args)
Dim sz

  If Args.Count = 1 Then _
    sz = Args.Item(1)

  shell_execute "open", "cmd.exe", , sz

End Function