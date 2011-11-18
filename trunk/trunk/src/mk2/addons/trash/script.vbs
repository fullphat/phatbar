
Sub pba_Init()

  AddOn.Hint = "Move to Trash"
  AddOn.SupportsFolders = True
  AddOn.SupportsFiles = True

End Sub

Function pba_Parse(Args)
Const RECYCLE_BIN = &HA

Dim oFolder
Dim oShell

  Set oShell = CreateObject("Shell.Application")
  Set oFolder = oShell.NameSpace(RECYCLE_BIN)

  If Args.Count = 1 Then _
    oFolder.MoveHere(Args.Item(1))

End Function