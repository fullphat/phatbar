
Function pbx_Parse(Args)
Dim oItem
Dim oApp

  Set oApp = CreateObject("Outlook.Application")
  Set oItem = oApp.CreateItem(0) 'olMailItem

  With oItem

    If Args.Count > 0 Then _
      .To = Args.Item(1)

    If Args.Count > 1 Then _
      .Subject = Args.Item(2)

    .ReadReceiptRequested = False
'    .HTMLBody = "Message<BR>Line 2"

  End With

  oItem.Display

End Function
