
' // |sys extension - system commands and info

Function pbx_Parse(Args)

  If Args.Count Then
    Select Case lCase(Args.Item(1))
    Case "shutdown"
      shutdown_dialog

    Case "logoff"
      logoff_dialog

    Case "whoami"
      pbx_Parse = who_am_i()

    Case "this"
      pbx_Parse = computer_name()

    Case Else
      pbx_Parse = "!Unknown command"

    End Select

  Else
    pbx_Parse = "!Arg missing"

  End If

End Function

Function pbx_Hint()

  pbx_Hint = "Provides system commands and information"

End Function
