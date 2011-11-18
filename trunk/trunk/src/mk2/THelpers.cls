VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "THelpers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LockWorkStation Lib "user32.dll" () As Long

Public Function pi() As Double

    pi = 4 * Atn(1)

End Function

Public Function sqrt(ByVal Number As Double) As Double

    sqrt = Sqr(Number)

End Function

Public Function who_am_i() As String

    who_am_i = g_GetUserName()

End Function

Public Function computer_name() As String

    computer_name = g_GetComputerName()

End Function

Public Function is_computer_on() As Boolean

    is_computer_on = True

End Function

Public Function is_computer_on_fire() As Boolean

    is_computer_on_fire = False

End Function

Public Sub shutdown_dialog()

    SHShutdownDialog gWindow.hWnd

End Sub

Public Sub logoff_dialog()

    LogoffWindowsDialog gWindow.hWnd

End Sub

Public Function shell_execute(Optional ByVal Verb As String = "open", Optional ByVal File As String, Optional ByVal Args As String, Optional ByVal Dir As String, Optional ByVal WindowStyle As Long = SW_SHOW) As Boolean

    If File <> "" Then _
        shell_execute = (ShellExecute(gWindow.hWnd, Verb, File, Args, Dir, WindowStyle) > 31)

End Function

Public Sub lock_computer()

    LockWorkStation

End Sub

Public Function quote(ByVal str As String) As String

    quote = g_Quote(str)

End Function
