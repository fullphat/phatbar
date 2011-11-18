VERSION 5.00
Begin VB.Form frmItems 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   197
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   309
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2670
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4515
   End
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event ItemSelected(ByVal Item As String)
Public Event Invoked(ByVal Item As String)
Public Event OpenSysMenu()
Public Event Closed()

Dim mSearch As String
Dim mFiles As BTagList
Dim mPath As String

Dim mView As mfxView

Implements BWndProcSink

Public Sub Go(ByVal hWndBar As Long, ByVal Path As String, Optional ByVal SearchFor As String)
Dim rc As RECT
Dim cx As Long

    GetClientRect hWndBar, rc
    cx = rc.Right - rc.Left - 40
    g_SizeWindow Me.hWnd, cx, 200

    GetWindowRect hWndBar, rc
    g_MoveWindow Me.hWnd, rc.Left + 20, rc.Bottom

    cx = GetWindowLong(List1.hWnd, GWL_STYLE)
    SetWindowLong List1.hWnd, GWL_STYLE, cx And (Not WS_BORDER)

    List1.Move 6, 6, Me.ScaleWidth - 12, Me.ScaleHeight - 12

    uRedraw

    mSearch = SearchFor
    mPath = g_MakePath(g_RemoveQuotes(Path))
    uGetContent

    g_ShowWindow Me.hWnd, True, True
    If List1.Enabled Then _
        SetFocusA List1.hWnd

End Sub

Public Sub Quit()

    RaiseEvent Closed
    Me.Hide

End Sub

Private Function uGetContent() As Boolean
Dim wfd As WIN32_FIND_DATA_API
Dim hFind As Long
Dim sz As String

    Set mFiles = new_BTagList()
    List1.Clear

    If mSearch <> "" Then
        List1.AddItem "Searching for '" & mSearch & "'..."
        List1.Enabled = False
        Exit Function

    Else
        List1.Enabled = True

    End If

    hFind = FindFirstFile(mPath & "*.*", wfd)
    If hFind = INVALID_HANDLE_VALUE Then _
        Exit Function

    Do
        sz = g_TrimStr(wfd.cFileName)
        If (sz <> ".") And (sz <> "..") Then
            mFiles.Add new_BTagItem(sz, "")
            List1.AddItem sz

'                LSet .Info = wfd
        End If

    Loop While FindNextFile(hFind, wfd) <> 0

    FindClose hFind
    uGetContent = True

End Function

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static pps As PAINTSTRUCT
Static hDC As Long

    Select Case uMsg
    Case WM_ERASEBKGND
        ReturnValue = -1
        BWndProcSink_WndProc = True

    Case WM_PAINT
        hDC = BeginPaint(hWnd, pps)
        draw_view mView, hDC
        EndPaint hWnd, pps
        ReturnValue = 0
        BWndProcSink_WndProc = True

    End Select

End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim pt As BTagItem
Dim sz As String

    Select Case KeyAscii
    Case 27, 9
        Me.Quit

    Case 8
        sz = g_GetPathParent(mPath)
        If sz <> "" Then
            mPath = sz
            RaiseEvent ItemSelected(g_RemovePath(sz))
            uGetContent

        End If

    Case 13
        Set pt = mFiles.TagAt(List1.ListIndex + 1)
        If NOTNULL(pt) Then
            sz = mPath & pt.Name
            If g_IsFolder(sz) Then
                mPath = g_MakePath(sz)
                uGetContent

            Else
                RaiseEvent Invoked(sz)
                Me.Quit

            End If
        End If

    End Select

End Sub

Private Sub Form_Load()

    Set mView = New mfxView
'    window_subclass Me.hWnd, Me

End Sub

Private Sub Form_Paint()

    draw_view mView, Me.hDC

End Sub

Private Sub Form_Unload(Cancel As Integer)

'    window_subclass Me.hWnd, Nothing

End Sub

Private Sub List1_Click()
Dim pt As BTagItem

    Set pt = mFiles.TagAt(List1.ListIndex + 1)
    RaiseEvent ItemSelected(mPath & pt.Name)

End Sub

Private Sub uRedraw()
Dim pr As BRect

    With mView

        Set pr = BW_Bounds(Me.hWnd)
        .SizeTo pr.Width, pr.Height

        .EnableSmoothing False
        .SetHighColour gWindow.BackgroundColour
        .FillRect .Bounds

        ' /* shading */
'        If mShading Then
            .SetHighColour rgba(0, 0, 0, 0)
            .SetLowColour rgba(0, 0, 0, 48)
            .FillRect .Bounds, MFX_VERT_GRADIENT

'        End If

        ' /* high/lowlight */
        .SetHighColour rgba(255, 255, 255, 56)
        .SetLowColour rgba(0, 0, 0, 56)
        .StrokeFancyRect .Bounds.InsetByCopy(1, 1)

        ' /* border */
        .SetHighColour rgba(0, 0, 0, 176)
        .StrokeRect .Bounds

        ' /* list box frame */

        pr.InsetBy 4, 4
        .SetHighColour rgba(0, 0, 0, 24)
        .SetLowColour rgba(255, 255, 255, 24)
        .StrokeFancyRect pr

        pr.InsetBy 1, 1
        .SetHighColour rgba(0, 0, 0, 84)
        .StrokeRect pr

    End With

End Sub
