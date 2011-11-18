VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WM_SYSKEYDOWN = &H104

Const CLASS_NAME = "w>phatBarMk2"

Const BAR_HEIGHT = 36
Const TEXT_HEIGHT = 14
Const ICON_SIZE = 16
Const MARGIN = 8
Const GRIPPER_WIDTH = ICON_SIZE + MARGIN + MARGIN

Dim mhWnd As Long
Dim mhWndText As Long
Dim mView As mfxView
Dim mhFont As Long
Dim mhKey As Long
Dim mWidth As Long

Dim mDoingSysMenu As Boolean
Dim mMenuOpen As Boolean

Dim mConfig As CConfFile3
Dim mShowingError As Boolean        ' // next keypress clears content

    ' /* config */
Dim mColBackground As Long
Dim mTooltip As CToolTip
Dim mShading As Boolean
Dim mAutoHide As Boolean

    ' /* lists */
Dim mHistory As BTagList
Dim mHistoryIndex As Long
Dim mExtensions As BTagList
Dim mAddons As BTagList
Dim mAlias As CConfSection

Dim mIcon As mfxBitmap
Dim mToolIcon As mfxBitmap
Dim mToolRect As BRect
Dim mToolButtonActive As Boolean

Dim mAction As String

Dim WithEvents theItemList As frmItems
Attribute theItemList.VB_VarHelpID = -1

Implements IDropTarget
Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean

    If hWnd = mhWndText Then
        BWndProcSink_WndProc = uWndProcText(hWnd, uMsg, wParam, lParam, ReturnValue)

    Else
        BWndProcSink_WndProc = uWndProc(hWnd, uMsg, wParam, lParam, ReturnValue)

    End If

End Function

Private Sub Class_Initialize()

    ' /* read the config first off */
    uLoadConfig g_MakePath(App.Path) & "phatbar.conf"

    EZRegisterClass CLASS_NAME

    mhWnd = EZ4AddWindow(CLASS_NAME, Me, , WS_POPUP Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS, WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
    g_SizeWindow mhWnd, mWidth, BAR_HEIGHT
    g_MoveWindow mhWnd, Fix((g_ScreenWidth - mWidth) / 2), 0

    g_SetAlwaysOnTop mhWnd, True
    RegisterDragDrop mhWnd, Me

    mhKey = register_system_key(mhWnd, vbKeySpace, B_SYSTEM_KEY_WINDOWS)

    ' /* edit control */

    mhWndText = EZ4AddWindow(WC_EDIT, Nothing, , WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or WS_VISIBLE Or ES_AUTOHSCROLL, , mhWnd)
    g_SizeWindow mhWndText, mWidth - (GRIPPER_WIDTH * 2) - (4 * 2), TEXT_HEIGHT
    g_MoveWindow mhWndText, GRIPPER_WIDTH + 4, Fix((BAR_HEIGHT - TEXT_HEIGHT) / 2)
    window_subclass mhWndText, Me

    mhFont = g_CreateFont("Tahoma", 8)
    SendMessage mhWndText, WM_SETFONT, mhFont, ByVal 0&

    Set theItemList = New frmItems

    Set mTooltip = New CToolTip
    mTooltip.Create TTS_ALWAYSTIP, mhWnd
'    mTooltip.Add "tip", TTF_TRACK Or TTF_TRANSPARENT

    Set mHistory = new_BTagList()
'    Set mExtensions = new_BTagList()

    Set mToolRect = new_BRect(0, 0, ICON_SIZE - 1, ICON_SIZE - 1)
    mToolRect.OffsetBy mWidth - (ICON_SIZE + MARGIN) - 1, Fix((BAR_HEIGHT - ICON_SIZE) / 2)
    Set mToolIcon = load_image_obj(g_MakePath(App.Path) & "icons\sysmenu.png")

    mTooltip.Add "tool", TTF_TRANSPARENT, "Tool menu [Alt+C]", mToolRect
    mTooltip.Add "icon", TTF_TRANSPARENT, "Action icon", new_BRect(0, MARGIN, GRIPPER_WIDTH - 1, BAR_HEIGHT - MARGIN - 1)

    ' /* can specify custom configuration file */

'    If Command$ <> "" Then
'        mConfig = Command$
'
'    Else
'        mConfig =
'
'    End If

    uLoadExtensions
    uLoadAddons

    ' /* create view, organise item frames and draw */

    Set mView = New mfxView
'    uRedraw


    uResult "Welcome to phatBar!"
    uSetIcon "!null"

    mShowingError = True

    ' /* make visible */

    g_ShowWindow mhWnd, True
    SetFocusA mhWndText

End Sub

Private Sub Class_Terminate()

    window_subclass mhWndText, Nothing
    EZ4RemoveWindow mhWndText
    DeleteObject mhFont

    mTooltip.RemoveAll
    mTooltip.Destroy
    Set mTooltip = Nothing

    unregister_system_key mhWnd, mhKey
    RevokeDragDrop mhWnd
    EZ4RemoveWindow mhWnd
    EZUnregisterClass CLASS_NAME

End Sub

Public Function hWnd() As Long

    hWnd = mhWnd

End Function

Private Sub uRedraw()
Dim pr As BRect

    With mView

        Set pr = BW_Bounds(mhWnd)
        .SizeTo pr.Width, pr.Height

        .EnableSmoothing False
        .SetHighColour mColBackground
        .FillRect .Bounds

        ' /* shading */
        If mShading Then
            .SetHighColour rgba(0, 0, 0, 0)
            .SetLowColour rgba(0, 0, 0, 48)
            .FillRect .Bounds, MFX_VERT_GRADIENT

        End If

        ' /* border */
        .SetHighColour rgba(0, 0, 0, 106)
        .StrokeRect .Bounds

        ' /* high/lowlight */
        .SetHighColour rgba(255, 255, 255, 44)
        .SetLowColour rgba(0, 0, 0, 44)
        .StrokeFancyRect .Bounds.InsetByCopy(1, 1)

        ' /* edit field frame */
        Set pr = .Bounds
        pr.InsetBy GRIPPER_WIDTH, 3

        pr.InsetBy 1, 1

        .SetHighColour rgba(0, 0, 0, 24)
        .SetLowColour rgba(255, 255, 255, 24)
        .StrokeFancyRect pr

        pr.InsetBy 1, 1
        .SetHighColour rgba(255, 255, 255)
        .FillRect pr

        .SetHighColour rgba(0, 0, 0, 84)
        .StrokeRect pr

        ' /* gripper etching */
        
'        Set pr = .Bounds
''        pr.Right = pr.Left + GRIPPER_WIDTH - 1
'        pr.InsetBy 6, 6
'        pr.Right = pr.Left
'
'        .SetHighColour rgba(54, 54, 54, 88)
'        .StrokeLine pr
'        pr.OffsetBy 1, 0
'        .SetHighColour rgba(255, 255, 255, 88)
'        .StrokeLine pr
'
'        pr.OffsetBy 3, 0
'        .SetHighColour rgba(54, 54, 54, 88)
'        .StrokeLine pr
'        pr.OffsetBy 1, 0
'        .SetHighColour rgba(255, 255, 255, 88)
'        .StrokeLine pr

        ' /* action icon */

        .DrawScaledImage mIcon, new_BPoint(MARGIN + 1, Fix((BAR_HEIGHT - ICON_SIZE) / 2)), new_BPoint(ICON_SIZE, ICON_SIZE)

        ' /* tool button */

        If (mMenuOpen) And (mDoingSysMenu) Then
            .EnableSmoothing True
            .SetHighColour rgba(0, 0, 0, 20)
            .FillRoundRect mToolRect.InsetByCopy(-4, -4), 7, 7

            .SetHighColour rgba(0, 0, 0, 30)
            .StrokeRoundRect mToolRect.InsetByCopy(-4, -4), 7, 7

        End If
        
        .DrawScaledImage mToolIcon, mToolRect.TopLeft, new_BPoint(ICON_SIZE, ICON_SIZE)
        

'        If mToolButtonActive Then
'            .EnableSmoothing True
'            .SetHighColour rgba(0, 0, 0, 120)
'            .StrokeRoundRect mToolRect.InsetByCopy(-3, -3), 6, 6
'
'        End If

    End With

    g_UpdateWindow mhWnd

End Sub

'Private Sub uArrange()
'
'    If mItems.CountItems = 0 Then _
'        Exit Sub
'
'    mTooltip.RemoveAll
'
'Static i As Long
'Dim pr As BRect
'Dim pi As TItem
'
'    Set pr = new_BRect(0, 0, 23, 23)
'    pr.OffsetBy GRIPPER_WIDTH + 2, ((BAR_HEIGHT - ITEM_SIZE) / 2)
'
'    For i = 1 To mItems.CountItems
'        Set pi = mItems.TagAt(i)
'        Set pi.Frame = pr.Duplicate
'        mTooltip.Add CStr(i), TTF_TRANSPARENT Or TTF_TRACK, pi.Label, pr.Duplicate
'        pr.OffsetBy ITEM_SIZE + ITEM_GAP, 0
'
'    Next i
'
'End Sub

Private Function uHitTest(ByRef Point As BPoint) As Long

'    If mItems.CountItems = 0 Then _
'        Exit Function
'
'Static i As Long
'Dim pi As TItem
'
'    For i = 1 To mItems.CountItems
'        Set pi = mItems.TagAt(i)
'        If pi.Frame.Contains(Point) Then
'            uHitTest = i
'            Exit Function
'
'        End If
'
'    Next i

End Function

'Private Function uGetIcon(ByRef Item As TItem) As MImage
'Dim pi As BIcon
'
'    With new_BIconContent(Item.Target)
'        If .GetIcon(B_GET_ICON_BIGGEST Or B_GET_ICON_MOST_COLOURS, pi) Then _
'            Set uGetIcon = pi.Render()
'
'    End With
'
'End Function

Private Sub uWriteConfig()

    mConfig.Save

End Sub

Private Sub IDropTarget_DragEnter(ByVal pDataObject As olelib.IDataObject, ByVal grfKeyState As Long, ByVal ptx As Long, ByVal pty As Long, pdwEffect As olelib.DROPEFFECTS)

End Sub

Private Sub IDropTarget_DragLeave()

End Sub

Private Sub IDropTarget_DragOver(ByVal grfKeyState As Long, ByVal ptx As Long, ByVal pty As Long, pdwEffect As olelib.DROPEFFECTS)

End Sub

Private Sub IDropTarget_Drop(ByVal pDataObject As olelib.IDataObject, ByVal grfKeyState As Long, ByVal ptx As Long, ByVal pty As Long, pdwEffect As olelib.DROPEFFECTS)
Dim pDrop As CDropContent
Dim pData As CDropItem
Dim sz As String

    Set pDrop = New CDropContent
    If pDrop.SetTo(pDataObject) Then
        With pDrop
            Debug.Print .CountFormats
            .Rewind
            Do While .GetNextFormat(sz)
                Debug.Print ": " & sz
            
            Loop

'            If .HasFormat("FileNameW") Then
'                If .GetData("FileNameW", pData) Then
'                    Debug.Print "Added '" & pData.GetAsString(True) & "'"
'                    uAddByTarget pData.GetAsString(True)
'                    uWriteConfig
'                    uArrange
'                    uRedraw
'
'                End If
'
'            End If

        End With

    Else
        Debug.Print "IDropTarget.Drop(): failed to obtain IDataObject"

    End If

End Sub

Private Sub uLoadConfig(ByVal Path As String)

    ' /* defaults */

    mShading = True
    mColBackground = rgba(244, 244, 244)
    mAutoHide = True
    mWidth = 384

Dim ps As CConfSection
Dim sz As String
Dim i As Long

    Set mConfig = New CConfFile3
    Set mAlias = New CConfSection
    mAlias.SetName "aliases"

    With mConfig
        .SetFile Path
        If .Load() Then
            ' /* general section */
            If .Find("general", ps) Then
                i = g_RGBAFromString(ps.GetValueWithDefault("col-background"))
                If i <> 0 Then _
                    mColBackground = i

                If ps.Find("shading", sz) Then _
                    mShading = (sz = "1")

                If ps.Find("autohide", sz) Then _
                    mAutoHide = (sz = "1")

                If ps.Find("width", sz) Then
                    mWidth = g_SafeLong(sz)
                    If mWidth < 192 Then
                        mWidth = 192

                    ElseIf mWidth > (g_ScreenWidth - 160) Then
                        mWidth = g_ScreenWidth - 160

                    End If
                End If

'                If ps.Find("autofocus", sz) Then _
'                    mAutoFocus = (sz = "1")

            Else
                Debug.Print "[general] not found"

            End If

            ' /* aliases */

            If Not .Find("aliases", mAlias) Then
                Debug.Print "[aliases] not found"
                .Add mAlias
                .Save

            End If

        Else
            Debug.Print "no config, creating..."

            Set ps = New CConfSection
            With ps
                .SetName "general"
                .Add "col-background", CStr(get_red(mColBackground)) & "," & CStr(get_green(mColBackground)) & "," & CStr(get_blue(mColBackground))
                .Add "shading", IIf(mShading, "1", "0")
                .Add "autohide", IIf(mAutoHide, "1", "0")
    
            End With

            mConfig.Add ps
            mConfig.Add mAlias
            mConfig.Save

        End If

    End With

End Sub

Private Sub uDoMenu()
Dim pmi As OMMenuItem
Dim rc As RECT
Dim dw As Long
Dim pm As OMMenu

'        .AddItem .CreateItem("focs", "Auto focus", , , mAutoFocus)
'        .AddItem .CreateItem("relc", "Reload config")

    With New OMMenu

        .AddItem .CreateItem("info", "About phatBar...")

'        .AddItem .CreateItem("", "AddOns", , (pm.CountItems > 0), , , , pm)
        .AddSeparator

        .AddItem .CreateItem("hide", "Hide")


        Set pm = New OMMenu
        With pm
            .AddItem .CreateItem("ahde", "Auto hide", , , mAutoHide)
            .AddSeparator
            .AddItem .CreateItem("shad", "Shading", , , mShading)
            .AddItem .CreateItem("bcol", "Background colour...")
            .AddSeparator
            .AddItem .CreateItem("addn", "Reload AddOns")
            .AddItem .CreateItem("extn", "Reload Extensions")
        
        End With

        .AddItem .CreateItem("", "Settings", , , , , , pm)
        .AddSeparator
        .AddItem .CreateItem("quit", "Close")

        GetWindowRect mhWnd, rc

        mDoingSysMenu = True

        Set pmi = .Track(mhWnd, new_BPoint(rc.Right - MARGIN - ICON_SIZE - 3, rc.Bottom - MARGIN + 1))

        mDoingSysMenu = False

        If Not (pmi Is Nothing) Then
'            If g_BeginsWith(pmi.Name, "+") Then
'                If mAddons.Find(g_SafeRightStr(pmi.Name, Len(pmi.Name) - 1), pa) Then
'                    pa.Parse g_MakeArgList(g_WindowText(mhWndText))
'                    uSetAction g_WindowText(mhWndText)
'
'                End If
'                Exit Sub
'
'            End If

            Select Case pmi.Name
            Case "quit"
                uQuit

            Case "shad"
                mShading = Not mShading
                uUpdateGeneralSetting "shading", IIf(mShading, "1", "0")
                uRedraw

            Case "ahde"
                mAutoHide = Not mAutoHide
                uUpdateGeneralSetting "autohide", IIf(mAutoHide, "1", "0")

            Case "hide"
                g_ShowWindow mhWnd, False

'            Case "focs"
'                mAutoFocus = Not mAutoFocus
'                uWriteConfig

            Case "extn"
                uLoadExtensions

            Case "addn"
                uLoadAddons

            Case "info"
                MsgBox "phatBar " & CStr(App.Major) & "." & CStr(App.Minor) & IIf(App.Comments <> "", " " & App.Comments, "") & vbCrLf & _
                        App.LegalCopyright & vbCrLf & vbCrLf & "http://www.fullphat.net/", _
                        vbInformation Or vbOKOnly, _
                        App.Title

            Case "bcol"
                dw = g_MakeRGB24(mColBackground)
                If g_PickColour(dw, mhWnd) Then
                    mColBackground = g_MakeRGBA(dw)
                    uRedraw

                    uUpdateGeneralSetting "col-background", CStr(get_red(mColBackground)) & "," & CStr(get_green(mColBackground)) & "," & CStr(get_blue(mColBackground))

                End If

            End Select

        End If

    End With

End Sub

Private Sub uDoAddOnsMenu(ByVal Text As String)
Dim pm As OMMenu
Dim pa As TAddOn
Dim b As Boolean

    Set pm = New OMMenu
    With mAddons
        .Rewind
        Do While .GetNextTag(pa) = B_OK

            If Text = "" Then
                b = False

            ElseIf pa.Info.SupportsAnything Then
                b = True

            ElseIf mAction = "!url" Then
                b = pa.Info.SupportsURLs

            ElseIf mAction = "!file" Then
                b = pa.Info.SupportsFiles

            ElseIf mAction = "!folder" Then
                b = pa.Info.SupportsFolders

            Else
                b = False

            End If

            pm.AddItem pm.CreateItem(pa.Name, pa.Hint, , b)

        Loop

    End With

Dim pmi As OMMenuItem
Dim rc As RECT

    GetWindowRect mhWndText, rc
    Set pmi = pm.Track(mhWnd, new_BPoint(rc.Left, rc.Bottom + 2))
    If Not (pmi Is Nothing) Then
        If mAddons.Find(pmi.Name, pa) Then
            pa.Parse g_MakeArgList(g_WindowText(mhWndText))
            uSetAction

        End If

    End If

End Sub

Private Function uWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnValue As Long) As Boolean
Static rcwa As RECT
Static pwp As WINDOWPOS
Static fButtonPressed As Boolean
Static i As Long

Static pps As PAINTSTRUCT
Static hDC As Long

    Select Case uMsg

    Case WM_PAINT
        hDC = BeginPaint(hWnd, pps)
        draw_view mView, hDC
        EndPaint hWnd, pps
        ReturnValue = 0
        uWndProc = True

    Case WM_ERASEBKGND
        ReturnValue = -1
        uWndProc = True

    Case WM_MOUSEACTIVATE
        SetFocusA mhWndText
        ReturnValue = MA_ACTIVATE
        uWndProc = True

    Case WM_ACTIVATE
        Select Case LoWord(wParam)
        Case WA_INACTIVE
            If (mAutoHide) And (Not theItemList.Visible) Then _
                g_ShowWindow hWnd, False

        Case WA_CLICKACTIVE
            SetFocusA mhWndText

        End Select

'    Case WM_SETCURSOR
'        If (mAutoFocus) Then
'            g_WindowToFront hwnd, True
'            uSetExpanded True
'
'        End If


    Case WM_LBUTTONDBLCLK, WM_NCLBUTTONDBLCLK
        If g_IsIDE() Then _
            PostQuitMessage 0


    Case WM_MOUSEMOVE
        If mToolRect.Contains(new_BPointFromInt32(lParam)) Then
            If Not mToolButtonActive Then
                mToolButtonActive = True
                uRedraw
                SetTimer hWnd, 1, 50, 0

            End If
        End If


    Case WM_LBUTTONDOWN
        If mToolRect.Contains(new_BPointFromInt32(lParam)) Then
            SetCapture hWnd
            fButtonPressed = True

        End If


    Case WM_LBUTTONUP
        If fButtonPressed Then
            If mToolRect.Contains(new_BPointFromInt32(lParam)) Then _
                uDoMenu

            fButtonPressed = False

        End If


    Case WM_CAPTURECHANGED
        fButtonPressed = False


    Case WM_TIMER
        Select Case wParam
        Case 1
            If Not mToolRect.Contains(uToClient(hWnd, new_BPointFromInt32(GetMessagePos()))) Then
                KillTimer hWnd, wParam
                mToolButtonActive = False
                uRedraw

            End If
        
        End Select


'    Case WM_WINDOWPOSCHANGING
'        If fDragging Then
'            CopyMemory pwp, ByVal lParam, Len(pwp)
'            If pwp.x < rcwa.Left Then
'                pwp.x = rcwa.Left
'
'            ElseIf (pwp.x + pwp.cx) > rcwa.Right Then
'                pwp.x = rcwa.Right - pwp.cx
'
'            End If
'
'            If pwp.y < rcwa.Top Then
'                pwp.y = rcwa.Top
'
'            ElseIf (pwp.y + pwp.cy) > rcwa.Bottom Then
'                pwp.y = rcwa.Bottom - pwp.cy
'
'            End If
'
'            CopyMemory ByVal lParam, pwp, Len(pwp)
'            ReturnValue = 0
'            uWndProc = True
'
'        End If
'
'    Case WM_ENTERSIZEMOVE
'        fDragging = True
'        g_GetWorkArea rcwa
'
'    Case WM_EXITSIZEMOVE
'        fDragging = False
'        uWriteConfig


    Case WM_HOTKEY
        Debug.Print "HOTKEY"
        g_WindowToFront hWnd, True
        SetFocusA mhWndText


    Case WM_COMMAND
        Select Case HiWord(wParam)
        Case EN_CHANGE
            uSetAction

        End Select


    Case WM_SYSCOMMAND
        Select Case wParam
        Case SC_KEYMENU
            uDoMenu

        End Select


    Case WM_ENTERMENULOOP
        mMenuOpen = True
        uRedraw


    Case WM_EXITMENULOOP
        mMenuOpen = False
        uRedraw


    Case Else
'        Debug.Print uMsg

    End Select

End Function

Private Function uWndProcText(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef ReturnValue As Long) As Boolean

    Select Case uMsg

    Case WM_SYSKEYDOWN
        Select Case wParam
        Case vbKeyLeft
            If mWidth >= 192 Then _
                uChangeWindowSize -16

        Case vbKeyRight
            If mWidth <= g_ScreenWidth() - 160 Then _
                uChangeWindowSize 16

        End Select


    Case WM_KEYDOWN
'        Debug.Print wParam

        Select Case wParam
        Case vbKeyUp
            If mHistoryIndex > 0 Then
                SetWindowText hWnd, mHistory.TagAt(mHistoryIndex).Name
                edit_CursorToEnd hWnd
                mShowingError = False

            End If

            mHistoryIndex = mHistoryIndex - 1
            ReturnValue = 0
            uWndProcText = True

        Case vbKeyDown
            ReturnValue = 0
            uWndProcText = True

        Case vbKeyPageUp
            edit_CursorToStart hWnd
            ReturnValue = 0
            uWndProcText = True

        Case vbKeyPageDown
            edit_CursorToEnd hWnd
            ReturnValue = 0
            uWndProcText = True

        End Select


    Case WM_PASTE
        If mShowingError Then
            SetWindowText hWnd, ""
            mShowingError = False

        End If


    Case WM_CHAR
        If (mShowingError) And ((wParam > 26) Or (wParam = 8)) Then
            SetWindowText hWnd, ""
            mShowingError = False

        End If
        
        Select Case wParam
        Case 1
            edit_SelectAll hWnd

        Case 9
            If g_IsPressed(VK_SHIFT) Then
                uDoAddOnsMenu g_WindowText(mhWndText)

            Else
                uDoItemList

            End If

        Case 13
            uParse

        Case 27
            SetWindowText hWnd, ""

        Case Else
            Debug.Print wParam
            Exit Function

        End Select

        ReturnValue = 0
        uWndProcText = True

    End Select

End Function

Private Sub uParse()
Dim sz As String
Dim hr As Long

    sz = g_WindowText(mhWndText)
    If sz = "" Then _
        Exit Sub

    mHistory.Add new_BTagItem(sz, "")
    mHistoryIndex = mHistory.CountItems

    ' /* check for an alias */

    mAlias.Find sz, sz

    ' /* process */

Dim pc As Collection

    Select Case g_SafeLeftStr(sz, 1)
    Case "="
        ' /* maths */
        uCalc g_SafeRightStr(sz, Len(sz) - 1)

    Case "|"
        ' /* extension */
        uDoExtension g_SafeRightStr(sz, Len(sz) - 1)

    Case "?"
        ' /* phatBar command */
        uDo g_MakeArgList(g_SafeRightStr(sz, Len(sz) - 1))

    Case Else
        ' /* treat as file/folder */
        uRun g_MakeArgList(sz)

    End Select

End Sub

Private Sub uRun(ByRef Args As Collection)
Dim sz As String
Dim i As Long

    If Args.Count > 1 Then
        For i = 2 To Args.Count
            sz = sz & g_Quote(Args.Item(i)) & " "

        Next i

        sz = g_SafeLeftStr(sz, Len(sz) - 1)

    End If

'    MsgBox Args.Item(1)

Dim hr As Long

    hr = ShellExecute(mhWnd, vbNullString, Args.Item(1), sz, vbNullString, SW_SHOW)
    If hr < 33 Then _
        uError hr

End Sub

Private Sub uDo(ByRef Args As Collection)

    ' /* uDo - process a ?phatBar command
    '
    '   Commands
    '   ""          - displays version and copyright
    '   alias       - add an alias <name> <command>
    '   colr        - changes bar colour to a random one (undoc'd)
    '   extn        - reloads extensions
    '   hint        - displays the hint for the specified extension
    '   quit        - quits
    '   list        - lists items ("alias" or "extn")
    '
    ' */

    If Args.Count = 0 Then
        uResult "phatBar " & App.Major & "." & App.Minor & " (Build " & App.Revision & ") " & App.LegalCopyright, True
        Exit Sub

    End If

Dim pce As CConfEntry
Dim pe As TExtension
Dim sz As String

    Select Case LCase$(Args.Item(1))
    Case "quit"
        uQuit

    Case "extn"
        uLoadExtensions
        uResult "Extensions reloaded", True

    Case "list"
        If Args.Count = 2 Then
            Select Case Args.Item(2)
            Case "alias"
                With mAlias
                    .Rewind
                    Do While .NextEntry(pce)
                        sz = sz & pce.Name & " "
    
                    Loop
    
                End With

                sz = g_SafeLeftStr(sz, Len(sz) - 1)
                If sz = "" Then _
                    sz = "No aliases defined"


            Case "extn"
                With mExtensions
                    .Rewind
                    Do While .GetNextTag(pe) = B_OK
                        sz = sz & "|" & pe.Name & " "
    
                    Loop
    
                End With

                sz = g_SafeLeftStr(sz, Len(sz) - 1)
                If sz = "" Then _
                    sz = "No extensions loaded"

            Case Else
                uError "Invalid argument '" & Args.Item(2) & "'"

            End Select

            uResult sz, True

        Else
            uError "Usage: ?list <alias|extn>"

        End If


    Case "colr"
        Randomize Timer
        mColBackground = rgba((Rnd * 64) + 191, (Rnd * 64) + 191, (Rnd * 64) + 191)
        uResult "#" & g_HexStr(mColBackground, 6), True
        uRedraw


    Case "hint"
        If Args.Count = 2 Then
            If mExtensions.Find(Args.Item(2), pe, False) Then
                uResult pe.Name & ": " & pe.Hint, True

            Else
                uError "Unknown extension '" & Args.Item(2) & "'"

            End If

        Else
            uError "Usage: ?hint <extension>"

        End If


    Case "alias"
        If Args.Count = 1 Then
            ' /* list them */

            With mAlias
                .Rewind
                Do While .NextEntry(pce)
                    sz = sz & pce.Name & " "

                Loop

            End With

            sz = g_SafeLeftStr(sz, Len(sz) - 1)
            If sz = "" Then
                uError "No aliases defined"

            Else
                uResult sz, True

            End If

        ElseIf Args.Count = 2 Then
            ' /* list value of specified alias */
            If mAlias.Find(Args.Item(2), sz) Then
                uResult sz, , False

            Else
                uError "Unknown alias '" & Args.Item(2) & "'"

            End If

        ElseIf Args.Count = 3 Then
            If mAlias.IndexOf(Args.Item(2)) Then
                ' /* update existing */
                mAlias.Update Args.Item(2), g_Quote(Args.Item(3))
                uResult "Changed alias '" & Args.Item(2) & "'", True

            Else
                ' /* add new */
                mAlias.Add Args.Item(2), g_Quote(Args.Item(3))
                uResult "Added alias '" & Args.Item(2) & "'", True

            End If

            mConfig.Save

        Else
            uError "Usage: ?alias <name> <command>"

        End If


    Case Else
        uError "Unknown command '" & Args.Item(1) & "'"

    End Select

End Sub

Private Sub uDoExtension(ByVal str As String)
Dim cmd As String
Dim res As String
Dim pe As TExtension
Dim i As Long

    i = InStr(str, " ")
    If i = 0 Then
        cmd = str
        str = ""

    Else
        cmd = g_SafeLeftStr(str, i - 1)
        str = g_SafeRightStr(str, Len(str) - i)

    End If

    ' /* is it registered? */

    Debug.Print "uDoExtension(): command='" & cmd & "' args='" & str & "'"

    If mExtensions.Find(cmd, pe, False) Then
        Debug.Print "uDoExtension(): found in roster"
        If pe.Parse(str, res) Then
            uResult res

        Else
            uError res

        End If

    ElseIf uCreateExt(cmd, pe) Then
        Debug.Print "uDoExtension(): was created"
        If pe.Parse(str, res) Then
            uResult res
            
        Else
            uError res
            
        End If

    Else
        uError "Bad extension"

    End If

End Sub

Private Sub uError(ByVal str As String)

    SetWindowText mhWndText, "ERR: " & str
    mShowingError = True

End Sub

Private Sub uResult(ByVal str As String, Optional ByVal ClearOnNextKeyPress As Boolean, Optional ByVal SetTickIcon As Boolean = True)

    SetWindowText mhWndText, str
    SendMessage mhWndText, EM_SETSEL, 16383, ByVal 16383&

    If SetTickIcon Then _
        uSetIcon IIf(ClearOnNextKeyPress, "!info", "!ok")

    mShowingError = ClearOnNextKeyPress

End Sub

Private Sub uCalc(ByVal Expression As String)
Dim sz As String

    On Error Resume Next

    Err.Clear
    With New ScriptControl
        .Language = "VBScript"
        .AddObject "helpers", New THelpers, True
        sz = .Eval(Expression)

    End With

    If Err.Number = 0 Then
        uResult "=" & sz

    Else
        uError Err.Description

    End If

End Sub

Private Sub uAddExtension(ByVal Name As String, ByVal Hint As String, ByRef Handler As BObject)
Dim pe As TExtension

    Set pe = New TExtension
    pe.Init Name, Hint, Handler
    mExtensions.Add pe

End Sub

Private Function uCreateExt(ByVal Name As String, ByRef Extension As TExtension) As Boolean
Dim po As BObject

    On Error Resume Next

    Err.Clear
    Set po = CreateObject(Name & ".phatbar")
    If (Err.Number = 0) And (Not (po Is Nothing)) Then
        Debug.Print "extension '" & Name & "' created ok"
        Set Extension = New TExtension
        Extension.Init Name, "", po
        uCreateExt = True

    Else
        Debug.Print "uCreateExt(): '" & Err.Description & "' creating '" & Name & ".phatbar'"

    End If

End Function

Private Sub uSetIcon(ByVal Name As String)

    Debug.Print "setting action/icon to '" & Name & "'..."

    mAction = Name

    If g_BeginsWith(Name, "!") Then
        Set mIcon = load_image_obj(g_MakePath(App.Path) & "icons\" & g_SafeRightStr(Name, Len(Name) - 1) & ".png")

    Else
        Set mIcon = load_image_obj(Name)

    End If

    uRedraw

End Sub

Private Sub uSetAction()
Dim szTxt As String
Dim pc As Collection
Dim sz As String

    szTxt = g_WindowText(mhWndText)

    mAlias.Find szTxt, szTxt

    ' /* pre-defined */

    If szTxt = "" Then
        sz = "!null"

    ElseIf g_BeginsWith(szTxt, "?") Then
        ' /* pb command */
        sz = "!null"

    ElseIf g_BeginsWith(szTxt, "ERR:") Then
        sz = "!fail"

    ElseIf g_BeginsWith(szTxt, "=") Then
        sz = "!calc"

    ElseIf g_IsURL(szTxt) Then
        sz = "!url"

    ElseIf uIsExtension(szTxt, sz) Then

'    ElseIf mAlias.Find(sztxt, "") Then
'        sz = "!alias"

    Else
        ' /* lastly, check to see if it's a file or folder */

        sz = "!null"

        Set pc = g_MakeArgList(szTxt)
        If pc.Count > 0 Then
            If uExecutableExists(pc.Item(1)) Then
                sz = "!app"

            ElseIf g_Exists(pc.Item(1)) Then
                sz = IIf(g_IsFolder(pc.Item(1)), "!folder", "!file")

            End If
        End If
    End If

    Debug.Print "action: " & sz

    If sz <> mAction Then _
        uSetIcon sz

End Sub

Private Function uIsExtension(ByVal Text As String, ByRef Icon As String) As Boolean

    uIsExtension = g_BeginsWith(Text, "|")
    If Not uIsExtension Then _
        Exit Function

Dim sz As String
Dim i As Long

    Icon = "!ext-unknown"

    i = InStr(Text, " ")
    If i Then
        sz = g_SafeLeftStr(Text, i - 1)

    Else
        sz = Text

    End If

    sz = g_SafeRightStr(sz, Len(sz) - 1)
    If sz = "" Then
        Icon = "!ext"
        Exit Function

    End If

Dim pe As TExtension

    If mExtensions.Find(sz, pe, False) Then _
        Icon = pe.Icon

End Function

Private Sub uLoadExtensions()

    Set mExtensions = new_BTagList()

    uAddExtension "test", "Dummy extension", New XTest
    uAddExtension "snarl", "Snarl", New XSnarl

    uGetScriptExtensions

End Sub

Private Sub uGetScriptExtensions()
Dim pe As TExtension
Dim sz As String

    With New CFolderContent2
        If .SetTo(g_MakePath(App.Path) & "extensions") Then
            .Rewind
            Do While .GetNextFile(sz)
                If uIsValidScriptExtension(sz) Then
                    Set pe = New TExtension
                    If pe.InitAsScript(sz) Then _
                        mExtensions.Add pe

                End If
            Loop
        End If
    End With

End Sub

Private Function uIsValidScriptExtension(ByVal Path As String)

    uIsValidScriptExtension = g_IsFolder(Path)

End Function

Private Sub uLoadAddons()

    Set mAddons = new_BTagList()
    uGetScriptAddons

End Sub

Private Sub uGetScriptAddons()
Dim pa As TAddOn
Dim sz As String

    With New CFolderContent2
        If .SetTo(g_MakePath(App.Path) & "addons") Then
            .Rewind
            Do While .GetNextFile(sz)
                If uIsValidScriptAddon(sz) Then
                    Set pa = New TAddOn
                    If pa.InitAsScript(sz) Then _
                        mAddons.Add pa

                End If
            Loop
        End If
    End With

End Sub

Private Function uIsValidScriptAddon(ByVal Path As String)

    uIsValidScriptAddon = g_IsFolder(Path)

End Function



Private Function uToClient(ByVal hWnd As Long, ByRef aPoint As BPoint) As BPoint
Static p As POINTAPI

    p.x = aPoint.x
    p.y = aPoint.y
    ScreenToClient hWnd, p
    Set uToClient = new_BPoint(p.x, p.y)

End Function

Private Sub uUpdateGeneralSetting(ByVal Name As String, ByVal Value As String)
Dim i As Long

    With mConfig
        i = .IndexOf("general")
        If i Then
            .SectionAt(i).Update Name, Value
            uWriteConfig

        End If

    End With

End Sub

Private Function uExecutableExists(ByVal Name As String) As Boolean

    If g_GetExtension(Name, True) <> "exe" Then _
        Name = Name & ".exe"

Dim sz As String

    ' /* check %system% */

    sz = g_GetSystemFolderStr(CSIDL_SYSTEM)
    If g_Exists(g_MakePath(sz) & Name) Then
        uExecutableExists = True
        Exit Function

    End If

    ' /* check HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths */

Dim hKey As Long

    If reg_OpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & Name, hKey) Then
        If reg_GetValue(hKey, "", sz) Then _
            uExecutableExists = g_Exists(sz)

        RegCloseKey hKey

    End If

End Function

Private Sub uChangeWindowSize(ByVal Increment As Long)
Dim rc As RECT

    mWidth = mWidth + Increment
    uUpdateGeneralSetting "width", CStr(mWidth)

    GetClientRect mhWndText, rc
    g_SizeWindow mhWndText, (rc.Right - rc.Left) + Increment, TEXT_HEIGHT

    mToolRect.OffsetBy Increment, 0
    g_SizeWindow mhWnd, mWidth, BAR_HEIGHT
    uRedraw

    g_MoveWindow mhWnd, Fix((g_ScreenWidth - mWidth) / 2), 0

End Sub

Private Sub theItemList_Closed()

    SetFocusA mhWndText
    uSetAction

End Sub

Private Sub theItemList_Invoked(ByVal Item As String)

    uParse

End Sub

Private Sub theItemList_ItemSelected(ByVal Item As String)

    SetWindowText mhWndText, g_Quote(Item)

End Sub

Private Sub theItemList_OpenSysMenu()

    uDoMenu

End Sub

Private Sub uDoItemList()

    If theItemList.Visible Then
        theItemList.Quit
        Exit Sub

    End If

Dim sz As String

    sz = g_WindowText(mhWndText)

    Select Case mAction
    Case "!file"
        theItemList.Go mhWnd, g_GetPath(sz)

    Case "!folder"
        theItemList.Go mhWnd, sz

    Case Else
        uSetIcon "!search"
        theItemList.Go mhWnd, "", sz

    End Select

End Sub

Private Sub uQuit()

    Unload theItemList
    PostQuitMessage 0

End Sub

Public Function BackgroundColour() As Long

    BackgroundColour = mColBackground

End Function
