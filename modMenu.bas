Attribute VB_Name = "modMenu"
Option Explicit
'*********************
'* Change MenuColors *
'* by Scythe         *
'*********************

'********************************************************************************
'* Changes in V1.0.1                                                            *
'* Added Seperator / Hotkey / Disabled / Checked support                        *
'* Changed Subclass Method so it calls the Original Routines without callbyname *
'* Rewrote some Routines to increase the speed                                  *
'********************************************************************************


'Searched the whole internet for an simple way to change the Fontcolor in Menus
'After hours of searchin i decided to write my own routine
'Took parts from PSC, ActiveVB, vb@rchiv
'Most code was from the never working example Menu Subclassing by Viper Tec./ DoS
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=6891&lngWId=1
'You can do much more with this sample
'Change Fonts,Colors,Style....
'I only needed Colors

'How To Use
'Call the ChangeMenu Routine in Form_Load or similar
'PopUp Menus you wont show in menubar
'DONT SET THE VISIBLE TO FALSE for PopUp menus / This take them out of subclassing
'Just give them the Name MnuPop.....

'The subclass methode has an selfunhook routine so it should not crash by closing window in ide

Private Const DT_LEFT           As Long = &H0
Private Const DT_SINGLELINE     As Long = &H20
Private Const DT_TOP            As Long = &H0
Private Const DT_CALCRECT       As Long = &H400
Private Const DT_RIGHT          As Long = &H2
Private Const FF_SCRIPT         As Integer = 64
Private Const LF_FACESIZE       As Integer = 32
Private Const SYSTEM_FONT       As Integer = 13
Private Const TRANSPARENT       As Integer = 1
'Menu
Private Const MF_BYPOSITION     As Long = &H400
Private Const MF_OWNERDRAW      As Long = &H100
Private Const MF_BYCOMMAND      As Long = &H0
Private Const ODS_SELECTED      As Long = &H1
Private Const ODS_DISABLED      As Long = &H4
Private Const ODS_CHECKED       As Long = &H8
Private Const ODT_MENU          As Long = &H1
Private Const MIIM_STATE        As Long = &H1
Private Const MIIM_FTYPE        As Long = &H100
Private Const MFS_GRAYED        As Long = &H3
'Subclassing
Private Const GWL_WNDPROC       As Integer = -4
Private Const WM_DRAWITEM       As Long = &H2B
Private Const WM_MEASUREITEM    As Long = &H2C
Private Const WM_DESTROY        As Long = &H2
Private Const WM_CLOSE          As Long = &H10
'Drawing
Private Type RECT
    Left                            As Long
    Top                             As Long
    Right                           As Long
    Bottom                          As Long
End Type
Private Type DRAWITEMSTRUCT
    CtlType                         As Long
    CtlID                           As Long
    itemID                          As Long
    itemAction                      As Long
    itemState                       As Long
    hwndItem                        As Long
    hdc                             As Long
    rcItem                          As RECT
    itemData                        As Long
End Type
Private Type LOGFONT
    lfHeight                        As Long
    lfWidth                         As Long
    lfEscapement                    As Long
    lfOrientation                   As Long
    lfWeight                        As Long
    lfItalic                        As Byte
    lfUnderline                     As Byte
    lfStrikeOut                     As Byte
    lfCharSet                       As Byte
    lfOutPrecision                  As Byte
    lfClipPrecision                 As Byte
    lfQuality                       As Byte
    lfPitchAndFamily                As Byte
    lfFaceName(1 To LF_FACESIZE)    As Byte
End Type
'Menu
Private Type MENUITEMINFO
    cbSize                          As Long
    fMask                           As Long
    fType                           As Long
    fState                          As Long
    wID                             As Long
    hSubMenu                        As Long
    hbmpChecked                     As Long
    hbmpUnchecked                   As Long
    dwItemData                      As Long
    dwTypeData                      As String
    cch                             As Long
End Type
Private Type MENUINFO
    cbSize                          As Long
    fMask                           As MENUINFO_MASKS
    dwStyle                         As MENUINFO_STYLES
    cyMax                           As Long
    hbrBack                         As Long
    dwContextHelpID                 As Long
    dwMenuData                      As Long
End Type
Private Enum MENUINFO_STYLES
    MNS_NOCHECK = &H80000000
    MNS_MODELESS = &H40000000
    MNS_DRAGDROP = &H20000000
    MNS_AUTODISMISS = &H10000000
    MNS_NOTIFYBYPOS = &H8000000
    MNS_CHECKORBMP = &H4000000
End Enum
Private Enum MENUINFO_MASKS
    MIM_MAXHEIGHT = &H1
    MIM_BACKGROUND = &H2
    MIM_HELPID = &H4
    MIM_MENUDATA = &H8
    MIM_STYLE = &H10
    MIM_APPLYTOSUBMENUS = &H80000000
End Enum
'Subclassing
Private Type MEASUREITEMSTRUCT
    CtlType                         As Long
    CtlID                           As Long
    itemID                          As Long
    itemWidth                       As Long
    itemHeight                      As Long
    itemData                        As Long
End Type
'Intern Types and Variables
'==========================
Private Type MenuInfos
    MenuText                        As String       'Visible Text i.e. File
    PopUp                           As Boolean      'Is this an PopUp Menu
    Hotkey                          As String       'The Hotkey if there is any
End Type
Private MenuItems()             As MenuInfos
Private OldDS                   As DRAWITEMSTRUCT
Private MnuCtr                  As Long
Private ColFor                  As Long
Private ColBack                 As Long
Private ColSel                  As Long
Private ColSelTxt               As Long
Private ColDisTxt               As Long
Private PrevProc                As Long
Private m_hWnd                  As Long
Private SubContainer            As Object
Private Type POINTAPI
    x                               As Long
    y                               As Long
End Type
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetObjectAPIBynum Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByVal lpObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, MI As MENUINFO) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Sub ApiLine(lngHdc As Long, x As Long, y As Long, X1 As Long, Y1 As Long, DrawColor As Long, Optional lngDrawWidth As Long = 1)

Dim PT       As POINTAPI
Dim hObject  As Long
Dim hPen     As Long
Dim hBrush   As Long
Dim OldBrush As Long
Dim OldPen   As Long

    OldPen = SelectObject(lngHdc, hPen)
    OldBrush = SelectObject(lngHdc, hBrush)
    hPen = CreatePen(0, lngDrawWidth, DrawColor)
    hObject = SelectObject(lngHdc, hPen)
    MoveToEx lngHdc, x, y, PT
    LineTo lngHdc, X1, Y1
    SelectObject lngHdc, OldPen
    SelectObject lngHdc, OldBrush
    DeleteObject hObject
    DeleteObject hPen

End Sub

'Init the whole thing
Public Sub ChangeMenu(Frm As Object, Optional ByVal MnuForeColor As Long = &HFFFFFF, Optional ByVal MnuBackColor As Long = 0, Optional ByVal MnuSelectColor As Long = &HFF00FF, Optional ByVal MnuSelectTextColor As Long = &HFFFFFF, Optional ByVal MnuDisabledTextColor As Long = &HC0C0C0)

'Check if we are in IDE
'If yes give a warning that subclassing could crash IDE

    If InIDE Then
        If MsgBox("Warnig:" & vbNewLine & _
          "You are in IDE" & vbNewLine & _
          "Any error could crash the IDE" & vbNewLine & _
          vbNewLine & _
          "Are you sure you want to start subclassing ?" & vbNewLine & _
          "If you select No the Programm will start without COLORMENUS", vbCritical + vbYesNo, "Menu Subclassing") = vbNo Then
            Exit Sub
        End If
    End If
    'Get the Menu Cations and disable PopUpMenus
    'So we can call them thru CallByName function
    GetMenuData Frm
    'Save the Form Object for CallByName
    Set SubContainer = Frm
    'Convert Menus to Ownerdrawn
    MnuCtr = 0
    SetAllMenusOwnerDraw GetMenu(Frm.hWnd)
    'Cange the Menus Backcolor
    ChangeMenucolor Frm.hWnd, MnuBackColor
    ColBack = MnuBackColor
    ColFor = MnuForeColor
    ColSel = MnuSelectColor
    ColSelTxt = MnuSelectTextColor
    ColDisTxt = MnuDisabledTextColor
    'Start Subclassing
    Hook Frm.hWnd

End Sub

'Change Menus Backcolor even for not subclassed parts
Private Sub ChangeMenucolor(lngHWnd As Long, Background As Long)

Dim MI As MENUINFO

    With MI
        .cbSize = Len(MI)
        .fMask = MIM_BACKGROUND Or MIM_APPLYTOSUBMENUS
        .hbrBack = CreateSolidBrush(Background)
    End With 'MI
    SetMenuInfo GetMenu(lngHWnd), MI

End Sub

Private Sub DoMenuStuff(ds As DRAWITEMSTRUCT, IsMainMenu As Boolean)

Dim MenuText   As String * 255
Dim Length     As Long
Dim Index      As Long
Dim UseBrush   As Long
Dim CurFnt     As Long
Dim NewFnt     As Long
Dim Lf         As LOGFONT
Dim TopRect    As RECT
Dim BrushColor As Long
Dim TextColor  As Long

    Length = GetMenuString(ds.hwndItem, ds.itemID, MenuText, 255, MF_BYCOMMAND)
    Index = Val(Left$(MenuText, Length))
    'Do nothing if we have an PopUp Menu
    If MenuItems(Index).PopUp Then
        Exit Sub
    End If
    'Dont draw the same thing twice....
    If OldDS.itemID = ds.itemID Then
        If OldDS.itemState = ds.itemState Then
            Exit Sub
        End If
    End If
    OldDS = ds
    'Draw Seperator if needed
    With ds.rcItem
        If .Bottom - .Top = 7 Then
            ApiLine ds.hdc, .Left + 1, .Top + 3, .Right - 1, .Top + 3, ColFor
            Exit Sub
        End If
    End With 'DS.RCITEM
    BrushColor = ColBack
    'check to see if our menu item is selected but not disabled
    If ds.itemState And ODS_SELECTED Then
        If Not ds.itemState And ODS_DISABLED Then
            BrushColor = ColSel
        End If
    End If
    UseBrush = CreateSolidBrush(BrushColor)
    'fill our selected area with color
    FillRect ds.hdc, ds.rcItem, UseBrush
    'then delete the brush we created
    If UseBrush Then
        DeleteObject UseBrush
    End If
    'again check to see if the menu item is selected
    If ds.itemState And ODS_DISABLED Then
        'set our menu Disabled text color
        TextColor = ColDisTxt
    ElseIf ds.itemState And ODS_SELECTED Then 'NOT DS.ITEMSTATE...
        'set our menu selected text color
        TextColor = ColSelTxt
    Else 'NOT DS.ITEMSTATE...
        'set our text unselected text color
        TextColor = ColFor
    End If
    With ds
        SetTextColor .hdc, TextColor
        SetBkMode .hdc, TRANSPARENT
        'create two regions. userect will be for drop down
        'menu items and toprect will be used for top menus.
        LSet TopRect = .rcItem
        'retrieve the system font
        CurFnt = SelectObject(.hdc, GetStockObject(SYSTEM_FONT))
    End With 'ds
    GetObjectAPIBynum CurFnt, Len(Lf), VarPtr(Lf)
    'i set the font to ff_script, but you can also use;
    'ff_dontcare, ff_modern, ff_roman, ff_swiss.
    With Lf
        .lfPitchAndFamily = FF_SCRIPT
        .lfFaceName(1) = 1
        'i set the font weight to 400 for normal text. it would
        'be bolded at 600.
        .lfWeight = 400
        'create and select our new font.
    End With 'Lf
    NewFnt = CreateFontIndirect(Lf)
    SelectObject ds.hdc, NewFnt
    'set our top and left corners for our regions.
    TopRect.Left = TopRect.Left + 17
    TopRect.Top = TopRect.Top + 2
    'now we will draw our text according to the menu's id.
    'we must specify the text and the region we are drawing.
    DrawText ds.hdc, MenuItems(Index).MenuText, Len(MenuItems(Index).MenuText), TopRect, DT_LEFT Or DT_TOP Or DT_SINGLELINE
    'Draw Hotkey if Needed
    TopRect.Right = TopRect.Right - 10
    If LenB(MenuItems(Index).Hotkey) <> 0 Then
        DrawText ds.hdc, MenuItems(Index).Hotkey, Len(MenuItems(Index).Hotkey), TopRect, DT_RIGHT Or DT_TOP Or DT_SINGLELINE
    End If
    'select our old font and delete the one we created.
    SelectObject ds.hdc, CurFnt
    If NewFnt Then
        DeleteObject NewFnt
    End If
    With ds.rcItem
        'Draw Arrow if needed
        If IsMenuHandle(ds.itemID) And Not IsMainMenu Then
            DrawPolyArrow ds.hdc, .Right - 10, .Top + 4, .Right - 10, .Top + 12, .Right - 6, .Top + 8, TextColor
        End If
        If ds.itemState And ODS_CHECKED Then
            ApiLine ds.hdc, .Left + 5, .Top + 7, .Left + 8, .Top + 10, TextColor, 1
            ApiLine ds.hdc, .Left + 8, .Top + 10, .Left + 13, .Top + 5, TextColor, 1
            ApiLine ds.hdc, .Left + 5, .Top + 8, .Left + 8, .Top + 11, TextColor, 1
            ApiLine ds.hdc, .Left + 8, .Top + 11, .Left + 13, .Top + 6, TextColor, 1
            ApiLine ds.hdc, .Left + 5, .Top + 9, .Left + 8, .Top + 12, TextColor, 1
            ApiLine ds.hdc, .Left + 8, .Top + 12, .Left + 13, .Top + 7, TextColor, 1
        End If
        If ds.CtlType = ODT_MENU Then
            ExcludeClipRect ds.hdc, .Left, .Top, .Right, .Bottom
        End If
    End With 'ds.rcItem

End Sub

Public Sub DrawPolyArrow(ByVal lngHdc As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, X3 As Long, Y3 As Long, DrawColor As Long)

Dim pts(3)   As POINTAPI
Dim OldBrush As Long
Dim OldPen   As Long
Dim hPen     As Long
Dim hBrush   As Long

    hPen = CreatePen(0, 1, DrawColor)
    hBrush = CreateSolidBrush(DrawColor)
    OldPen = SelectObject(lngHdc, hPen)
    OldBrush = SelectObject(lngHdc, hBrush)
    pts(0).x = X1
    pts(0).y = Y1
    pts(1).x = X2
    pts(1).y = Y2
    pts(2).x = X3
    pts(2).y = Y3
    Polygon lngHdc, pts(0), 3
    SelectObject lngHdc, OldPen
    SelectObject lngHdc, OldBrush
    DeleteObject hPen
    DeleteObject hBrush

End Sub

'Get the Menus name
'So we can create a call to it´s Mehtod
Private Sub GetMenuData(Frm As Form)

Dim nControl As Control

    For Each nControl In Frm.Controls
        If TypeOf nControl Is Menu Then
            Newdim
            MenuItems(UBound(MenuItems)).MenuText = nControl.Caption
            'Now set a new caption that represents the index of our MenuItems Array
            nControl.Caption = UBound(MenuItems)
            'Is this an hidden PopUp menu
            'If yes then we have to name it mnupop.....
            If Left$(LCase$(nControl.Name), 6) = "mnupop" Then
                nControl.Enabled = False
                MenuItems(UBound(MenuItems)).PopUp = True
            End If
        End If
    Next nControl

End Sub

Private Function GetTextWidth(ByVal TestText As String) As Long

Dim CurFnt  As Long
Dim NewFnt  As Long
Dim Lf      As LOGFONT
Dim UseRect As RECT

    CurFnt = SelectObject(SubContainer.hdc, GetStockObject(SYSTEM_FONT))
    GetObjectAPIBynum CurFnt, Len(Lf), VarPtr(Lf)
    With Lf
        .lfPitchAndFamily = FF_SCRIPT
        .lfFaceName(1) = 1
        .lfWeight = 400
    End With 'Lf
    NewFnt = CreateFontIndirect(Lf)
    SelectObject SubContainer.hdc, NewFnt
    DrawText SubContainer.hdc, TestText, Len(TestText), UseRect, DT_CALCRECT Or DT_LEFT Or DT_TOP Or DT_SINGLELINE
    GetTextWidth = UseRect.Right - UseRect.Left
    SelectObject SubContainer.hdc, CurFnt
    If NewFnt Then
        DeleteObject NewFnt
    End If

End Function

Private Sub Hook(lngHWnd As Long)

    m_hWnd = lngHWnd
    PrevProc = SetWindowLong(lngHWnd, GWL_WNDPROC, AddressOf WindowProc)

End Sub

'Check if in IDE
'Needed for errortesting and Message in subclassing
Private Function InIDE() As Boolean

    On Error GoTo DivideError
    Debug.Print 1 / 0
    InIDE = False

Exit Function

DivideError:
    InIDE = True

End Function

Private Function IsMenuHandle(ByVal hMenu As Long) As Boolean

    IsMenuHandle = Not (GetMenuItemCount(hMenu) = -1)

End Function

Private Function MenuHasSubMenues(ByVal hMenu As Long) As Boolean

    MenuHasSubMenues = GetMenuItemCount(hMenu) <> -1

End Function

'Redim preserve without errors
Private Sub Newdim()

    On Error GoTo ErrOut
    ReDim Preserve MenuItems(UBound(MenuItems) + 1)

Exit Sub

ErrOut:
    ReDim MenuItems(0)

End Sub

Private Function SetAllMenusOwnerDraw(ByVal hMenu As Long)

Dim mii      As MENUITEMINFO
Dim i        As Integer
Dim Length   As Long
Dim MenuText As String * 255
Dim x        As Long

    mii.cbSize = Len(mii)
    mii.fMask = MIIM_FTYPE
    For i = 0 To GetMenuItemCount(hMenu)
        'By the way check for hotkey´s
        Length = GetMenuString(hMenu, i, MenuText, 255, MF_BYPOSITION)
        x = InStr(1, MenuText, vbTab)
        If x > 0 Then
            MenuItems(MnuCtr).Hotkey = Mid$(MenuText, x + 1, Length - x)
        End If
        'Now set the menu Ownerdraw
        MnuCtr = MnuCtr + 1
        GetMenuItemInfo hMenu, i, True, mii
        mii.fType = mii.fType Or MF_OWNERDRAW
        SetMenuItemInfo hMenu, i, True, mii
        If MenuHasSubMenues(GetSubMenu(hMenu, i)) Then
            SetAllMenusOwnerDraw GetSubMenu(hMenu, i)
        End If
    Next i

End Function

Public Sub UnHook()

    SetWindowLong m_hWnd, GWL_WNDPROC, PrevProc

End Sub

Private Function WindowProc(ByVal hW As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim di       As DRAWITEMSTRUCT
Dim ms       As MEASUREITEMSTRUCT
Dim Length   As Long
Dim MenuText As String * 255
Dim Index    As Long

    Select Case uMsg
    Case WM_DRAWITEM
        'with our menus now being owner drawn, we can now
        'process the wm_drawitem message. wm_drawitem is
        'sent to our hwnd when the menu is being drawn.
        'by intercepting this message, we can change
        'various attributes of the menu including; the
        'text color and selection colors.
        CopyMemory di, ByVal lParam, Len(di)
        DoMenuStuff di, (GetMenu(hW) = di.hwndItem)
    Case WM_MEASUREITEM
        'with our menus now being owner drawn, we can now
        'process the wm_measureitem message. this message
        'is also sent to our hwnd when the menu is being
        'drawn. we must intercept this message with owner
        'drawn menus in order to specify their size.
        CopyMemory ms, ByVal lParam, Len(ms)
        'make sure our structure is owner drawn.
        If ms.CtlType = ODT_MENU Then
            'set initial values for the structure.
            Length = GetMenuString(GetMenu(hW), ms.itemID, MenuText, 255, MF_BYCOMMAND)
            Index = Val(Left$(MenuText, Length))
            If MenuItems(Index).MenuText = "-" Then 'Seperatorline
                ms.itemHeight = 7
                ms.itemWidth = 1
            Else 'NOT MENUITEMS(INDEX).MENUTEXT...
                ms.itemWidth = GetTextWidth(MenuItems(Index).MenuText) + 18
                If LenB(MenuItems(Index).Hotkey) > 0 Then
                    ms.itemWidth = ms.itemWidth + GetTextWidth(MenuItems(Index).Hotkey) + 15
                End If
                ms.itemHeight = 17
            End If
        End If
        CopyMemory ByVal lParam, ms, Len(ms)
    Case WM_DESTROY, WM_CLOSE
        UnHook
    End Select
    WindowProc = CallWindowProc(PrevProc, hW, uMsg, wParam, lParam)

End Function
