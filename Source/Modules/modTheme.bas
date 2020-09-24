Attribute VB_Name = "modTheme"
'//---------------------------------------------------------------------------------------
'//--Module    : modTheme - (mGlobal)
'//--DateTime  : 01/02/2005
'//--Author    : Gary Noble   ©2005 Telecom Direct Limited
'//--Purpose   : Global Subs And Functions Relating To Drawing And Windows Theme State
'//--Assumes   :
'//--Notes     : Updated removed unused API and Variables (Silvellis N. Cawaling)
'//--Revision  : 2.0
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Gary Noble  01/02/2005
'//---------------------------------------------------------------------------------------
Option Explicit
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type

Private Const IDC_HAND As Long = 32649
Private Const IDC_ARROW As Long = 32512
'//-- Current Theme Name
Public m_sCurrentSystemThemename As String
'Public colTrackMouse As New Collection
'Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, _
                                                                                  ByVal lpfn As Long, _
                                                                                  ByVal hMod As Long, _
                                                                                  ByVal dwThreadId As Long) As Long
'Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
                                                      ByVal ncode As Long, _
                                                      ByVal wParam As Long, _
                                                      lParam As Any) As Long
'Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
'Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, _
                                                                      ByVal lpCursorName As Long) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                                               ByVal ptX As Long, _
                                               ByVal ptY As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, _
                                                     lpPoint As POINTAPI) As Long
''Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, _
                                                               ByVal hPalette As Long, _
                                                               pccolorref As Long) As Long
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, _
                                                          ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As Long, _
                                                                ByVal dwMaxNameChars As Long, _
                                                                ByVal pszColorBuff As Long, _
                                                                ByVal cchMaxColorChars As Long, _
                                                                ByVal pszSizeBuff As Long, _
                                                                ByVal cchMaxSizeChars As Long) As Long
Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long


Public glColorOneNormal As Long
Public glColorTwoNormal As Long
'Public glColorTopNormal As Long
'Public glColorButtomNormal As Long
Public glColorBorder As Long
Public glColorHeaderColorOne As Long
Public glColorHeaderColorTwo As Long
Public glColorOneSelected As Long
Public glColorTwoSelected As Long

'//---------------------------------------------------------------------------------------
'//--Procedure : AppThemed
'//--Type      : Function
'//--DateTime  : 01/02/2005
'//--Author    : Gary Noble
'//--Purpose   : Determines If The Current Window is Themed
'//--Returns   : Boolean
'//--Notes     :
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Gary Noble  01/02/2005
'//---------------------------------------------------------------------------------------
Public Function AppThemed() As Boolean

    On Error Resume Next
    AppThemed = IsAppThemed()
    On Error GoTo 0

End Function

'//---------------------------------------------------------------------------------------
'//--Procedure : BlendColor
'//--Type      : Property
'//--DateTime  : 03/02/2005
'//--Author    : Gary Noble
'//--Purpose   : Blends Two Colours Together
'//--Returns   : Long
'//--Notes     :
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Gary Noble  03/02/2005
'//---------------------------------------------------------------------------------------
Public Property Get BlendColor(ByVal oColorFrom As OLE_COLOR, _
                               ByVal oColorTo As OLE_COLOR, _
                               Optional ByVal Alpha As Long = 128) As Long

    Dim lSrcR As Long

    Dim lSrcG As Long
    Dim lSrcB As Long
    Dim lDstR As Long
    Dim lDstG As Long
    Dim lDstB As Long
    Dim lCFrom As Long
    Dim lCTo As Long
    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    BlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))

End Property

Public Function GetKeyName(ByVal KeyCode As KeyCodeConstants, _
                           Optional ByVal Alt As Boolean, _
                           Optional ByVal Ctrl As Boolean, _
                           Optional ByVal Shift As Boolean) As String


    Dim strLeft As String

    If ((KeyCode >= vbKeyF1) And (KeyCode <= vbKeyF16)) Then
        GetKeyName = "F" & (KeyCode - vbKeyF1) + 1
    ElseIf ((KeyCode >= vbKeyA) And (KeyCode <= vbKeyZ)) Then
        GetKeyName = Chr$((KeyCode - vbKeyA) + 65)
    ElseIf ((KeyCode >= vbKey0) And (KeyCode <= vbKey9)) Then
        GetKeyName = (KeyCode - vbKey0)
    ElseIf ((KeyCode >= vbKeyNumpad0) And (KeyCode <= vbKeyNumpad9)) Then
        GetKeyName = "Numpad" & (KeyCode - vbKeyNumpad0)
    ElseIf (KeyCode = vbKeyDelete) Then
        GetKeyName = "Delete"
    ElseIf (KeyCode = vbKeyTab) Then
        GetKeyName = "Tab"
    ElseIf (KeyCode = vbKeyEscape) Then
        GetKeyName = "Escape"
    End If
    If GetKeyName <> vbNullString Then
        If Alt Then
            strLeft = "Alt"
        End If
        If Ctrl Then
            If strLeft = vbNullString Then
                strLeft = "Ctrl"
            Else
                strLeft = strLeft & "+" & "Ctrl"
            End If
        End If
        If Shift Then
            If strLeft = vbNullString Then
                strLeft = "Shift"
            Else
                strLeft = strLeft & "+" & "Shift"
            End If
        End If
        If strLeft <> vbNullString Then
            GetKeyName = strLeft & "+" & GetKeyName
        End If
    End If

End Function

'//---------------------------------------------------------------------------------------
'//--Procedure : GetThemeName
'//--Type      : Sub
'//--DateTime  : 01/02/2005
'//--Author    : Gary Noble
'//--Purpose   : Returns The current Windows Theme Name
'//--Returns   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Gary Noble  01/02/2005
'//---------------------------------------------------------------------------------------
Public Sub GetThemeName(lngHwnd As Long)


    Dim hTheme As Long
    Dim sShellStyle As String
    Dim sThemeFile As String
    Dim lPtrThemeFile As Long
    Dim lPtrColorName As Long

    'Dim hres As Long
    Dim iPos As Long
    On Error Resume Next
    hTheme = OpenThemeData(lngHwnd, StrPtr("ExplorerBar"))
    If Not hTheme = 0 Then
        ReDim bThemeFile(0 To 260 * 2) As Byte
        lPtrThemeFile = VarPtr(bThemeFile(0))
        ReDim bColorName(0 To 260 * 2) As Byte
        lPtrColorName = VarPtr(bColorName(0))
        GetCurrentThemeName lPtrThemeFile, 260, lPtrColorName, 260, 0, 0
        sThemeFile = bThemeFile
        iPos = InStr(sThemeFile, vbNullChar)
        If iPos > 1 Then
            sThemeFile = Left$(sThemeFile, iPos - 1)
        End If
        m_sCurrentSystemThemename = bColorName
        iPos = InStr(m_sCurrentSystemThemename, vbNullChar)
        If iPos > 1 Then
            m_sCurrentSystemThemename = Left$(m_sCurrentSystemThemename, iPos - 1)
        End If
        sShellStyle = sThemeFile
        For iPos = Len(sThemeFile) To 1 Step -1
            If (Mid$(sThemeFile, iPos, 1) = "\") Then
                sShellStyle = Left$(sThemeFile, iPos)
                Exit For
            End If
        Next iPos
        sShellStyle = sShellStyle & "Shell\" & m_sCurrentSystemThemename & "\ShellStyle.dll"
        CloseThemeData hTheme
    Else
        m_sCurrentSystemThemename = "Classic"
    End If
    On Error GoTo 0

End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : SetScreenCursor
'//--Type      : Sub
'//--DateTime  : 01/02/2005
'//--Author    : Gary Noble
'//--Purpose   : Sets The Screen To Use The Windows Hand Or Normal Cursor
'//--Returns   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Gary Noble  01/02/2005
'//---------------------------------------------------------------------------------------
Public Sub SetScreenCursor(ByVal bHand As Boolean)

    If bHand Then
        SetCursor LoadCursor(0, IDC_HAND)
    Else
        SetCursor LoadCursor(0, IDC_ARROW)
    End If

End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : TranslateColor
'//--Type      : Function
'//--DateTime  : 03/02/2005
'//--Author    : Gary Noble
'//--Purpose   : Convert Automation color to Windows color
'//--Returns   : Long
'//--Notes     :
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Gary Noble  03/02/2005
'//---------------------------------------------------------------------------------------
Public Function TranslateColor(ByVal oClr As OLE_COLOR, _
                               Optional hPal As Long = 0) As Long

    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If

End Function

Public Sub GetGradientColor(lhWnd As Long)

    GetThemeName lhWnd
    
    If AppThemed Then   '/Check if themed.
        Select Case m_sCurrentSystemThemename
            Case "HomeStead"
                glColorOneNormal = RGB(228, 235, 200)
                glColorTwoNormal = RGB(175, 194, 142)
                glColorBorder = RGB(100, 144, 88)
                glColorHeaderColorOne = RGB(228, 235, 200)
                glColorHeaderColorTwo = RGB(175, 194, 142)
            Case "NormalColor"
                glColorOneNormal = RGB(197, 221, 250)
                glColorTwoNormal = RGB(128, 167, 225)
                glColorBorder = RGB(0, 45, 150)
                glColorHeaderColorOne = RGB(197, 221, 250)
                glColorHeaderColorTwo = RGB(128, 167, 225)
            Case "Metallic"
                glColorOneNormal = RGB(219, 220, 232)
                glColorTwoNormal = RGB(149, 147, 177)
                glColorBorder = RGB(119, 118, 151)
                glColorHeaderColorOne = RGB(219, 220, 232)
                glColorHeaderColorTwo = RGB(149, 147, 177)
            Case Else
                glColorOneNormal = BlendColor(vbButtonFace, vbWhite, 120)
                glColorTwoNormal = vbButtonFace
                glColorHeaderColorOne = BlendColor(vbButtonFace, vbWhite, 120)
                glColorHeaderColorTwo = vbButtonFace
                glColorBorder = TranslateColor(vbInactiveTitleBar)
        End Select
        'm_lngColorHotOne = BlendColor(vbWindowBackground, vbButtonFace, 220)
        'm_lngColorHotTwo = RGB(248, 216, 126)
        glColorOneSelected = RGB(240, 160, 38)
        glColorTwoSelected = RGB(248, 216, 126)
    Else
        'm_lngColorHotTwo = BlendColor(vbInactiveTitleBar, BlendColor(vbButtonFace, vbWhite, 50), 10)
        'm_lngColorHotOne = m_lngColorHotTwo
        glColorOneNormal = BlendColor(vbButtonFace, vbWhite, 120)
        glColorTwoNormal = vbButtonFace
        glColorBorder = BlendColor(vbButtonFace, vbBlack, 200)
        glColorHeaderColorOne = BlendColor(vbButtonFace, vbWhite, 120)
        glColorHeaderColorTwo = vbButtonFace
        glColorOneSelected = BlendColor(vbInactiveTitleBar, BlendColor(vbButtonFace, vbWhite, 150), 100)
        glColorTwoSelected = glColorOneSelected
    End If

End Sub




