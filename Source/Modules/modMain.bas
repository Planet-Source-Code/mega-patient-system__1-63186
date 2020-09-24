Attribute VB_Name = "modMain"
'//---------------------------------------------------------------------------------------
'//--Module    : modMain
'//--DateTime  : 7/12/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   : Global Variables and Subroutines.
'//--Notes     :
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation Silvellis N. Cawaling 7/12/2005
'//---------------------------------------------------------------------------------------
'
Option Explicit

'-- Global Enumerations.
Public Enum g_eMenu
    [Patient] = 1
    [Fee] = 2
    [User] = 3
    [Visitation] = 4
    [Today] = 5
    [Report] = 6
End Enum
Public Enum g_eIcon
    [Add] = 1
    [Edit] = 2
    [Delete] = 3
    [Refresh] = 4
    [Search] = 5
    [Print] = 6
    [Options] = 7
    [Save] = 8
    [Cancel] = 9
    [Help] = 10
    [Ok] = 11
    [Patient] = 1
    [Fee] = 2
    [User] = 3
    [Visitation] = 4
    [Today] = 5
    [Report] = 6
    [BOF] = 12
    [Previous] = 13
    [Next] = 14
    [EOF] = 15
End Enum
Public Enum g_eLevel
    [Personnel] = 0
    [Administrator] = 1
End Enum

'-- Local Type Declarations.
Private Type tUser
    ID          As Long
    UserID      As String
    Name        As String
    Password    As String
    Level       As Long
End Type
Private Type tSetting
    Splash      As Long
    LastLog     As Long
End Type

'-- Global Constants.
Public Const SX_MALE        As String = "Male"
Public Const SX_FEMALE      As String = "Female"

Public Const CS_SINGLE      As String = "Single"
Public Const CS_MARRIED     As String = "Married"
Public Const CS_WIDOW       As String = "Widow"
Public Const CS_SEP         As String = "Separated"

Public Const CPYRYT         As String = "Silvellis N. Cawaling © 2005"
Public Const FMT_DT         As String = "MM/DD/YYYY"
Public Const FMT_CUR        As String = "#0.00"
'Public Const DFT_DT         As String = "01/01/1900"
Public Const CPT_ADD        As String = "-[ADD]"
Public Const CPT_EDIT       As String = "-[EDIT]"

Public Const DB_NAME        As String = "PatientSys.mdb"
Public Const INI_NAME       As String = "PatientSys.ini"


'-- Global Variables.
Public g_sAppName           As String                   '-- Application Name.
Public g_cn                 As ADODB.Connection         '-- Connection.
Public g_rsPatient          As ADODB.Recordset          '-- Patient Recordset.
Public g_rsFee              As ADODB.Recordset          '-- Fee Recordset.
Public g_rsUser             As ADODB.Recordset          '-- User Recordset.
Public g_rsVisitation       As ADODB.Recordset          '-- Visition Recordset.
Public g_rsToday            As ADODB.Recordset          '-- Today.
Public g_sConnectionString  As String                   '-- Connection String.
Public g_sSQL               As String
Public g_tUser              As tUser                    '-- User Info.
Public g_tSetting           As tSetting                 '-- System Settings.
Public g_sSysDate           As String                   '-- System Date.
Public g_lLogOff            As Long                     '-- LogOff flag.
    
'-- Global Declarations.
Public Declare Sub InitCommonControls Lib "Comctl32.dll" ()    '-- XP UI.
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
        ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" ( _
        ByVal hLibModule As Long) As Long
        
'-- Local Declarations.
Private Declare Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'

Private Sub Main()

On Error GoTo errHandler

    '-- Init global variables.
    g_sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & DB_NAME & ";Persist Security Info=False;"
    g_sAppName = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision
    g_tSetting.Splash = GetINI(INI_NAME, "Option", "LastLog", 1)
    g_tSetting.LastLog = GetINI(INI_NAME, "Option", "LastLog", 1)
    
    If g_tSetting.Splash = 1 Then frmSplash.Show vbModal
    frmLogin.Show

errHandler:
    If Err.Number = 364 Then '-- Object Unloaded.
        Err.Clear
    ElseIf Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "modMain", "Main", True)
    End If
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : CenterForm
'//--DateTime  : 5/12/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   : Center the a form.
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Public Sub CenterForm(ByRef Frm As Form)
    Frm.Left = (Screen.Width - Frm.Width) / 2
    Frm.Top = (Screen.Height - Frm.Height) / 2
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : initFrame
'//--DateTime  : 7/6/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Public Sub initFrame(fraBackground As ucGradContainer, _
                    Optional imlIcons As ImageList, _
                    Optional lImageIdx As Long)

On Error GoTo errHandler
    
    With fraBackground
        .BorderColor = glColorBorder
        .BackColor1 = vbButtonFace
        .BackColor2 = vbButtonFace
        .HeaderColor1 = glColorHeaderColorTwo
        .HeaderColor2 = glColorHeaderColorOne
        .CaptionAlignment = Center
        If lImageIdx <> 0 Then
            Set .Icon = imlIcons.ListImages(lImageIdx).Picture
        End If
    End With

errHandler:
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : PromptError
'//--DateTime  : 4/17/2005
'//--Author    : Silvellis N. Cawaling (Ellis)
'//--Purpose   : Show error message and log error.
'//--Returns   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
Public Sub PromptError(Optional ByVal ErrNumber As String = vbNullString, _
                        Optional ByVal ErrDescription As String = vbNullString, _
                        Optional ByVal ErrSource As String = vbNullString, _
                        Optional ByVal ModuleName As String = vbNullString, _
                        Optional ByVal ProcName As String = vbNullString, _
                        Optional ByVal DisplayPrompt As Boolean)

    Dim tString As String
    Dim fnum As Integer

    '-- Show Error Message.
    If Err.Number <> 0 Then
        If DisplayPrompt Then
            tString = "Error occured: "
            If Trim$(ErrNumber) <> "" Then tString = tString & ErrNumber & vbNewLine Else tString = tString & vbNewLine
            If Trim$(ErrDescription) <> "" Then tString = tString & "Description: " & ErrDescription & vbNewLine
            If Trim$(ErrSource) <> "" Then tString = tString & "Source: " & ErrSource & vbNewLine
            If Trim$(ModuleName) <> "" Then tString = tString & "Module: " & ModuleName & vbNewLine
            If Trim$(ProcName) <> "" Then tString = tString & "Function: " & ProcName
            MsgBox tString, vbCritical, App.Title
        End If
        '-- Write error log.
        fnum = FreeFile
        Open App.Path & "\ErrorLog.txt" For Append As #fnum
        Write #fnum, Date, ErrNumber, ErrDescription, ErrSource, ModuleName, ProcName, Environ("username"), Environ("computername")
        Close #fnum
    End If

End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : FormatButton
'//--DateTime  : 4/17/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   : Set/format Lavolpe button properties. (Themed Button)
'//--Returns   :
'//--Notes     : Toolbar
'//---------------------------------------------------------------------------------------

Public Sub initToolbar(ByRef cmdLavolpe As lvButtons_H, _
                        imlIcons As ImageList, _
                        Optional lFooter As Long, _
                        Optional lPictureIdx As Long, _
                        Optional lPictureSize As Long = lv_16x16)

    With cmdLavolpe
        If lFooter = 0 Then '-- Header.
            .BackColor = glColorTwoNormal
            .GradientColor = glColorOneNormal
            .HoverBackColor = glColorTwoSelected
            .HoverBackColorEnd = glColorOneSelected
            .PictureSize = lPictureSize
        Else    '-- Footer.
            .BackColor = glColorOneNormal
            .GradientColor = glColorTwoNormal
            .HoverBackColor = glColorOneSelected
            .HoverBackColorEnd = glColorTwoSelected
            .PictureSize = lPictureSize
        End If

        '-- Button picture.
        If lPictureIdx > 0 Then
            Set .Picture = imlIcons.ListImages(lPictureIdx).Picture
        End If

        .GradientMode = lv_Bottom2Top
        .ButtonStyle = lv_hover
        .PictureAlign = lv_LeftOfCaption
    End With

End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : initButton
'//--DateTime  : 4/29/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   : Set/format Lavolpe button properties. (Hover button)
'//--Returns   :
'//--Notes     : Lookup
'//---------------------------------------------------------------------------------------
'
Public Sub initButton(ByRef cmdLavolpe As lvButtons_H, _
                        Optional imlIcons As ImageList, _
                        Optional lPictureIdx As Long, _
                        Optional lPictureSize As Long = lv_16x16)
                        
    Set cmdLavolpe.Picture = imlIcons.ListImages(lPictureIdx).Picture
    
    cmdLavolpe.ButtonStyle = lv_hover
    cmdLavolpe.PictureSize = lPictureSize
    cmdLavolpe.ShowFocusRect = False
    'Set cmdLavolpe.MouseIcon = LoadResPicture(101, 2)
    
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : StartBusy
'//--DateTime  : 4/17/2005
'//--Author    : Silvellis N. Cawaling (Ellis)
'//--Purpose   : Start system busy state.
'//--Returns   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Silvellis N. Cawaling (Ellis)  4/17/2005
'//---------------------------------------------------------------------------------------
Public Sub StartBusy()
    Screen.MousePointer = vbHourglass                   '-- Change screen mouse cursor.
    'frmMain.ucWait.PlayWait                             '-- Play wait control.
    frmMain.ucStatus.PanelText(1) = "Processing..."       '-- Change main form caption.
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : EndBusy
'//--DateTime  : 4/17/2005
'//--Author    : Silvellis N. Cawaling (Ellis)
'//--Purpose   : End system busy state.
'//--Returns   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'//--History   : Initial Implementation    Silvellis N. Cawaling (Ellis)  4/17/2005
'//---------------------------------------------------------------------------------------
Public Sub EndBusy()
    Screen.MousePointer = vbDefault                     '-- Change screen mouse cursor.
    'frmMain.ucWait.StopWait                             '-- Play wait control.
    frmMain.ucStatus.PanelText(1) = CPYRYT               '-- Change main form caption.
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : ClearAll
'//--DateTime  : 7/22/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   : Clear all form objects
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Public Sub ClearAll(Frm As Form)
    Dim ctl As Control
    For Each ctl In Frm
        If TypeOf ctl Is TextBox Then
            ctl.Text = vbNullString
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.ListIndex = -1
        End If
    Next
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : Numeric
'//--DateTime  : 11/01/2004
'//--Author    : Silvellis N. Cawaling
'//--Purpose   : Numeric inputs.
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Public Sub Numeric(txt As TextBox, KeyAscii As Integer)
    Select Case KeyAscii
        Case "48" To "57", vbKeyBack                        '-- Numbers 0-9, Backspace
        Case "46"                                           '-- Period(.) ,45 -> Negative Sign
            If Len(txt) = txt.SelLength Then Exit Sub       '-- Check if whole text is highlighted.
            If InStr(1, txt, ".") > 0 Then KeyAscii = 0     '-- Lock Decimal.
        Case "13"                                           '-- Enter Key.
            SendKeys "{tab}"
        Case Else
            KeyAscii = 0
    End Select
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : Navigate
'//--DateTime  : 5/23/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   : Navigate Record.
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Public Sub Navigate(ByVal Index As Integer, _
                    ByRef rs As ADODB.Recordset)

On Error GoTo errHandler

    If Not (rs.BOF And rs.EOF) Then
        Select Case Index
        Case 0 '-- First
            rs.MoveFirst
        Case 1 '-- Previous
            rs.MovePrevious
            If rs.BOF Then
                rs.MoveFirst
            End If
        Case 2 '-- Next
            If rs.EOF Then
                rs.MoveLast
            Else
                rs.MoveNext
                If rs.EOF Then
                    rs.MoveLast
                End If
            End If
        Case 3 '-- Last
            rs.MoveLast
        End Select
    End If

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "modGlobal", "Navigate", True)
    End If
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : isValidData
'//--DateTime  : 11/01/2004
'//--Author    : Silvellis N. Cawaling
'//--Purpose   : Check for null and valid data type. Used to validate textbox and combobox
'//--Notes     : 0 = None.
'//              1 = Numeric.
'//              2 = Date.
'//---------------------------------------------------------------------------------------
'
Public Function isValidData(ctl As Control, sDesc As String, Optional lType As Long) As Boolean
    
On Error GoTo errHandler

    If TypeOf ctl Is ComboBox Then '-- Combobox.
        '-- Check for empty values
        If ctl.ListIndex = -1 Then
            MsgBox "Kindly select a valid " & LCase(sDesc) & ".", vbExclamation + vbOKOnly
            ctl.SetFocus
            Exit Function
        End If
    Else '-- Textbox.
        '-- Check for empty values.
        If Trim$(ctl) = vbNullString Then
            MsgBox "Kindly enter " & LCase(sDesc) & ".", vbExclamation + vbOKOnly
            ctl.SetFocus
            Exit Function
        End If
    End If
    
    If lType = 1 Then '-- Number.
        If Not IsNumeric(ctl) Then
            MsgBox sDesc & " is not a valid number.", vbExclamation + vbOKOnly
            ctl.SetFocus
            Exit Function
        End If
    ElseIf lType = 2 Then '-- Date.
        If Not IsDate(ctl) Then
            MsgBox sDesc & " is not a valid date.", vbExclamation + vbOKOnly
            ctl.SetFocus
            Exit Function
        End If
    End If
    
    isValidData = True

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "modMain", "isValidData", True)
    End If

End Function

'//---------------------------------------------------------------------------------------
'//--Procedure : OpenThisFile
'//--DateTime  : 5/12/2005
'//--Author    : Lavolpe
'//--Purpose   : Simply attempts to open a file with the Shell command, if errors are encountered then...
'                tries to open it with an API call using extensions to find associated executables, if error then...
'                prompts user with the "Open With ..." routine
'//--Notes     : Updated (Silvellis N. Cawaling)
'//---------------------------------------------------------------------------------------
'
Public Function OpenThisFile(ByVal stFile As String, _
                        ByVal lShowHow As Long, _
                        ByVal sParams As String, _
                        ByRef lhWnd As Long) As Variant

    Dim lRet As Long, stRet As String, ErrID As Long

On Error GoTo TryAPIcall
    lRet = -1   ' set default value -- meaning failure
    If Len(sParams) > 0 Then sParams = " " & sParams    ' if no optional parameters, then format with a space
    lRet = Shell(stFile & sParams, lShowHow)       ' attempt simple shell command
    OpenThisFile = lRet
    Exit Function

TryAPIcall:
    Err.Clear
    ' if above shell function failed, then try an association open based on the file extension
    ErrID = apiShellExecute(lhWnd, "OPEN", _
            stFile, sParams, App.Path, lShowHow)
    ' Errors will be a retruned value of <32
    If ErrID < 32& Then
        Select Case ErrID
            Case 31&:
                'Try the OpenWith dialog
                lRet = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " _
                        & stFile, 1)
            Case 0&:
                stRet = "Out of Memory/Resources. Could not Execute."
            Case 2&:
                stRet = "File not found. Could not Execute."
            Case 3&:
                stRet = "Path not found. Could not Execute."
            Case 11&:
                stRet = "Bad File Format. Could not Execute."
            Case Else:
        End Select
        If ErrID <> 31 Then
            lRet = -1 ' failure
            MsgBox stRet, vbExclamation + vbOKOnly  ' display error
        End If
        OpenThisFile = lRet
    Else
        lRet = 69
    End If
Resume Next
End Function

'//---------------------------------------------------------------------------------------
'//--Procedure : isExisting
'//--DateTime  : 11/07/2004
'//--Author    : Silvellis N. Cawaling
'//--Purpose   : Check if a cetain record exists.
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Public Function isExisting(cn As ADODB.Connection, sTableName As String, sCondition As String) As Boolean
    Dim rs As ADODB.Recordset
    
On Error GoTo errHandler

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM " & sTableName & " WHERE " & sCondition, cn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then isExisting = True
    
errHandler:
    Set rs = Nothing
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "modDB", "isExisting", True)
    End If
End Function

'//---------------------------------------------------------------------------------------
'//--Procedure : PointToRecord
'//--DateTime  : 11/07/2004
'//--Author    : Silvellis N. Cawaling
'//--Purpose   : Procedure used to point to the record
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Public Sub PointToRecord(ByRef rs As ADODB.Recordset, ByVal sField As String, ByVal isString As Boolean, ByVal sStr As String, ByVal sNum As Long)
    Dim lOldPosition As Long
    Dim sqlParam As String
    
On Error GoTo errHandler

    With rs
        lOldPosition = .AbsolutePosition
        rs.Filter = adFilterNone
        rs.Requery
        .MoveFirst
        '/Check if string or number.
        If isString Then
            sqlParam = sField & " = '" & sStr & "'"
        Else
            sqlParam = sField & " = " & sNum
        End If
        .Find sqlParam
        If .EOF Then .AbsolutePosition = lOldPosition
    End With

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "modDB", "PointToRecord", True)
    End If
End Sub

