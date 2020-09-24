VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "Sys Info"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   3960
      Width           =   1065
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   3540
      Width           =   1065
   End
   Begin VB.CommandButton cmdEmail 
      Caption         =   "Email"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   4380
      Width           =   1065
   End
   Begin VB.ListBox lstCredits 
      Height          =   1815
      Left            =   150
      TabIndex        =   0
      Top             =   1500
      Width           =   5715
   End
   Begin VB.Label lblCopyright 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "%Copyright%"
      Height          =   255
      Left            =   150
      TabIndex        =   5
      Top             =   1230
      Width           =   4695
   End
   Begin VB.Image imgMe 
      Height          =   795
      Left            =   510
      Top             =   180
      Width           =   4545
   End
   Begin VB.Label lblComments 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "%Comments%"
      Height          =   1185
      Left            =   150
      TabIndex        =   4
      Top             =   3480
      Width           =   4575
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    'Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    'Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                               ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub AddCredits()
    With lstCredits
      .AddItem "Developed by:"
      .AddItem vbTab & "Silvellis N. Cawaling (e-Shock Technology)"
      .AddItem vbTab & "Transco NLOM ISTSS Philippines"
      .AddItem vbTab & "Email: silvellis@yahoo.com"
      .AddItem "________________________________________________________________"
      .AddItem "Credits to:"
      .AddItem vbTab & "Keith Fox (Lavolpe)"
      .AddItem vbTab & "- Lavolpe buttons control and some of his code"
      .AddItem vbTab & "  used in the system."
      .AddItem vbTab & "Gary Noble (Phantom Man)"
      .AddItem vbTab & "- Phantom Panel control and the theme module"
      .AddItem vbTab & "  used in the system."
      .AddItem vbTab & "Carles PV"
      .AddItem vbTab & "- For the ucStatusbar control."
      .AddItem vbTab & "  used in the system."
      .AddItem vbTab & "Neeraj Agrawal"
      .AddItem vbTab & "- For the XTab control."
      .AddItem vbTab & "  used in the system."
      .AddItem vbTab & "Matthew R. Usner"
      .AddItem vbTab & "- For the ucGradContainer control."
      .AddItem vbTab & "  used in the system."
      .AddItem vbTab & "RegX"
      .AddItem vbTab & "- For the modINI module."
      .AddItem vbTab & "  used in the system."
      .AddItem vbTab & "Fredrik Qvarfort"
      .AddItem vbTab & "- For the Huffman Encoding/Decoding Class"
      .AddItem vbTab & "  used in the system."
      .AddItem "________________________________________________________________"
      .AddItem "Special Thanks to:"
      .AddItem vbTab & "Planet Source Code"
      .AddItem vbTab & "Free VB Code"
      .AddItem vbTab & "VB Accelerator"
      .AddItem vbTab & "KPD-Team (API-Guide)"
      .AddItem vbTab & "MZ Tools"
      .AddItem vbTab & "Dean Camera (DeepLook)"
    End With
End Sub

Private Sub cmdEmail_Click()
On Error Resume Next
    If OpenThisFile("mailto:silvellis@yahoo.com", 1, "", hWnd) = -1 Then
        InputBox "Sending email failed. Here's the email address if you want to contact me.", "Email Address", "silvellis@yahoo.com"
    End If
End Sub

Private Sub cmdOk_Click()
    Call Unload(Me)
End Sub

Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub

Private Sub Form_Load()

On Error GoTo errHandler

    Call CenterForm(Me)
    
    Call AddCredits
    
    Caption = "About " & g_sAppName

    lblCopyright = App.LegalCopyright
    lblComments = "Warning: " & App.Comments
    
    Icon = frmMain.Icon
    
    imgMe = LoadResPicture(104, 0)
    'imgAbout = LoadResPicture(105, 0)

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmAbout", "Form_Load", True)
        Call Unload(Me)
    End If
End Sub


