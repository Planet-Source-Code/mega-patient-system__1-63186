VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin prjPatientSystem.ucHorizontal3DLine ucHorizontal3DLine1 
      Height          =   30
      Left            =   -360
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   53
   End
   Begin prjPatientSystem.ucGradContainer fraBackground 
      Height          =   1995
      Left            =   30
      TabIndex        =   7
      Top             =   390
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   3519
      IconSize        =   0
      CaptionColor    =   128
      Caption         =   "User Info"
      CaptionAlignment=   2
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtUserID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   510
         Width           =   2745
      End
      Begin VB.ComboBox cboLevel 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmUser.frx":0000
         Left            =   1680
         List            =   "frmUser.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1500
         Width           =   3015
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1170
         Width           =   2745
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   840
         Width           =   4395
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   90
         TabIndex        =   15
         Top             =   570
         Width           =   2085
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   14
         Top             =   1560
         Width           =   2085
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   13
         Top             =   1230
         Width           =   2085
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   8
         Top             =   900
         Width           =   2085
      End
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine2 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   503
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine1 
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjPatientSystem.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Save"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjPatientSystem.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Cancel"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjPatientSystem.lvButtons_H cmdHelp 
      Height          =   375
      Left            =   5940
      TabIndex        =   11
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjPatientSystem.lvButtons_H cmdHeader 
      Height          =   375
      Left            =   -30
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   661
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- Global Variables.
Public g_lAddState  As Long
Public g_lOldID     As Long
Public g_sOldUserID   As String
'

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdSave_Click()
    Dim sUserID As String

On Error GoTo errHandler

    '-- User ID.
    If Not isValidData(txtUserID, "User ID") Then Exit Sub
    
    sUserID = Trim$(txtUserID)
    
    '-- User Name.
    If Not isValidData(txtName, "User Name") Then Exit Sub
    
    '-- Level.
    If Not isValidData(cboLevel, "User Level") Then Exit Sub
    
    '-- Check for Unique UserID.
    If g_lAddState = 1 Then
        If isExisting(g_cn, "tblUser", "UserID='" & sUserID & "'") Then
            MsgBox "User User ID already exists.", vbExclamation
            Exit Sub
        End If
    Else
        '-- Check if user updated the User UserID.
        If InStr(1, sUserID, g_sOldUserID) = 0 Then
            If isExisting(g_cn, "tblUser", "UserID='" & sUserID & "' AND ID<>" & g_lOldID) Then
                MsgBox "User User ID already exists.", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    '-- Save Data.
    With g_rsUser
        If g_lAddState = 1 Then
            .AddNew
            .Fields("DateCreated") = g_sSysDate
        End If
        .Fields("UserID") = sUserID
        .Fields("Name") = txtName
        .Fields("Password") = txtPassword
        .Fields("Level") = cboLevel.ListIndex
        .Fields("DateModified") = g_sSysDate
        .Update
    End With
    
    If g_lAddState = 1 Then
        Dim lReply As Long
        lReply = MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo)
        If lReply = vbYes Then
            Call ClearAll(Me)
            txtUserID.SetFocus
        Else
            Call Unload(Me)
        End If
    Else
        Call PointToRecord(g_rsUser, "ID", False, g_lOldID, 0)
        If g_lOldID = g_tUser.ID Then
            g_tUser.UserID = txtUserID
            g_tUser.Password = txtPassword
            g_tUser.Level = cboLevel.ListIndex
        End If
        Call Unload(Me)
    End If

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmUser", "cmdSave_Click", True)
    End If
    
End Sub

Private Sub Form_Load()
    Call StartBusy
    
    Icon = frmMain.Icon
    Call CenterForm(Me)
    
    Call initToolbar(cmdHeader, frmMain.imlTrans)
    Call initToolbar(cmdSave, frmMain.imlTrans, , g_eIcon.Save)
    Call initToolbar(cmdCancel, frmMain.imlTrans, , g_eIcon.Cancel)
    Call initToolbar(cmdHelp, frmMain.imlTrans, , g_eIcon.Help)
    Call initFrame(fraBackground, frmMain.imlPanel, g_eIcon.User)
    
    Call EndBusy
End Sub

