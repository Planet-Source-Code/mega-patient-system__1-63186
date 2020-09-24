VERSION 5.00
Begin VB.Form frmPatient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
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
   ScaleHeight     =   7245
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin prjPatientSystem.ucHorizontal3DLine ucHorizontal3DLine1 
      Height          =   30
      Left            =   -360
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   360
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   53
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine2 
      Height          =   285
      Left            =   2160
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   503
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine1 
      Height          =   255
      Left            =   1080
      TabIndex        =   17
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
      TabIndex        =   13
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
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   14
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
      Left            =   8520
      TabIndex        =   25
      TabStop         =   0   'False
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
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjPatientSystem.ucGradContainer fraBackground 
      Height          =   6825
      Left            =   30
      TabIndex        =   19
      Top             =   390
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   12039
      IconSize        =   0
      CaptionColor    =   128
      Caption         =   "Patient Info"
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
      Begin VB.TextBox txtOccupation 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Top             =   3900
         Width           =   3165
      End
      Begin VB.TextBox txtAge 
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Top             =   2910
         Width           =   1725
      End
      Begin VB.TextBox txtBirthday 
         Height          =   315
         Left            =   1920
         TabIndex        =   4
         Top             =   2580
         Width           =   3165
      End
      Begin VB.ComboBox cboCivilStatus 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2250
         Width           =   3195
      End
      Begin VB.ComboBox cboSex 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1920
         Width           =   3195
      End
      Begin VB.TextBox txtAddress 
         Height          =   1005
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   900
         Width           =   3975
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   570
         Width           =   3165
      End
      Begin VB.TextBox txtContactNo 
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   3240
         Width           =   3165
      End
      Begin VB.TextBox txtEMail 
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   3570
         Width           =   3165
      End
      Begin VB.TextBox txtUdf1 
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Top             =   4230
         Width           =   3165
      End
      Begin VB.TextBox txtUdf2 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   4560
         Width           =   6795
      End
      Begin VB.TextBox txtUdf3 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   4890
         Width           =   6795
      End
      Begin VB.TextBox txtComments 
         Height          =   1455
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   5220
         Width           =   6795
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         Height          =   285
         Left            =   60
         TabIndex        =   33
         Top             =   3930
         Width           =   2085
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   285
         Left            =   60
         TabIndex        =   32
         Top             =   2940
         Width           =   2085
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Birthday"
         Height          =   285
         Left            =   60
         TabIndex        =   31
         Top             =   2610
         Width           =   2085
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Civil Status"
         Height          =   285
         Left            =   60
         TabIndex        =   30
         Top             =   2280
         Width           =   2085
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         Height          =   285
         Left            =   60
         TabIndex        =   29
         Top             =   1950
         Width           =   2085
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   285
         Left            =   60
         TabIndex        =   28
         Top             =   930
         Width           =   2085
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   60
         TabIndex        =   27
         Top             =   600
         Width           =   2085
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No"
         Height          =   285
         Left            =   60
         TabIndex        =   26
         Top             =   3270
         Width           =   2085
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail"
         Height          =   285
         Left            =   60
         TabIndex        =   24
         Top             =   3600
         Width           =   2085
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Userdefined field1"
         Height          =   285
         Left            =   60
         TabIndex        =   23
         Top             =   4260
         Width           =   2085
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Userdefined field2"
         Height          =   285
         Left            =   60
         TabIndex        =   22
         Top             =   4590
         Width           =   2085
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Userdefined field3"
         Height          =   285
         Left            =   60
         TabIndex        =   21
         Top             =   4920
         Width           =   2085
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         Height          =   285
         Left            =   60
         TabIndex        =   20
         Top             =   5250
         Width           =   2085
      End
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine3 
      Height          =   255
      Left            =   3330
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjPatientSystem.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   2160
      TabIndex        =   34
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Print"
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   8925
      _ExtentX        =   15743
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
Attribute VB_Name = "frmPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- Global Variables.
Public g_lAddState  As Long
Public g_lOldID     As Long
Public g_sOldName   As String
'

'//---------------------------------------------------------------------------------------
'//--Procedure : init
'//--DateTime  : 7/22/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   : Initialize.
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Public Sub init()

On Error GoTo errHandler
    
    '-- Fill with fixed Data.
    cboSex.AddItem SX_FEMALE
    cboSex.AddItem SX_MALE
    cboCivilStatus.AddItem CS_SINGLE
    cboCivilStatus.AddItem CS_MARRIED
    cboCivilStatus.AddItem CS_WIDOW
    cboCivilStatus.AddItem CS_SEP

errHandler:
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdSave_Click()
    Dim sName As String

On Error GoTo errHandler

    '-- Required Field.
    If Not isValidData(txtName, "Patient Name") Then Exit Sub
    
    sName = Trim$(txtName)
    
    '-- Valid Birthday.
    If Trim$(txtBirthday) <> vbNullString Then
        If Not isValidData(txtBirthday, " Birthday", 2) Then Exit Sub
    End If
    
    '-- Age.
    If Trim$(txtAge) <> vbNullString Then
        If Not isValidData(txtAge, "Age", 1) Then Exit Sub
    End If
    
    '-- Check for Unique Name.
    If g_lAddState = 1 Then
        If isExisting(g_cn, "tblPatient", "Name='" & sName & "'") Then
            MsgBox "Patient Name already exists.", vbExclamation
            Exit Sub
        End If
    Else
        '-- Check if user updated the Patient Name.
        If InStr(1, sName, g_sOldName) = 0 Then
            If isExisting(g_cn, "tblPatient", "Name='" & sName & "' AND ID<>" & g_lOldID) Then
                MsgBox "Patient Name already exists.", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    '-- Save Data.
    With g_rsPatient
        If g_lAddState = 1 Then
            .AddNew
            .Fields("DateCreated") = g_sSysDate
        End If
        .Fields("Name") = sName
        .Fields("Address") = txtAddress
        .Fields("Sex") = cboSex
        .Fields("CivilStatus") = cboCivilStatus
        If Trim$(txtBirthday) <> vbNullString Then .Fields("Birthday") = txtBirthday
        If Trim$(txtAge) <> vbNullString Then .Fields("Age") = txtAge
        .Fields("ContactNo") = txtContactNo
        .Fields("EMail") = txtEMail
        .Fields("Occupation") = txtOccupation
        .Fields("Udf1") = txtUdf1
        .Fields("Udf2") = txtUdf2
        .Fields("Udf3") = txtUdf3
        .Fields("Comments") = txtComments
        .Fields("DateModified") = g_sSysDate
        .Fields("UserID") = g_tUser.ID
        .Update
    End With
    
    '-- Prompt user.
    If g_lAddState = 1 Then
        Dim lReply As Long
        lReply = MsgBox("Do you want to add a new record?", vbQuestion + vbYesNo)
        If lReply = vbYes Then
            Call ClearAll(Me)
            txtName.SetFocus
        Else
            Unload Me
        End If
    Else
        Call PointToRecord(g_rsPatient, "ID", False, g_lOldID, 0)
        Call Unload(Me)
    End If

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmPatient", "cmdSave_Click", True)
    End If
    
End Sub

Private Sub Form_Load()
    Call StartBusy

    Icon = frmMain.Icon
    Call CenterForm(Me)
    
    Call init
    
    Call initToolbar(cmdHeader, frmMain.imlTrans)
    Call initToolbar(cmdSave, frmMain.imlTrans, , g_eIcon.Save)
    Call initToolbar(cmdCancel, frmMain.imlTrans, , g_eIcon.Cancel)
    Call initToolbar(cmdPrint, frmMain.imlTrans, , g_eIcon.Print)
    Call initToolbar(cmdHelp, frmMain.imlTrans, , g_eIcon.Help)
    Call initFrame(fraBackground, frmMain.imlPanel, g_eIcon.Patient)

    Call EndBusy
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case "48" To "57", vbKeyBack                        '-- Numbers 0-9, Backspace
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtBirthday_LostFocus()
    txtBirthday = Format(txtBirthday, FMT_DT)
End Sub
