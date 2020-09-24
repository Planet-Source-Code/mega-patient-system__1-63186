VERSION 5.00
Begin VB.Form frmFee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fee"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin prjPatientSystem.ucHorizontal3DLine ucHorizontal3DLine1 
      Height          =   30
      Left            =   -360
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   53
   End
   Begin prjPatientSystem.ucGradContainer fraBackground 
      Height          =   1365
      Left            =   30
      TabIndex        =   5
      Top             =   390
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   2408
      IconSize        =   0
      CaptionColor    =   128
      Caption         =   "Fee Info"
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
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
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
         Left            =   1920
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   900
         Width           =   2115
      End
      Begin VB.TextBox txtDesc 
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
         Left            =   1920
         TabIndex        =   0
         Top             =   570
         Width           =   4305
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         TabIndex        =   7
         Top             =   930
         Width           =   2085
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         TabIndex        =   6
         Top             =   600
         Width           =   2085
      End
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine2 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   503
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine1 
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjPatientSystem.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   30
      TabIndex        =   2
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
      TabIndex        =   3
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
      Left            =   6120
      TabIndex        =   10
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
      TabIndex        =   11
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
Attribute VB_Name = "frmFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- Global Variables.
Public g_lAddState  As Long
Public g_lOldID     As Long
Public g_sOldDesc   As String
'

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdSave_Click()
    Dim sDesc As String

On Error GoTo errHandler

    sDesc = Trim$(txtDesc)

    '-- Required Field.
    If sDesc = vbNullString Then
        MsgBox "Kindly enter Fee Description", vbExclamation
        Exit Sub
    End If
    '-- Check for Unique Desc.
    If g_lAddState = 1 Then
        If isExisting(g_cn, "tblFee", "Desc='" & sDesc & "'") Then
            MsgBox "Fee Description already exists.", vbExclamation
            Exit Sub
        End If
    Else
        '-- Check if user updated the Fee Desc.
        If InStr(1, sDesc, g_sOldDesc) = 0 Then
            If isExisting(g_cn, "tblFee", "Desc='" & sDesc & "' AND ID<>" & g_lOldID) Then
                MsgBox "Fee Description already exists.", vbExclamation
                Exit Sub
            End If
        End If
    End If
    
    '-- Save Data.
    With g_rsFee
        If g_lAddState = 1 Then
            .AddNew
            .Fields("DateCreated") = g_sSysDate
        End If
        .Fields("Desc") = sDesc
        .Fields("Amount") = txtAmount
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
            txtDesc.SetFocus
        Else
            Unload Me
        End If
    Else
        Call PointToRecord(g_rsFee, "ID", False, g_lOldID, 0)
        Call Unload(Me)
    End If

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmFee", "cmdSave_Click", True)
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
    Call initFrame(fraBackground, frmMain.imlPanel, g_eIcon.Fee)
    
    Call EndBusy
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    Call Numeric(txtAmount, KeyAscii)
End Sub
