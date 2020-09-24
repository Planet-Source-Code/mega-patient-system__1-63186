VERSION 5.00
Begin VB.Form frmDate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Date"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   1
      Top             =   870
      WhatsThisHelpID =   3056
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      TabIndex        =   2
      Top             =   870
      WhatsThisHelpID =   3055
      Width           =   1140
   End
   Begin VB.TextBox txtDate 
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
      Left            =   1500
      TabIndex        =   0
      Top             =   180
      Width           =   1725
   End
   Begin prjPatientSystem.ucHorizontal3DLine ucHorizontal3DLine1 
      Height          =   30
      Left            =   90
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   690
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   53
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "System Date"
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
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1905
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdOk_Click()
    
    If Not isValidData(txtDate, "System Date", 2) Then Exit Sub
    
    g_sSysDate = txtDate
    frmMain.cmdDate.Caption = Format(g_sSysDate, FMT_DT)
    Call Unload(Me)
    
End Sub

Private Sub Form_Load()
    Icon = frmMain.Icon
    Call CenterForm(Me)
    
    txtDate = Format(g_sSysDate, FMT_DT)
End Sub

Private Sub txtDate_LostFocus()
On Error GoTo errHandler

    txtDate = Format(txtDate, FMT_DT)

errHandler:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub
