VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin prjPatientSystem.ucGradContainer fraBackground 
      Height          =   2235
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   3942
      IconSize        =   0
      CaptionColor    =   128
      Caption         =   "Search Info"
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
      Begin VB.ComboBox cboOperator 
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
         ItemData        =   "frmSearch.frx":0000
         Left            =   1320
         List            =   "frmSearch.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   810
         Width           =   4725
      End
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
         Left            =   3690
         TabIndex        =   3
         Top             =   1740
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
         Left            =   4890
         TabIndex        =   4
         Top             =   1740
         WhatsThisHelpID =   3055
         Width           =   1140
      End
      Begin prjPatientSystem.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   60
         TabIndex        =   8
         Top             =   1620
         Width           =   6015
         _ExtentX        =   4260
         _ExtentY        =   53
      End
      Begin VB.TextBox txtText 
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
         Left            =   1320
         TabIndex        =   2
         Top             =   1140
         Width           =   4725
      End
      Begin VB.ComboBox cboField 
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
         ItemData        =   "frmSearch.frx":0098
         Left            =   1320
         List            =   "frmSearch.frx":009A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   4725
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Search Text"
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
         Top             =   1200
         Width           =   2085
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Search by"
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
         TabIndex        =   6
         Top             =   540
         Width           =   2085
      End
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- Global Variables.
Public g_rs         As ADODB.Recordset
Public g_lCancel    As Long
'

Private Sub cmdCancel_Click()
    Call Unload(Me)
    g_lCancel = 1
End Sub

Private Sub cmdOk_Click()
    Dim sCondition  As String
    
    sCondition = g_rs.Fields.Item(cboField.ItemData(cboField.ListIndex)).Name
    sCondition = Replace(sCondition, " ", "")
    sCondition = "[" & sCondition & "]"

    Select Case cboOperator.ListIndex
        Case 0: sCondition = sCondition & " LIKE '%" & txtText.Text & "%'"
        Case 1: sCondition = sCondition & " = '" & txtText.Text & "'"
        Case 2: sCondition = sCondition & " <> '" & txtText.Text & "'"
        Case 3: sCondition = sCondition & " > '" & txtText.Text & "'"
        Case 4: sCondition = sCondition & " >= '" & txtText.Text & "'"
        Case 5: sCondition = sCondition & " < '" & txtText.Text & "'"
        Case 6: sCondition = sCondition & " <= '" & txtText.Text & "'"
    End Select
    
    g_rs.Filter = sCondition
    g_rs.Requery
    
    Call Unload(Me)
End Sub

Private Sub Form_Load()
    
    Call CenterForm(Me)
    Icon = frmMain.Icon
    Call initFrame(fraBackground)

End Sub

