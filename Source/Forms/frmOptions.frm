VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5595
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
   ScaleHeight     =   4380
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3525
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   5385
      Begin VB.CheckBox chkOptions 
         Caption         =   "Remember Last Login"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Top             =   540
         Width           =   1875
      End
      Begin VB.CheckBox chkOptions 
         Caption         =   "Show Splash"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   240
         Width           =   1995
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   2940
      ScaleHeight     =   555
      ScaleWidth      =   2535
      TabIndex        =   1
      Top             =   3750
      Width           =   2535
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   360
         Left            =   90
         TabIndex        =   3
         Top             =   120
         WhatsThisHelpID =   3056
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   360
         Left            =   1305
         TabIndex        =   2
         Top             =   120
         WhatsThisHelpID =   3055
         Width           =   1140
      End
   End
   Begin prjPatientSystem.ucHorizontal3DLine ucHorizontal3DLine1 
      Height          =   30
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3720
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   53
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdOk_Click()
        
    '-- Update values.
    g_tSetting.Splash = chkOptions(0).Value
    g_tSetting.LastLog = chkOptions(1).Value
    
    '-- Update ini file.
    Call SaveINI(INI_NAME, "Option", "Splash", CStr(g_tSetting.Splash))
    Call SaveINI(INI_NAME, "Option", "LastLog", CStr(g_tSetting.LastLog))
    
    Call Unload(Me)

End Sub

Private Sub Form_Load()
    
    Call CenterForm(Me)
    Icon = frmMain.Icon
    
    '-- Set values.
    chkOptions(0).Value = g_tSetting.Splash
    chkOptions(1).Value = g_tSetting.LastLog
    
End Sub
