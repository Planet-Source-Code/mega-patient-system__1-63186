VERSION 5.00
Begin VB.Form frmPath 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Path"
   ClientHeight    =   4770
   ClientLeft      =   7095
   ClientTop       =   3420
   ClientWidth     =   5925
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2120
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5925
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   2910
      ScaleHeight     =   555
      ScaleWidth      =   2925
      TabIndex        =   4
      Top             =   4170
      Width           =   2925
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   360
         Left            =   510
         TabIndex        =   6
         Top             =   120
         WhatsThisHelpID =   3056
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   360
         Left            =   1725
         TabIndex        =   5
         Top             =   120
         WhatsThisHelpID =   3055
         Width           =   1140
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   135
      TabIndex        =   1
      Top             =   525
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   2415
   End
   Begin prjPatientSystem.ucHorizontal3DLine ucHorizontal3DLine1 
      Height          =   30
      Left            =   120
      TabIndex        =   7
      Top             =   4140
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   53
   End
   Begin VB.Label lblFileName 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   2760
      TabIndex        =   2
      Top             =   525
      Width           =   3000
   End
   Begin VB.Label lblPath 
      Caption         =   "Current Selected Path "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   3
      Top             =   225
      Width           =   2895
   End
End
Attribute VB_Name = "frmPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdOk_Click()
    frmBackup.txtPath.Text = lblFileName.Caption
    Call Unload(Me)
End Sub

Private Sub Dir1_Change()
    lblFileName.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo A1:
    Dir1.Path = Drive1.Drive
    Exit Sub
A1:
    MsgBox "Drive Can not be Accessed ...", vbInformation, "Drive not Accessed ..."
    Drive1.Drive = "c:"
End Sub

Private Sub Form_Load()
    Drive1.Drive = "c:"
    lblFileName.Caption = Dir1.Path
    Call CenterForm(Me)
End Sub

