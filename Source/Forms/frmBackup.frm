VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5625
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin prjPatientSystem.ucHorizontal3DLine ucHorizontal3DLine1 
      Height          =   30
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2610
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   53
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   2940
      ScaleHeight     =   555
      ScaleWidth      =   2535
      TabIndex        =   6
      Top             =   2640
      Width           =   2535
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   360
         Left            =   1305
         TabIndex        =   4
         Top             =   120
         WhatsThisHelpID =   3055
         Width           =   1140
      End
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
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   90
      Width           =   5385
      Begin VB.PictureBox picBG 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2205
         Left            =   120
         ScaleHeight     =   2205
         ScaleWidth      =   5175
         TabIndex        =   8
         Top             =   180
         Width           =   5175
         Begin VB.TextBox txtPath 
            Height          =   315
            Left            =   0
            TabIndex        =   0
            Top             =   270
            Width           =   4725
         End
         Begin VB.CommandButton cmdPath 
            Caption         =   "..."
            Height          =   315
            Left            =   4770
            TabIndex        =   1
            Top             =   270
            Width           =   345
         End
         Begin VB.TextBox txtFName 
            ForeColor       =   &H80000012&
            Height          =   315
            Left            =   0
            MaxLength       =   20
            TabIndex        =   2
            Top             =   930
            Width           =   4725
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   315
            Left            =   0
            TabIndex        =   9
            Top             =   1350
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label Label1 
            Caption         =   "Select Path Where to Store Backup"
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
            Index           =   1
            Left            =   15
            TabIndex        =   12
            Top             =   0
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Back up File Name"
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
            Left            =   0
            TabIndex        =   11
            Top             =   690
            Width           =   3615
         End
         Begin VB.Label lblMsg 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Left            =   30
            TabIndex        =   10
            Top             =   1710
            Width           =   5115
         End
      End
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- Local Variables.
Private WithEvents Huffman As clsHuffman
Attribute Huffman.VB_VarHelpID = -1
'

Private Sub cmdOk_Click()

    If Not isValidData(txtPath, "Back Up File Path") Then Exit Sub
    If Not isValidData(txtFName, "Back Up File Name") Then Exit Sub
    
    Dim dest        As String
    Dim OldTimer    As Single
    
    lblMsg.Visible = True
    
    dest = txtPath.Text & "\" & txtFName & ".BK"
    
    ProgressBar1.Visible = True
    OldTimer = Timer
  
    Call Huffman.EncodeFile(App.Path & "\" & DB_NAME, dest)
    ProgressBar1.Value = 0
  
    Unload Me
    Exit Sub
    
End Sub

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdPath_Click()
    Me.Hide
    frmPath.Show vbModal
    Me.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        If ProgressBar1.Value = 0 Or ProgressBar1.Value Then
            Call Unload(Me)
        End If
    End If
End Sub

Private Sub Form_Load()

    Call CenterForm(Me)
    Icon = frmMain.Icon
    
    Set Huffman = New clsHuffman
    
    lblMsg.Visible = False
    ProgressBar1.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Huffman = Nothing
End Sub

Private Sub Huffman_Progress(Procent As Integer)
    lblMsg.Caption = "Compressing Database"
    ProgressBar1.Value = Procent
    If ProgressBar1.Value = 100 Then
        lblMsg.Caption = "Saving Compressed File ..."
    End If
    DoEvents
End Sub


