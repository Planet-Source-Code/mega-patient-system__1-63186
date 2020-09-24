VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRestore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Restore Database"
   ClientHeight    =   5190
   ClientLeft      =   1635
   ClientTop       =   4755
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2130
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   5760
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5835
      Left            =   120
      ScaleHeight     =   5835
      ScaleWidth      =   5895
      TabIndex        =   5
      Top             =   90
      Width           =   5895
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   360
         Left            =   4380
         TabIndex        =   4
         Top             =   4650
         WhatsThisHelpID =   3055
         Width           =   1140
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Default         =   -1  'True
         Height          =   360
         Left            =   3180
         TabIndex        =   3
         Top             =   4650
         WhatsThisHelpID =   3056
         Width           =   1140
      End
      Begin prjPatientSystem.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4500
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         TabIndex        =   0
         Top             =   360
         Width           =   2565
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   45
         TabIndex        =   1
         Top             =   735
         Width           =   2565
      End
      Begin VB.FileListBox File1 
         Height          =   1455
         Left            =   2745
         TabIndex        =   2
         Top             =   735
         Width           =   2715
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   0
         TabIndex        =   6
         Top             =   4050
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   556
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Selected Back Up File Path"
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
         Left            =   30
         TabIndex        =   10
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label lblFileName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   30
         TabIndex        =   9
         Top             =   3150
         Width           =   5190
      End
      Begin VB.Label Label1 
         Caption         =   "Select Backed Up File and Click on the Ok Button"
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
         Index           =   2
         Left            =   30
         TabIndex        =   8
         Top             =   0
         Width           =   5415
      End
      Begin VB.Label lblMsg 
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
         TabIndex        =   7
         Top             =   3720
         Width           =   5445
      End
   End
End
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- Local Variables.
Private WithEvents Huffman  As clsHuffman
Attribute Huffman.VB_VarHelpID = -1
Dim sFileName               As String
'

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdOk_Click()
  
    Dim Frm         As Form
    Dim OldTimer    As Single
    
On Error GoTo errHandler

    If Trim$(lblFileName.Caption) = vbNullString Then
        MsgBox "Select Backup File Path and then Click Ok Button...", vbInformation, "Select Path ..."
        Exit Sub
    Else
        '-- Check if valid file name.
        If UCase(Right$(sFileName, 3)) <> ".BK" Then
            MsgBox "Kindly change back up file." & vbCr & vbCr & "Not a valid back up file.", vbExclamation, "Valid Filename"
            File1.SetFocus
            Exit Sub
        End If
    End If
  
    ProgressBar1.Visible = True
    lblMsg.Visible = True
  
    OldTimer = Time
    
    '-- Close connection.
    g_cn.Close
    
    '-- Decompress DB.
    Call Huffman.DecodeFile(lblFileName.Caption, App.Path & "\" & DB_NAME)
    
    '-- DB Connection.
    'Set g_cn = New ADODB.Connection
    g_cn.Open g_sConnectionString
    
    '-- Reload Recordsets.
    frmMain.Form_Load
    
    ProgressBar1.Value = 0
  
    MsgBox "Database succesfully restored.", vbInformation
  
    Call Unload(Me)
    Exit Sub

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmRestore", "cmdOk_Click", True)
    End If
    
End Sub

Private Sub Dir1_Change()
    lblFileName.Caption = ""
On Error GoTo A1:
    File1.Path = Dir1.Path
    Exit Sub
A1:
    MsgBox "Folder Can not be Accessed ...", vbInformation, "Drive not Accessed ..."
    Drive1.Drive = "c:"
End Sub

Private Sub Drive1_Change()
    lblFileName.Caption = ""
On Error GoTo A1:
    Dir1.Path = Drive1.Drive
    Exit Sub
A1:
    MsgBox "Drive Can not be Accessed ...", vbInformation, "Drive not Accessed ..."
    Drive1.Drive = "c:"

End Sub

Private Sub File1_Click()
    sFileName = File1.FileName
    lblFileName.Caption = File1.Path & "\" & File1.FileName
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
    
    lblFileName.Caption = ""
    lblMsg.Visible = False
    ProgressBar1.Visible = False
    
    Set Huffman = New clsHuffman
    
    Drive1.Drive = "c:"
    
End Sub

Private Sub Huffman_Progress(Procent As Integer)
    lblMsg.Caption = "Uncompressing Database"
    ProgressBar1.Value = Procent
    If ProgressBar1.Value = 100 Then
        lblMsg.Caption = "Restoring Uncompressed File ..."
        End If
    DoEvents
End Sub

