VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8109F586-E9B5-4B96-994B-493253626B1C}#1.0#0"; "XTab.ocx"
Begin VB.Form frmVisitation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visitation"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10515
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
   ScaleHeight     =   8925
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin prjPatientSystem.ucHorizontal3DLine ucHorizontal3DLine1 
      Height          =   30
      Left            =   -360
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   53
   End
   Begin prjPatientSystem.ucGradContainer fraBG 
      Height          =   8505
      Left            =   30
      TabIndex        =   7
      Top             =   390
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   15002
      IconSize        =   0
      CaptionColor    =   128
      Caption         =   "Visitation Info"
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
      Begin VB.TextBox txtCivilStatus 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2130
         Width           =   3165
      End
      Begin VB.TextBox txtSex 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1800
         Width           =   3165
      End
      Begin VB.TextBox txtOccupation 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   7140
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1410
         Width           =   3165
      End
      Begin VB.TextBox txtAge 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   7140
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   750
         Width           =   1725
      End
      Begin VB.TextBox txtBirthday 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   7140
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   420
         Width           =   3165
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H8000000F&
         Height          =   1005
         Left            =   1410
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   780
         Width           =   3975
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   450
         Width           =   3165
      End
      Begin VB.TextBox txtContactNo 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   7140
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   3165
      End
      Begin VB.TextBox txtVisitationDate 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7170
         TabIndex        =   8
         Top             =   2130
         Width           =   1725
      End
      Begin prjPatientSystem.lvButtons_H cmdPatient_lu 
         Height          =   315
         Left            =   4590
         TabIndex        =   15
         Top             =   450
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
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
      Begin prjXTab.XTab tbBG 
         Height          =   5955
         Left            =   30
         TabIndex        =   16
         Top             =   2520
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   10504
         TabCount        =   2
         TabCaption(0)   =   "&Diagnosis Info"
         TabAccessKey(0) =   68
         TabContCtrlCnt(0)=   1
         Tab(0)ContCtrlCap(1)=   "picDiagnosis"
         TabCaption(1)   =   "&Fee Info"
         TabAccessKey(1) =   70
         TabContCtrlCnt(1)=   1
         Tab(1)ContCtrlCap(1)=   "picFee"
         ActiveTab       =   1
         TabStyle        =   1
         TabTheme        =   2
         ActiveTabBackStartColor=   16777215
         InActiveTabBackStartColor=   16777215
         InActiveTabBackEndColor=   -2147483626
         ActiveTabForeColor=   128
         InActiveTabForeColor=   -2147483631
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterBorderColor=   16777215
         DisabledTabBackColor=   -2147483633
         DisabledTabForeColor=   -2147483627
         Begin VB.PictureBox picDiagnosis 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5505
            Left            =   -74880
            ScaleHeight     =   5505
            ScaleWidth      =   10185
            TabIndex        =   24
            Top             =   420
            Width           =   10185
            Begin VB.TextBox txtDiagnosis 
               Height          =   5295
               Left            =   60
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   25
               Top             =   120
               Width           =   10065
            End
         End
         Begin VB.PictureBox picFee 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5535
            Left            =   90
            ScaleHeight     =   5535
            ScaleWidth      =   10215
            TabIndex        =   17
            Top             =   390
            Width           =   10215
            Begin VB.CommandButton cmdAdd 
               Caption         =   "Add"
               Height          =   360
               Left            =   30
               TabIndex        =   22
               Top             =   90
               Width           =   1140
            End
            Begin VB.TextBox txtTotalAmount 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7290
               Locked          =   -1  'True
               TabIndex        =   21
               Text            =   "0.00"
               Top             =   5160
               Width           =   2895
            End
            Begin VB.CommandButton cmdEdit 
               Caption         =   "Edit"
               Enabled         =   0   'False
               Height          =   360
               Left            =   30
               TabIndex        =   19
               Top             =   510
               Width           =   1140
            End
            Begin VB.CommandButton cmdDelete 
               Caption         =   "Delete"
               Enabled         =   0   'False
               Height          =   360
               Left            =   30
               TabIndex        =   18
               Top             =   930
               Width           =   1140
            End
            Begin MSComctlLib.ListView lvwFee 
               Height          =   4995
               Left            =   1260
               TabIndex        =   20
               Top             =   90
               Width           =   8925
               _ExtentX        =   15743
               _ExtentY        =   8811
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   3
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Fee "
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Comment"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Amount"
                  Object.Width           =   2540
               EndProperty
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total Amount"
               Height          =   285
               Left            =   6060
               TabIndex        =   23
               Top             =   5250
               Width           =   1125
            End
         End
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
         Height          =   285
         Left            =   5850
         TabIndex        =   34
         Top             =   1500
         Width           =   2085
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   285
         Left            =   5850
         TabIndex        =   33
         Top             =   840
         Width           =   2085
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Birthday"
         Height          =   285
         Left            =   5850
         TabIndex        =   32
         Top             =   510
         Width           =   2085
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Civil Status"
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   2190
         Width           =   2085
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   1860
         Width           =   2085
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   810
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
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   2085
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No"
         Height          =   285
         Left            =   5850
         TabIndex        =   27
         Top             =   1170
         Width           =   2085
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   5850
         TabIndex        =   26
         Top             =   2220
         Width           =   2085
      End
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine2 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   503
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine1 
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjPatientSystem.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   1
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
      Left            =   10140
      TabIndex        =   5
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
   Begin prjPatientSystem.lvButtons_H cmdSave 
      Height          =   375
      Left            =   0
      TabIndex        =   0
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
   Begin prjPatientSystem.lvButtons_H cmdHeader 
      Height          =   375
      Left            =   -30
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
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
Attribute VB_Name = "frmVisitation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- Global Variables.
Public g_lAddState          As Long
Public g_lOldID             As Long

'-- Local Variables.
Private m_rsVisitation      As ADODB.Recordset
Private m_rsPatient         As ADODB.Recordset
Private m_rsFee             As ADODB.Recordset
Private m_lPatientID        As Long
'

'//---------------------------------------------------------------------------------------
'//--Procedure : ComputeTotal
'//--DateTime  : 8/11/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Public Sub ComputeTotal()
    Dim lCtr        As Long
    Dim cTotal      As Currency

On Error GoTo errHandler

    For lCtr = 1 To lvwFee.ListItems.Count
        cTotal = cTotal + lvwFee.ListItems(lCtr).SubItems(2)
    Next
    
    txtTotalAmount = Format(cTotal, FMT_CUR)

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmVisitation", "ComputeTotal", True)
    End If
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : CountCheckedItems
'//--DateTime  : 8/11/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Function CountCheckedItems() As Long
    Dim lvw         As ListItem
    Dim lCount      As Long
    
On Error GoTo errHandler
    
    For Each lvw In lvwFee.ListItems
        If lvw.Checked Then lCount = lCount + 1
    Next
    
    CountCheckedItems = lCount

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmVisitation", "CountCheckedItems", True)
    End If
End Function

'//---------------------------------------------------------------------------------------
'//--Procedure : initListView
'//--DateTime  : 8/11/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub initListView()

On Error GoTo errHandler

    With lvwFee
        .ColumnHeaders(1).Width = .Width / 3
        .ColumnHeaders(2).Width = .Width / 3
        .ColumnHeaders(3).Width = .Width / 3
        .ColumnHeaders(3).Alignment = lvwColumnRight
    End With

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmVisitation", "initListView", True)
    End If
End Sub


'//---------------------------------------------------------------------------------------
'//--Procedure : getVisitationID
'//--DateTime  : 8/12/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   : Get the next Visitation ID.
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Function getVisitationID() As Long
    Dim rs As ADODB.Recordset
    
On Error GoTo errHandler

    Set rs = g_cn.Execute("SELECT Max(ID) + 1 AS NextID FROM tblVisitation")
    If Not rs.EOF Then
        getVisitationID = rs.Fields("NextID")
    Else
        getVisitationID = 1
    End If

errHandler:
    Set rs = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Function


Private Sub cmdAdd_Click()
    With frmFee_ae
        .g_lAddState = 1
        .Show vbModal
    End With
End Sub

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdDelete_Click()
    Dim lvw         As ListItem
    Dim lResult     As Long
    
    lResult = MsgBox("Are you sure you want to delete?", vbQuestion + vbYesNo)
    If lResult = vbYes Then
    
        For Each lvw In lvwFee.ListItems
            If lvw.Checked Then
                lvwFee.ListItems.Remove lvw.Index
            End If
        Next
        
    End If
    
End Sub

Private Sub cmdEdit_Click()
    Dim lvw As ListItem
    
    For Each lvw In lvwFee.ListItems
        If lvw.Checked Then
            With frmFee_ae
                .g_lAddState = 0
                .g_lIdx = lvw.Index
                .cboDesc = lvw.Text
                .txtAmount = lvw.SubItems(2)
                .txtComments = lvw.SubItems(1)
                .Show vbModal
            End With
            Exit For
        End If
    Next
    
End Sub

Private Sub cmdPatient_lu_Click()
    frmPatient_lu.Show vbModal
End Sub

Private Sub cmdSave_Click()
    
On Error GoTo errHandler

    '-- Patient.
    If Not isValidData(txtName, "Patient Name") Then Exit Sub
    
    '-- Visitation Date.
    If Not isValidData(txtVisitationDate, "Visitation Date", 2) Then Exit Sub
    
    '-- Diagnosis.
    If Not isValidData(txtDiagnosis, "Diagnosis Info") Then Exit Sub
    
    '-- Fee.
    If lvwFee.ListItems.Count <= 0 Then
        MsgBox "Kindly enter Fee Info.", vbExclamation
        tbBG.ActiveTab = 1
        cmdAdd.SetFocus
        Exit Sub
    End If
    
    Dim lTrans          As Long
    Dim lvw             As ListItem
    Dim lVisitationID   As Long
    
    '-- Start of transaction.
    g_cn.BeginTrans
    lTrans = 1
    
    '-- Save Data.
    With m_rsVisitation
        
        '-- Visitation.
        If g_lAddState = 1 Then
            .AddNew
            .Fields("DateCreated") = Date
            lVisitationID = getVisitationID
        Else
            lVisitationID = g_lOldID
        End If
        .Fields("PatientID") = txtName.Tag
        .Fields("Diagnosis") = Trim$(txtDiagnosis)
        .Fields("DateModified") = Date
        .Fields("TotalAmount") = txtTotalAmount
        .Fields("VisitationDate") = txtVisitationDate
        .Fields("UserID") = g_tUser.ID
        .Update
        
        '-- Visitation Fee.
        If g_lAddState = 0 Then g_cn.Execute "DELETE * FROM tblVisitation_Fee WHERE VisitationID = " & g_lOldID
        For Each lvw In lvwFee.ListItems
            g_cn.Execute "INSERT INTO tblVisitation_Fee (VisitationID, FeeID, Amount, Comments) VALUES(" & _
                lVisitationID & "," & lvw.Tag & "," & lvw.SubItems(2) & ",'" & lvw.SubItems(1) & "')"
        Next
        
    End With
    
    '-- End of transcation.
    g_cn.CommitTrans
    lTrans = 0
    
    Call Unload(Me)
    
errHandler:
    If Err.Number <> 0 Then
        If lTrans = 1 Then g_cn.RollbackTrans
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmVisitation", "cmdSave_Click", True)
    End If
    
End Sub

Private Sub Form_Load()
On Error GoTo errHandler

    Call StartBusy

    Icon = frmMain.Icon
    Call CenterForm(Me)
    
    '-- Initialize.
    Call initFrame(fraBG, frmMain.imlPanel, g_eIcon.Visitation)
    Call initToolbar(cmdHeader, frmMain.imlTrans)
    Call initToolbar(cmdSave, frmMain.imlTrans, , g_eIcon.Save)
    Call initToolbar(cmdCancel, frmMain.imlTrans, , g_eIcon.Cancel)
    Call initToolbar(cmdHelp, frmMain.imlTrans, , g_eIcon.Help)
    Call initButton(cmdPatient_lu, frmMain.imlTrans, g_eIcon.Search, lv_Fill_Stretch)
    Call initListView
    
    tbBG.ActiveTab = 0
    
    Set m_rsPatient = New ADODB.Recordset
    Set m_rsVisitation = New ADODB.Recordset
    Set m_rsFee = New ADODB.Recordset
    m_rsVisitation.CursorLocation = adUseClient
    m_rsPatient.CursorLocation = adUseClient
    m_rsFee.CursorLocation = adUseClient
    
    '-- New State.
    If g_lAddState = 1 Then
        
        txtVisitationDate = Format(Date, FMT_DT)
        g_sSQL = "SELECT * FROM tblVisitation"
        m_rsVisitation.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
        
    Else    '-- Edit State.
        
        g_sSQL = "SELECT * FROM tblVisitation WHERE ID = " & g_lOldID
        m_rsVisitation.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
        
        '-- Load Records.
        If Not m_rsVisitation.EOF Then
            
            '-- Patient Info.
            g_sSQL = "SELECT * FROM tblPatient"
            m_rsPatient.Open g_sSQL, g_cn, adOpenForwardOnly, adLockReadOnly
            m_lPatientID = m_rsVisitation.Fields("PatientID")
            m_rsPatient.Find "ID = " & m_lPatientID
            If Not m_rsPatient.EOF Then
                txtName = m_rsPatient.Fields("Name")
                txtName.Tag = m_lPatientID
                txtAddress = IIf(IsNull(m_rsPatient.Fields("Address")), vbNullString, m_rsPatient.Fields("Address"))
                txtSex = IIf(IsNull(m_rsPatient.Fields("Sex")), vbNullString, m_rsPatient.Fields("Sex"))
                txtCivilStatus = IIf(IsNull(m_rsPatient.Fields("CivilStatus")), vbNullString, m_rsPatient.Fields("CivilStatus"))
                txtBirthday = IIf(IsNull(m_rsPatient.Fields("Birthday")), vbNullString, m_rsPatient.Fields("Birthday"))
                txtAge = IIf(IsNull(m_rsPatient.Fields("Age")), vbNullString, m_rsPatient.Fields("Age"))
            End If
            
            '-- Visitation Info.
            txtVisitationDate = m_rsVisitation.Fields("VisitationDate")
            txtDiagnosis = m_rsVisitation.Fields("Diagnosis")
            
            '-- Fee Info.
            Dim lIdx As Long
            g_sSQL = "SELECT * FROM qryVisitation_Fee WHERE VisitationID = " & g_lOldID
            m_rsFee.Open g_sSQL, g_cn, adOpenForwardOnly, adLockReadOnly
            If Not m_rsFee.EOF Then
                While Not m_rsFee.EOF
                    With lvwFee
                        .ListItems.Add Text:=m_rsFee.Fields("Desc")
                        lIdx = .ListItems.Count
                        .ListItems.Item(lIdx).Tag = m_rsFee.Fields("FeeID")
                        .ListItems.Item(lIdx).SubItems(1) = IIf(IsNull(m_rsFee.Fields("Comments")), vbNullString, m_rsFee.Fields("Comments"))
                        .ListItems.Item(lIdx).SubItems(2) = Format(m_rsFee.Fields("Amount"), FMT_CUR)
                    End With
                    m_rsFee.MoveNext
                Wend
            End If
            
        End If
        
    End If
    


errHandler:

    Call EndBusy
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmVisitation", "Form_Load", True)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_rsPatient = Nothing
    Set m_rsVisitation = Nothing
End Sub

Private Sub lvwFee_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Item.Selected = True
    Call lvwFee_ItemClick(Item)
End Sub

Private Sub lvwFee_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If CountCheckedItems <= 0 Then
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    ElseIf CountCheckedItems = 1 Then
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
    Else '-- >1
        cmdEdit.Enabled = False
        cmdDelete.Enabled = True
    End If
End Sub

