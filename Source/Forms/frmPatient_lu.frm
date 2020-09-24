VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmPatient_lu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patients"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8385
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
   ScaleHeight     =   6105
   ScaleWidth      =   8385
   StartUpPosition =   3  'Windows Default
   Begin prjPatientSystem.ucGradContainer fraBackground 
      Height          =   6045
      Left            =   30
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   30
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   10663
      IconSize        =   0
      CaptionColor    =   128
      Caption         =   "Search Patient"
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
      Begin prjPatientSystem.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   150
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   5430
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   53
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   5550
         Width           =   1140
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Default         =   -1  'True
         Height          =   360
         Left            =   5880
         TabIndex        =   2
         Top             =   5550
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   360
         Left            =   7080
         TabIndex        =   3
         Top             =   5550
         Width           =   1140
      End
      Begin VB.TextBox txtFind 
         Height          =   315
         Left            =   1770
         TabIndex        =   0
         Top             =   450
         Width           =   6405
      End
      Begin MSDataGridLib.DataGrid dgData 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   7858
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   3390.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2009.764
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2085.166
            EndProperty
         EndProperty
      End
      Begin VB.Label lblFind 
         BackStyle       =   0  'Transparent
         Caption         =   "Find by:"
         Height          =   285
         Left            =   150
         TabIndex        =   6
         Top             =   510
         Width           =   2085
      End
   End
End
Attribute VB_Name = "frmPatient_lu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- Local Variables.
Private m_rsPatient     As ADODB.Recordset
Private m_lFieldIdx     As Long
Private m_sSort         As String
'

'//---------------------------------------------------------------------------------------
'//--Procedure : ldRecords
'//--DateTime  : 8/12/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub ldRecords()
    
On Error GoTo errHandler
    Call StartBusy

    m_rsPatient.Requery
    With dgData
        Set .DataSource = m_rsPatient
        .Columns(0).Visible = False             '-- ID.
        .Columns(1).Caption = "Name"            '-- Name.
        .Columns(2).Caption = "Address"         '-- Address.
        .Columns(3).Caption = "Sex"             '-- Sex.
        .Columns(4).Caption = "Civil Status"    '-- Civil Status.
        .Columns(5).Caption = "Birthday"        '-- Birthday.
        .Columns(6).Caption = "Age"             '-- Age
        .Columns(7).Caption = "Contact No"      '-- ContactNo.
        .Columns(8).Visible = False             '-- Email.
        .Columns(9).Caption = "Occupation"      '-- Occupation.
        .Columns(10).Visible = False            '-- Udf1.
        .Columns(11).Visible = False            '-- Udf2.
        .Columns(12).Visible = False            '-- Udf3.
        .Columns(13).Visible = False            '-- Comments.
        .Columns(14).Visible = False            '-- Date Created.
        .Columns(15).Visible = False            '-- Date Modified.
        .Columns(16).Visible = False            '-- UserID
    End With

errHandler:
    Call EndBusy
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmPatient_lu", "ldRecords", True)
    End If
End Sub

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdOk_Click()
    
    '-- No selected Patient.
    If m_rsPatient.EOF Then Exit Sub

    With frmVisitation
        .txtName = m_rsPatient.Fields("Name")
        .txtName.Tag = m_rsPatient.Fields("ID")       '-- ID.
        .txtAddress = IIf(IsNull(m_rsPatient.Fields("Address")), vbNullString, m_rsPatient.Fields("Address"))
        .txtSex = IIf(IsNull(m_rsPatient.Fields("Sex")), vbNullString, m_rsPatient.Fields("Sex"))
        .txtCivilStatus = IIf(IsNull(m_rsPatient.Fields("CivilStatus")), vbNullString, m_rsPatient.Fields("CivilStatus"))
        .txtBirthday = IIf(IsNull(m_rsPatient.Fields("Birthday")), vbNullString, m_rsPatient.Fields("Birthday"))
        .txtAge = IIf(IsNull(m_rsPatient.Fields("Age")), vbNullString, m_rsPatient.Fields("Age"))
        .txtContactNo = IIf(IsNull(m_rsPatient.Fields("ContactNo")), vbNullString, m_rsPatient.Fields("ContactNo"))
        .txtOccupation = IIf(IsNull(m_rsPatient.Fields("Occupation")), vbNullString, m_rsPatient.Fields("Occupation"))
    End With
    
    Call Unload(Me)
    
End Sub

Private Sub cmdRefresh_Click()
    m_rsPatient.Filter = adFilterNone
    Call ldRecords
End Sub

Private Sub dgData_DblClick()
    Call cmdOk_Click
End Sub

Private Sub dgData_HeadClick(ByVal ColIndex As Integer)
    If m_lFieldIdx = ColIndex Then
        If m_sSort = "ASC" Then
            m_sSort = "DESC"
        Else
            m_sSort = "ASC"
        End If
    Else
        m_sSort = "ASC"
    End If
    m_lFieldIdx = ColIndex
    lblFind.Caption = "Find by " & dgData.Columns(ColIndex).Caption
    lblFind.Tag = dgData.Columns(ColIndex).DataField
    m_rsPatient.Sort = lblFind.Tag & " " & m_sSort
End Sub

Private Sub Form_Load()

    Icon = frmMain.Icon
    Call CenterForm(Me)
    
    Set fraBackground.Icon = frmMain.imlPanel.ListImages(g_eIcon.Patient).Picture
    Call initFrame(fraBackground)
    
    Set m_rsPatient = New ADODB.Recordset
    m_rsPatient.CursorLocation = adUseClient
    g_sSQL = "SELECT * FROM tblPatient ORDER BY Name"
    m_rsPatient.Open g_sSQL, g_cn, adOpenForwardOnly, adLockReadOnly
    
    Call ldRecords
    
    '-- Default.
    m_lFieldIdx = 1     '-- Name.
    m_sSort = "ASC"
    lblFind.Caption = "Find by " & dgData.Columns(m_lFieldIdx).Caption
    lblFind.Tag = dgData.Columns(m_lFieldIdx).DataField
    m_rsPatient.Sort = lblFind.Tag & " " & m_sSort
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_rsPatient = Nothing
End Sub

Private Sub txtFind_Change()
    If Trim$(txtFind) = vbNullString Then
        m_rsPatient.Filter = adFilterNone
    Else
        m_rsPatient.Filter = lblFind.Tag & " LIKE *" & txtFind.Text & "*"
    End If
    
    Call ldRecords
End Sub
