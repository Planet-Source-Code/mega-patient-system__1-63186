VERSION 5.00
Begin VB.Form frmFee_ae 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fee"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7785
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
   ScaleHeight     =   2805
   ScaleWidth      =   7785
   StartUpPosition =   3  'Windows Default
   Begin prjPatientSystem.ucGradContainer fraBG 
      Height          =   2745
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   4842
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
      Begin prjPatientSystem.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2100
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   53
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1950
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   840
         Width           =   2745
      End
      Begin VB.TextBox txtComments 
         Height          =   795
         Left            =   1950
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1170
         Width           =   5655
      End
      Begin VB.ComboBox cboDesc 
         Height          =   315
         Left            =   1950
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   510
         Width           =   5655
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   360
         Left            =   6465
         TabIndex        =   4
         Top             =   2250
         WhatsThisHelpID =   3055
         Width           =   1140
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   360
         Left            =   5250
         TabIndex        =   3
         Top             =   2250
         WhatsThisHelpID =   3056
         Width           =   1140
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         TabIndex        =   8
         Top             =   870
         Width           =   2085
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fee Description"
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
         TabIndex        =   7
         Top             =   540
         Width           =   2085
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   2085
      End
   End
End
Attribute VB_Name = "frmFee_ae"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- Global Variables.
Public g_lAddState  As Long
Public g_lIdx       As Long

'-- local Variables.
Private m_rsFee     As ADODB.Recordset
'

'//---------------------------------------------------------------------------------------
'//--Procedure : ldFee
'//--DateTime  : 8/11/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   : Load fee into the combobox.
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub ldFee()
    
On Error GoTo errHandler

    cboDesc.Clear

    m_rsFee.MoveFirst
    While Not m_rsFee.EOF
        cboDesc.AddItem m_rsFee.Fields("Desc")
        cboDesc.ItemData(cboDesc.NewIndex) = m_rsFee.Fields("ID")
        m_rsFee.MoveNext
    Wend

errHandler:
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : isFeeExist
'//--DateTime  : 8/11/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   : Check if a fee aleady exists in the list.
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Function isFeeExist(lFeeID As Long) As Long
    Dim lvw As ListItem
    
On Error GoTo errHandler

    With frmVisitation
        For Each lvw In .lvwFee.ListItems
            If lvw.Tag = lFeeID Then
                isFeeExist = 1
                Exit For
            End If
        Next
    End With

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmFee_ae", "isFeeExist", True)
    End If
End Function

Private Sub cboDesc_Change()
    If cboDesc.ListIndex = -1 Then Exit Sub
    m_rsFee.Requery
    m_rsFee.Find "ID=" & cboDesc.ItemData(cboDesc.ListIndex)
    If Not m_rsFee.EOF Then
        txtAmount = Format(m_rsFee.Fields("Amount"), FMT_CUR)
    End If
End Sub

Private Sub cboDesc_Click()
    Call cboDesc_Change
End Sub

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdOk_Click()
    Dim lResult     As Long
        
On Error GoTo errHandler

    '-- Fee
    If cboDesc.ListIndex = -1 Then
        MsgBox "Kindly enter Fee.", vbExclamation
        cboDesc.SetFocus
        Exit Sub
    Else
        '-- New state.
        If g_lAddState = 1 Then
            '-- Check if fee alreadys exists.
            If isFeeExist(cboDesc.ItemData(cboDesc.ListIndex)) = 1 Then
                lResult = MsgBox(cboDesc & " already exists." & vbCr & vbCr & _
                    "Do you want to continue?", vbQuestion + vbYesNo)
                If lResult = vbNo Then Exit Sub
            End If
        End If
    End If
    
    '-- Amount
    If Trim$(txtAmount) = vbNullString Then
        MsgBox "Kindly enter Amount.", vbExclamation
        txtAmount.SetFocus
        Exit Sub
    Else    '-- Check for valid amount.
        If CCur(txtAmount) <= 0 Then
            MsgBox "That is not a valid amount.", vbExclamation
            txtAmount.SetFocus
            Exit Sub
        End If
    End If

    Dim lidx        As Long
    
    '-- Save Data.
    With frmVisitation
        
        If g_lAddState = 1 Then
        
            '-- Add Fee.
            .lvwFee.ListItems.Add Text:=cboDesc
            lidx = .lvwFee.ListItems.Count
            .lvwFee.ListItems.Item(lidx).Tag = cboDesc.ItemData(cboDesc.ListIndex)
            .lvwFee.ListItems.Item(lidx).SubItems(1) = txtComments
            .lvwFee.ListItems.Item(lidx).SubItems(2) = Format(txtAmount, FMT_CUR)
            
        Else
            
            '-- Update Fee.
            lidx = g_lIdx
            .lvwFee.ListItems(lidx).Text = cboDesc
            .lvwFee.ListItems(lidx).Tag = cboDesc.ItemData(cboDesc.ListIndex)
            .lvwFee.ListItems(lidx).SubItems(1) = txtComments
            .lvwFee.ListItems(lidx).SubItems(2) = Format(txtAmount, FMT_CUR)
            
        End If
        
        '-- Re-compute total amount.
        .ComputeTotal
        
    End With
    
    '-- Prompt user.
    If g_lAddState = 1 Then
        lResult = MsgBox("Add another fee?", vbQuestion + vbYesNo)
        If lResult = vbYes Then
            cboDesc.ListIndex = -1
            txtComments = vbNullString
            txtAmount = "0.00"
            cboDesc.SetFocus
            Exit Sub
        End If
    End If
    
    Call Unload(Me)

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmFee_ae", "cmdOk_Click", True)
    End If
    
End Sub

Private Sub Form_Load()

    Set m_rsFee = New ADODB.Recordset
    m_rsFee.CursorLocation = adUseClient
    
    g_sSQL = "SELECT * FROM tblFee ORDER BY tblFee.Desc"
    m_rsFee.Open g_sSQL, g_cn, adOpenForwardOnly, adLockReadOnly
        
    Call initFrame(fraBG, frmMain.imlPanel, g_eIcon.Fee)
    Icon = frmMain.Icon
    Call CenterForm(Me)
    
    Call ldFee
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_rsFee = Nothing
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    Call Numeric(txtAmount, KeyAscii)
End Sub

Private Sub txtAmount_LostFocus()
    txtAmount = Format(txtAmount, FMT_CUR)
End Sub
