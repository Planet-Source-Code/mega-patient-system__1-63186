VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5715
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
   ScaleHeight     =   1770
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBG 
      BorderStyle     =   0  'None
      Height          =   1845
      Left            =   30
      ScaleHeight     =   1845
      ScaleWidth      =   5685
      TabIndex        =   5
      Top             =   30
      Width           =   5685
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   360
         Left            =   3210
         TabIndex        =   3
         Top             =   1230
         WhatsThisHelpID =   3056
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   360
         Left            =   4425
         TabIndex        =   4
         Top             =   1230
         WhatsThisHelpID =   3055
         Width           =   1140
      End
      Begin prjPatientSystem.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   90
         TabIndex        =   8
         Top             =   1050
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   53
      End
      Begin VB.OptionButton optType 
         Caption         =   "Printer"
         Height          =   375
         Index           =   1
         Left            =   3030
         TabIndex        =   2
         Top             =   660
         Width           =   945
      End
      Begin VB.OptionButton optType 
         Caption         =   "Sreen"
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   660
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   270
         Width           =   4485
      End
      Begin VB.Label Label2 
         Caption         =   "Print to"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   690
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Title"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdOk_Click()
    Dim sTitle  As String
    
    Select Case Tag
        Case g_eMenu.Patient    '-- Patient
            
            With rptPatient
                .Icon = frmMain.Icon
                sTitle = Trim$(txtTitle)
                .Caption = IIf(sTitle = vbNullString, App.Title, sTitle)
                Set .DataSource = g_rsPatient
                If Len(sTitle) = 0 Then
                    .Sections("Section4").Visible = False '-- Hide Header.
                Else
                    .Sections("Section4").Controls("lblTitle").Caption = sTitle
                    .Sections("Section4").Controls("lblTitle_Shadow").Caption = sTitle
                End If
                Me.Hide
                If optType(0).Value = True Then '-- Screen.
                    .Show vbModal, frmMain
                Else '-- Printer
                    .PrintReport True
                End If
                Call Unload(Me)
            End With
            
        Case g_eMenu.Fee    '-- Fee.
        
            With rptFee
                .Icon = frmMain.Icon
                sTitle = Trim$(txtTitle)
                .Caption = IIf(sTitle = vbNullString, App.Title, sTitle)
                Set .DataSource = g_rsFee
                If Len(sTitle) = 0 Then
                    .Sections("Section4").Visible = False '-- Hide Header.
                Else
                    .Sections("Section4").Controls("lblTitle").Caption = sTitle
                    .Sections("Section4").Controls("lblTitle_Shadow").Caption = sTitle
                End If
                Me.Hide
                If optType(0).Value = True Then '-- Screen.
                    .Show vbModal, frmMain
                Else '-- Printer
                    .PrintReport True
                End If
                Call Unload(Me)
            End With
            
        Case g_eMenu.User   '-- User.
        
            With rptUser
                .Icon = frmMain.Icon
                sTitle = Trim$(txtTitle)
                .Caption = IIf(sTitle = vbNullString, App.Title, sTitle)
                Set .DataSource = g_rsUser
                If Len(sTitle) = 0 Then
                    .Sections("Section4").Visible = False '-- Hide Header.
                Else
                    .Sections("Section4").Controls("lblTitle").Caption = sTitle
                    .Sections("Section4").Controls("lblTitle_Shadow").Caption = sTitle
                End If
                Me.Hide
                If optType(0).Value = True Then '-- Screen.
                    .Show vbModal, frmMain
                Else '-- Printer
                    .PrintReport True
                End If
                Call Unload(Me)
            End With
        
        Case g_eMenu.Visitation
        
            With rptVisitation
                .Icon = frmMain.Icon
                sTitle = Trim$(txtTitle)
                .Caption = IIf(sTitle = vbNullString, App.Title, sTitle)
                Set .DataSource = g_rsVisitation
                If Len(sTitle) = 0 Then
                    .Sections("Section4").Visible = False '-- Hide Header.
                Else
                    .Sections("Section4").Controls("lblTitle").Caption = sTitle
                    .Sections("Section4").Controls("lblTitle_Shadow").Caption = sTitle
                End If
                Me.Hide
                If optType(0).Value = True Then '-- Screen.
                    .Show vbModal, frmMain
                Else '-- Printer
                    .PrintReport True
                End If
                Call Unload(Me)
            End With
            
        Case g_eMenu.Today
        
            With rptToday
                .Icon = frmMain.Icon
                sTitle = Trim$(txtTitle)
                .Caption = IIf(sTitle = vbNullString, App.Title, sTitle)
                Set .DataSource = g_rsToday
                If Len(sTitle) = 0 Then
                    .Sections("Section4").Visible = False '-- Hide Header.
                Else
                    .Sections("Section4").Controls("lblTitle").Caption = sTitle
                    .Sections("Section4").Controls("lblTitle_Shadow").Caption = sTitle
                End If
                Me.Hide
                If optType(0).Value = True Then '-- Screen.
                    .Show vbModal, frmMain
                Else '-- Printer
                    .PrintReport True
                End If
                Call Unload(Me)
            End With
                
    End Select
End Sub

Private Sub Form_Load()

    Call CenterForm(Me)
    Icon = frmMain.Icon
End Sub


