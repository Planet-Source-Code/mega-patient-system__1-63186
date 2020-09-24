VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F689EFD5-B674-4C69-843B-03C8FC1D7583}#1.0#0"; "Phantom.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12105
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin prjPatientSystem.ucHorizontal3DLine ucToolbar 
      Height          =   30
      Left            =   -30
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   360
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   53
   End
   Begin prjPatientSystem.ucVertical3DLine ucDate_Sep 
      Height          =   255
      Left            =   11340
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin MSComctlLib.TreeView tvwReport 
      Height          =   1455
      Left            =   4080
      TabIndex        =   26
      Top             =   3270
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2566
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjPatientSystem.ucStatusbar ucStatus 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      Top             =   7245
      Width           =   12105
      _ExtentX        =   21352
      _ExtentY        =   609
   End
   Begin MSComctlLib.ImageList imlPanel 
      Left            =   4980
      Top             =   6150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3BBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3ED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C28
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":73C4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine8 
      Height          =   255
      Left            =   9390
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine7 
      Height          =   255
      Left            =   8220
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin PhantomPanel.PanelControl pMenu 
      Height          =   3225
      Left            =   180
      TabIndex        =   17
      Top             =   1230
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   5689
      PanelIconPictureSize=   32
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PanelStyle      =   1
      SelectedItemColor=   128
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine6 
      Height          =   255
      Left            =   7020
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine5 
      Height          =   255
      Left            =   5850
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine4 
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine3 
      Height          =   255
      Left            =   3510
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine2 
      Height          =   255
      Left            =   2340
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjPatientSystem.ucVertical3DLine ucVertical3DLine1 
      Height          =   255
      Left            =   1170
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjPatientSystem.lvButtons_H cmdTools 
      Height          =   375
      Left            =   7050
      TabIndex        =   7
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Tools"
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
   Begin prjPatientSystem.lvButtons_H cmdRefresh 
      Height          =   375
      Left            =   3510
      TabIndex        =   4
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Refresh"
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
   Begin prjPatientSystem.lvButtons_H cmdDelete 
      Height          =   375
      Left            =   2340
      TabIndex        =   3
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Delete"
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
   Begin prjPatientSystem.lvButtons_H cmdEdit 
      Height          =   375
      Left            =   1170
      TabIndex        =   2
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Edit"
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
   Begin prjPatientSystem.lvButtons_H cmdAdd 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Add"
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
   Begin VB.PictureBox picData 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   4620
      ScaleHeight     =   2415
      ScaleWidth      =   4875
      TabIndex        =   10
      Top             =   1860
      Width           =   4875
      Begin MSDataGridLib.DataGrid dgData 
         Height          =   1725
         Left            =   -30
         TabIndex        =   0
         Top             =   0
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   3043
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BorderStyle     =   0
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
         ColumnCount     =   2
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin prjPatientSystem.lvButtons_H cmdNavigate 
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "First Record"
         Top             =   2100
         Width           =   375
         _ExtentX        =   661
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
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin prjPatientSystem.lvButtons_H cmdNavigate 
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Previous Record"
         Top             =   2100
         Width           =   375
         _ExtentX        =   661
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
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin prjPatientSystem.lvButtons_H cmdNavigate 
         Height          =   315
         Index           =   2
         Left            =   720
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Next Record"
         Top             =   2100
         Width           =   375
         _ExtentX        =   661
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
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin prjPatientSystem.lvButtons_H cmdNavigate 
         Height          =   315
         Index           =   3
         Left            =   1080
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Last Record"
         Top             =   2100
         Width           =   375
         _ExtentX        =   661
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
         Focus           =   0   'False
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
      Begin prjPatientSystem.lvButtons_H cmdFooter 
         Height          =   315
         Left            =   0
         TabIndex        =   21
         Top             =   2070
         Width           =   4845
         _ExtentX        =   8546
         _ExtentY        =   556
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
   Begin MSComctlLib.ImageList imlTrans 
      Left            =   3690
      Top             =   6060
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9F16
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A2B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A64A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A9E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AD7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B118
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B4B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FDEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10186
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15A08
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1641A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16734
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17268
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17802
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin prjPatientSystem.lvButtons_H cmdHelp 
      Height          =   375
      Left            =   8220
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Help"
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
   Begin prjPatientSystem.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   5850
      TabIndex        =   6
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Print"
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
   Begin prjPatientSystem.lvButtons_H cmdDate 
      Height          =   375
      Left            =   10470
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "System Date"
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Date"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   16711680
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjPatientSystem.lvButtons_H cmdSearch 
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Search"
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
   Begin prjPatientSystem.ucVertical3DLine ucUser_Sep 
      Height          =   255
      Left            =   9600
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjPatientSystem.lvButtons_H cmdUser 
      Height          =   375
      Left            =   9480
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   0
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
      Caption         =   "User"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   16711680
      cFHover         =   16711680
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjPatientSystem.lvButtons_H cmdToolbar 
      Height          =   375
      Left            =   9060
      TabIndex        =   8
      Top             =   0
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   661
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuTools_Item 
         Caption         =   "Backup"
         Index           =   0
      End
      Begin VB.Menu mnuTools_Item 
         Caption         =   "Restore"
         Index           =   1
      End
      Begin VB.Menu mnuTools_Item 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuTools_Item 
         Caption         =   "Calculator"
         Index           =   3
      End
      Begin VB.Menu mnuTools_Item 
         Caption         =   "Notepad"
         Index           =   4
      End
      Begin VB.Menu mnuTools_Item 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuTools_Item 
         Caption         =   "Options"
         Index           =   6
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu mnuHelp_Item 
         Caption         =   "Function Keys"
         Index           =   0
      End
      Begin VB.Menu mnuHelp_Item 
         Caption         =   "User's Manual"
         Index           =   1
      End
      Begin VB.Menu mnuHelp_Item 
         Caption         =   "Contents"
         Index           =   2
      End
      Begin VB.Menu mnuHelp_Item 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuHelp_Item 
         Caption         =   "About"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'-- Local Variables.
Private m_rsCurrent     As ADODB.Recordset
Private m_sSort         As String
Private m_lFieldIdx     As Long
'

'//---------------------------------------------------------------------------------------
'//--Procedure : initPanel
'//--DateTime  : 7/12/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   : Initialize our Panel.
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub initPanel()

On Error GoTo errHandler

    Dim c As pTab

    Me.pMenu.Clear

    Me.pMenu.LockUpdate = True
    
    Set c = Me.pMenu.Panels.Add("List of patients.", imlPanel.ListImages(1).Picture, "Patient List", 1, picData, True)
    Set c = Me.pMenu.Panels.Add("List of fees.", imlPanel.ListImages(2).Picture, "Fee List", 2, picData)
    Set c = Me.pMenu.Panels.Add("List of users.", imlPanel.ListImages(3).Picture, "User List", 3, picData)
    Set c = Me.pMenu.Panels.Add("List of all Visitations occured.", imlPanel.ListImages(4).Picture, "Visitations", 4, picData)
    Set c = Me.pMenu.Panels.Add("Current date visitations.", imlPanel.ListImages(5).Picture, "Today's Visitations", 5, picData)
    Set c = Me.pMenu.Panels.Add("System reports.", imlPanel.ListImages(6).Picture, "Reports", 6, tvwReport)
    Me.pMenu.LockUpdate = False

errHandler:
    Set c = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : initStatusbar
'//--DateTime  : 7/18/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub initStatusbar()
    Dim tmpFont As StdFont
    
On Error GoTo errHandler
    
    With ucStatus
        
        '-- Initialize statusbar
        Call .Initialize(SizeGrip:=True, ToolTips:=True)
        
        .Font.Size = 8
        .Font.Name = "Tahoma"
        
        '-- Initialize icons list
'        Call .InitializeIconList
        
        '-- Add icons
'        Call .AddIcon(LoadResPicture("MAIL", vbResIcon))
'        Call .AddIcon(LoadResPicture("USER", vbResIcon))
'        Call .AddIcon(LoadResPicture("TIP", vbResIcon))
                             
        '-- Add panels
        Call .AddPanel(, , , [sbSpring], CPYRYT, , 0)
        Call .AddPanel(, , , [sbContents], "")
        Call .AddPanel(, , , [sbContents], "")
        
    End With

errHandler:
    Set tmpFont = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : initTreeview
'//--DateTime  : 9/2/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub initTreeview()

On Error GoTo errHandler

    With tvwReport
        .Nodes.Clear
        
        .Nodes.Add Text:="Patients Visition"
        .Nodes.Add Text:="Monthly Sales"
        .Nodes.Add Text:="Yearly Sales"
        
    End With

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmMain", "initTreeview", True)
    End If
End Sub


'//---------------------------------------------------------------------------------------
'//--Procedure : setToolbar
'//--DateTime  : 7/12/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub setToolbar()

On Error GoTo errHandler

    Select Case pMenu.SelectedItem.Key
        Case g_eMenu.Report
            cmdAdd.Enabled = False
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdRefresh.Enabled = False
            cmdSearch.Enabled = False
        Case Else
            cmdAdd.Enabled = True
            cmdEdit.Enabled = True
            cmdDelete.Enabled = True
            cmdRefresh.Enabled = True
            cmdSearch.Enabled = True
    End Select

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmMain", "setToolbar", True)
    End If
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : setPatient
'//--DateTime  : 7/22/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub setPatient()
    Dim bVal As Boolean
    
On Error GoTo errHandler

    '-- No Record.
    If g_rsPatient.EOF Then
        bVal = False
    Else
        bVal = True
    End If
    cmdEdit.Enabled = bVal
    cmdDelete.Enabled = bVal

    g_rsPatient.Requery

    Set m_rsCurrent = g_rsPatient
    
    With dgData
        Set .DataSource = g_rsPatient
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
    
    Call ShowRecordInfo(g_rsPatient, cmdFooter)
    
    '-- Clear Panel.
    With ucStatus
       .PanelText(2) = vbNullString
       .PanelText(3) = vbNullString
    End With
    
errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmMain", "setPatient", True)
    End If
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : setFee
'//--DateTime  : 7/22/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub setFee()
    Dim bVal As Boolean
    
On Error GoTo errHandler

    '-- No Record.
    If g_rsFee.EOF Then
        bVal = False
    Else
        bVal = True
    End If
    cmdEdit.Enabled = bVal
    cmdDelete.Enabled = bVal

    g_rsFee.Requery
    
    Set m_rsCurrent = g_rsFee
    
    With dgData
        Set .DataSource = g_rsFee
        .Columns(0).Visible = False             '-- ID.
        .Columns(1).Caption = "Description"     '-- Description.
        .Columns(2).Caption = "Amount"          '-- Amount
        .Columns(2).Alignment = dbgRight
        .Columns(3).Visible = False             '-- Date Created.
        .Columns(4).Visible = False             '-- Date Modified.
        .Columns(5).Visible = False
    End With

    Call ShowRecordInfo(g_rsFee, cmdFooter)

    '-- Clear Panel.
    With ucStatus
       .PanelText(2) = vbNullString
       .PanelText(3) = vbNullString
    End With

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmMain", "setFee", True)
    End If
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : setUser
'//--DateTime  : 7/22/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub setUser()
    Dim bVal As Boolean
    
On Error GoTo errHandler
    
    '-- No Record.
    If g_rsUser.EOF Then
        bVal = False
    Else
        bVal = True
    End If
    cmdEdit.Enabled = bVal
    cmdDelete.Enabled = bVal
    
    g_rsUser.Requery
    
    Set m_rsCurrent = g_rsUser
    
    With dgData
        Set .DataSource = g_rsUser
        .Columns(0).Visible = False             '-- ID.
        .Columns(1).Caption = "User ID"         '-- User ID.
        .Columns(2).Caption = "User Name"       '-- User Name.
        .Columns(3).Visible = False             '-- Password.
        .Columns(4).Visible = False             '-- Level
        .Columns(5).Visible = False             '-- Date Created.
        .Columns(6).Visible = False             '-- Date Modified.
    End With
    
    Call ShowRecordInfo(g_rsUser, cmdFooter)
    
    '-- Clear Panel.
    With ucStatus
       .PanelText(2) = vbNullString
       .PanelText(3) = vbNullString
    End With

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmMain", "setUser", True)
    End If
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : setVisitation
'//--DateTime  : 8/12/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub setVisitation()
    Dim bVal As Boolean
    
On Error GoTo errHandler

    '-- No Record.
    If g_rsVisitation.EOF Then
        bVal = False
    Else
        bVal = True
    End If
    cmdEdit.Enabled = bVal
    cmdDelete.Enabled = bVal
    
    g_rsVisitation.Requery

    Set m_rsCurrent = g_rsVisitation
        
    With dgData
        Set .DataSource = g_rsVisitation
        .Columns(0).Visible = False             '-- ID.
        .Columns(1).Visible = False             '-- PatientID.
        .Columns(2).Caption = "Name"            '-- Name.
        .Columns(3).Caption = "Address"         '-- Address.
        .Columns(4).Caption = "Sex"             '-- Sex.
        .Columns(5).Caption = "Civil Status"    '-- CivilStatus.
        .Columns(6).Caption = "Age"             '-- Age
        .Columns(7).Caption = "Contact No"      '-- ContactNo.
        .Columns(8).Caption = "Occupation"      '-- Occupation.
        .Columns(9).Caption = "Date"            '-- VisitationDate.
    End With
    
    Call ShowRecordInfo(g_rsVisitation, cmdFooter)
    
    '-- Clear Panel.
    With ucStatus
       .PanelText(2) = vbNullString
       .PanelText(3) = vbNullString
    End With

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmMain", "setVisitation", True)
    End If
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : setToday
'//--DateTime  : 8/15/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub setToday()
    Dim bVal                  As Boolean
    Dim rsTemp                As ADODB.Recordset
    Dim lPatient_Count        As Long
    Dim lPatient_GrandTotal   As Long
    
On Error GoTo errHandler

    '-- No Record.
    If g_rsVisitation.EOF Then
        bVal = False
    Else
        bVal = True
    End If
    cmdEdit.Enabled = bVal
    cmdDelete.Enabled = bVal

    g_rsToday.Requery

    Set m_rsCurrent = g_rsToday
    
    With dgData
        Set .DataSource = g_rsToday
        .Columns(0).Visible = False             '-- ID.
        .Columns(1).Visible = False             '-- PatientID.
        .Columns(2).Caption = "Name"            '-- Name.
        .Columns(3).Caption = "Address"         '-- Address.
        .Columns(4).Caption = "Sex"             '-- Sex.
        .Columns(5).Caption = "Civil Status"    '-- CivilStatus.
        .Columns(6).Caption = "Age"             '-- Age
        .Columns(7).Caption = "Contact No"      '-- ContactNo.
        .Columns(8).Caption = "Occupation"      '-- Occupation.
        .Columns(9).Caption = "Date"            '-- VisitationDate.
        .Columns(10).Caption = "Total Amount"   '-- TotalAmount.
    End With
    
    Call ShowRecordInfo(g_rsToday, cmdFooter)
    
    '-- Get Grand Total Amount.
    Set rsTemp = New ADODB.Recordset
    g_sSQL = "SELECT SUM(TotalAmount) AS AmountSum FROM tblVisitation WHERE VisitationDate = #" & Date & "#"
    rsTemp.Open g_sSQL, g_cn, adOpenForwardOnly, adLockReadOnly
    If Not rsTemp.EOF Then
        With ucStatus
            .PanelText(2) = "Grand Total"
            .PanelText(3) = Format(IIf(IsNull(rsTemp.Fields("AmountSum")), 0, rsTemp.Fields("AmountSum")), FMT_CUR)
        End With
    End If

errHandler:
    Set rsTemp = Nothing
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmMain", "setToday", True)
    End If
End Sub

'//---------------------------------------------------------------------------------------
'//--Procedure : dlVisitation
'//--DateTime  : 8/12/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub dlVisitation(lID As Long)
    Dim lTrans As Long
    
On Error GoTo errHandler

    g_cn.BeginTrans
    lTrans = 1
    
    g_cn.Execute "DELET * FROM tblVisitation WHERE ID=" & lID
    g_cn.Execute "DELET * FROM tblVisitation_Fee WHERE VisitationID=" & lID
    
    g_cn.CommitTrans
    lTrans = 0

errHandler:
    If Err.Number <> 0 Then
        If lTrans = 1 Then g_cn.RollbackTrans
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmMain", "dlVisitation", True)
    End If
End Sub


'//---------------------------------------------------------------------------------------
'//--Procedure : ShowRecordInfo
'//--DateTime  : 8/26/2005
'//--Author    : Silvellis N. Cawaling
'//--Purpose   :
'//--Notes     :
'//---------------------------------------------------------------------------------------
'
Private Sub ShowRecordInfo(rs As ADODB.Recordset, cmdLavolpe As lvButtons_H)

On Error GoTo errHandler

    Dim lPos    As Long
    Dim lCount  As Long
    
    lPos = rs.AbsolutePosition
    lCount = rs.RecordCount
    
    cmdLavolpe.Caption = IIf(lCount < 1, 0, lPos) & " of " & lCount

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmMain", "ShowRecordInfo", True)
    End If
End Sub


Private Sub cmdAdd_Click()
    Select Case pMenu.SelectedItem.Key
        Case g_eMenu.Patient
            With frmPatient
                .g_lAddState = 1    '-- Add State.
                .g_lOldID = 0
                .g_sOldName = vbNullString
                .Caption = .Caption & CPT_ADD
                .Show vbModal
            End With
        Case g_eMenu.Fee
            With frmFee
                .g_lAddState = 1    '-- Add State.
                .g_lOldID = 0
                .g_sOldDesc = vbNullString
                .Caption = .Caption & CPT_ADD
                .Show vbModal
            End With
        Case g_eMenu.User
            With frmUser
                .g_lAddState = 1    '-- Add State.
                .g_lOldID = 0
                .g_sOldUserID = vbNullString
                .Caption = .Caption & CPT_ADD
                .Show vbModal
            End With
        Case g_eMenu.Visitation, g_eMenu.Today
            With frmVisitation
                .g_lAddState = 1    '-- Add State.
                .g_lOldID = 0
                .Caption = .Caption & CPT_ADD
                .Show vbModal
            End With
    End Select
End Sub

Private Sub cmdDate_Click()
    frmDate.Show vbModal, Me
End Sub

Private Sub cmdDelete_Click()
    Dim lResult As Long
       
    Select Case pMenu.SelectedItem.Key
        Case g_eMenu.Patient
            lResult = MsgBox("Are you sure you want to delete?", vbQuestion + vbYesNo)
            If lResult = vbYes Then
                '(*) Check for delete constraints.
                g_rsPatient.Delete
            End If
        Case g_eMenu.Fee
            lResult = MsgBox("Are you sure you want to delete?", vbQuestion + vbYesNo)
            If lResult = vbYes Then
                '(*) Check for delete constraints.
                g_rsFee.Delete
            End If
        Case g_eMenu.User
            '-- Check user level.
            If g_tUser.Level = 0 Then
                MsgBox "You do not have enough privilage to delete user info.", vbExclamation
                Exit Sub
            Else
                lResult = MsgBox("Are you sure you want to delete?", vbQuestion + vbYesNo)
                If lResult = vbYes Then
                    '(*) Check for delete constraints.
                    g_rsUser.Delete
                End If
            End If
        Case g_eMenu.Visitation
            lResult = MsgBox("Are you sure you want to delete?", vbQuestion + vbYesNo)
            If lResult = vbYes Then
                '(*) Check for delete constraints.
                Call dlVisitation(g_rsVisitation.Fields("ID"))
            End If
        Case g_eMenu.Today
            lResult = MsgBox("Are you sure you want to delete?", vbQuestion + vbYesNo)
            If lResult = vbYes Then
                '(*) Check for delete constraints.
                Call dlVisitation(g_rsToday.Fields("ID"))
            End If
     End Select
    
End Sub

Private Sub cmdEdit_Click()
    Select Case pMenu.SelectedItem.Key
        Case g_eMenu.Patient
            If g_rsPatient.EOF Then Exit Sub
            With frmPatient
                .g_lAddState = 0 '-- Edit State.
                .g_lOldID = g_rsPatient.Fields("ID")
                .g_sOldName = g_rsPatient.Fields("Name")
                .init
                .txtName = .g_sOldName
                .txtAddress = g_rsPatient.Fields("Address")
                If g_rsPatient.Fields("Sex") <> vbNullString Then
                    .cboSex = g_rsPatient.Fields("Sex")
                End If
                If g_rsPatient.Fields("CivilStatus") <> vbNullString Then
                    .cboCivilStatus = g_rsPatient.Fields("CivilStatus")
                End If
                If g_rsPatient.Fields("Birthday") <> "1/1/1900" Then
                    .txtBirthday = g_rsPatient.Fields("Birthday")
                End If
                .txtAge = IIf(g_rsPatient.Fields("Age") = 0, "", g_rsPatient.Fields("Age"))
                .txtContactNo = IIf(IsNull(g_rsPatient.Fields("ContactNo")), vbNullString, g_rsPatient.Fields("ContactNo"))
                .txtEMail = IIf(IsNull(g_rsPatient.Fields("EMail")), vbNullString, g_rsPatient.Fields("EMail"))
                .txtOccupation = IIf(IsNull(g_rsPatient.Fields("Occupation")), vbNullString, g_rsPatient.Fields("Occupation"))
                .txtUdf1 = IIf(IsNull(g_rsPatient.Fields("Udf1")), vbNullString, g_rsPatient.Fields("Udf1"))
                .txtUdf2 = IIf(IsNull(g_rsPatient.Fields("Udf2")), vbNullString, g_rsPatient.Fields("Udf2"))
                .txtUdf3 = IIf(IsNull(g_rsPatient.Fields("Udf3")), vbNullString, g_rsPatient.Fields("Udf3"))
                .txtComments = IIf(IsNull(g_rsPatient.Fields("Comments")), vbNullString, g_rsPatient.Fields("Comments"))
                .Caption = .Caption & CPT_EDIT
                .Show vbModal
                '-- Reset the patient grid.
                Call setPatient
            End With
        Case g_eMenu.Fee
            If g_rsFee.EOF Then Exit Sub
            With frmFee
                .g_lAddState = 0 '-- Edit State.
                .g_lOldID = g_rsFee.Fields("ID")
                .g_sOldDesc = g_rsFee.Fields("Desc")
                .txtDesc = .g_sOldDesc
                .txtAmount = g_rsFee.Fields("Amount")
                .Caption = .Caption & CPT_EDIT
                .Show vbModal
                '-- Reset the fee grid.
                Call setFee
            End With
        Case g_eMenu.User
            If g_rsUser.EOF Then Exit Sub
            With frmUser
                .g_lAddState = 0 '-- Edit State.
                If g_tUser.Level = 0 Then   '-- Personnel
                    .g_lOldID = g_tUser.ID
                    .g_sOldUserID = g_tUser.UserID
                    .txtUserID = g_tUser.UserID
                    .txtName = g_tUser.Name
                    .txtPassword = g_tUser.Password
                    .cboLevel.ListIndex = g_tUser.Level
                    .cboLevel.Enabled = False
                Else
                    .g_lOldID = g_rsUser.Fields("ID")
                    .g_sOldUserID = g_rsUser.Fields("UserID")
                    .txtUserID = .g_sOldUserID
                    .txtName = g_rsUser.Fields("Name")
                    .txtPassword = g_rsUser.Fields("Password")
                    .cboLevel.ListIndex = g_rsUser.Fields("Level")
                End If
                .Caption = .Caption & CPT_EDIT
                .Show vbModal
                '-- Reset the user grid.
                Call setUser
            End With
        Case g_eMenu.Visitation
            If g_rsVisitation.EOF Then Exit Sub
            With frmVisitation
                .g_lAddState = 0 '-- Edit State.
                .g_lOldID = g_rsVisitation.Fields("ID")
                .Caption = .Caption & CPT_EDIT
                .Show vbModal
            End With
        Case g_eMenu.Today
            If g_rsToday.EOF Then Exit Sub
            With frmVisitation
                .g_lAddState = 0 '-- Edit State.
                .g_lOldID = g_rsToday.Fields("ID")
                .Caption = .Caption & CPT_EDIT
                .Show vbModal
            End With
    End Select
End Sub

Private Sub cmdHelp_Click()
    Call PopupMenu(mnuHelp, , cmdHelp.Left, cmdHelp.Top + cmdHelp.Height)
End Sub

Private Sub cmdNavigate_Click(Index As Integer)
    Call Navigate(Index, m_rsCurrent)
    Call ShowRecordInfo(m_rsCurrent, cmdFooter)
End Sub

Private Sub cmdPrint_Click()
    With frmPrint
        .Tag = pMenu.SelectedItem.Key
        Select Case pMenu.SelectedItem.Key
        Case g_eMenu.Patient
            .txtTitle = "List of Patients"
        Case g_eMenu.Fee
            .txtTitle = "List of Fees"
        Case g_eMenu.User
            .txtTitle = "List of Users"
        Case g_eMenu.Visitation
            .txtTitle = "List of Visitations"
        Case g_eMenu.Today
            .txtTitle = "List of Todays Visitations"
        End Select
        .Show vbModal
    End With
End Sub

Private Sub cmdRefresh_Click()
    Select Case pMenu.SelectedItem.Key
        Case g_eMenu.Patient
            g_rsPatient.Filter = adFilterNone
            g_rsPatient.Requery
            Call setPatient
        Case g_eMenu.Fee
            g_rsFee.Filter = adFilterNone
            g_rsFee.Requery
            Call setFee
        Case g_eMenu.User
            g_rsUser.Filter = adFilterNone
            g_rsUser.Requery
            Call setUser
        Case g_eMenu.Visitation
            g_rsVisitation.Filter = adFilterNone
            g_rsVisitation.Requery
            Call setVisitation
        Case g_eMenu.Today
            g_rsToday.Filter = adFilterNone
            g_rsToday.Requery
            Call setToday
    End Select
End Sub

Private Sub cmdSearch_Click()
    Dim i   As Long
    
    With frmSearch
        .g_lCancel = 0
        
        Select Case pMenu.SelectedItem.Key
            Case g_eMenu.Patient
                Set .g_rs = g_rsPatient
            Case g_eMenu.Fee
                Set .g_rs = g_rsFee
            Case g_eMenu.User
                Set .g_rs = g_rsUser
            Case g_eMenu.Visitation
                Set .g_rs = g_rsVisitation
            Case g_eMenu.Today
                Set .g_rs = g_rsToday
        End Select
        
        For i = 0 To dgData.Columns.Count - 1
            If dgData.Columns(i).Visible Then
                .cboField.AddItem dgData.Columns(i).Caption
                .cboField.ItemData(.cboField.NewIndex) = i
            End If
        Next
        
        .Show vbModal
        
        '-- Check if user canceled the search.
        If .g_lCancel = 0 Then
            Select Case pMenu.SelectedItem.Key
                Case g_eMenu.Patient
                    Call setPatient
                Case g_eMenu.Fee
                    Call setFee
                Case g_eMenu.User
                    Call setUser
                Case g_eMenu.Visitation
                    Call setVisitation
                Case g_eMenu.Today
                    Call setToday
            End Select
        End If
        
    End With
End Sub

Private Sub cmdTools_Click()
    Call PopupMenu(mnuTools, , cmdTools.Left, cmdTools.Top + cmdTools.Height)
End Sub

Private Sub cmdUser_Click()
    If MsgBox("Are you sure you want to Log-Off?", vbQuestion + vbYesNo) = vbYes Then
        g_lLogOff = 1   '-- Set LogOff flag.
        Call Unload(Me)
    End If
End Sub

Private Sub dgData_DblClick()
    Call cmdEdit_Click
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
    m_rsCurrent.Sort = dgData.Columns(ColIndex).DataField & " " & m_sSort
End Sub

Private Sub dgData_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Then Call ShowRecordInfo(m_rsCurrent, cmdFooter)
End Sub

Private Sub dgData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call ShowRecordInfo(m_rsCurrent, cmdFooter)
End Sub

Public Sub Form_Load()
    Dim lidx   As Long
    
On Error GoTo errHandler

    Caption = g_sAppName
     
    '-- Initialize.
    Call initPanel
    Call initStatusbar
    Call initTreeview
    Call initToolbar(cmdToolbar, imlTrans)
    Call initToolbar(cmdAdd, imlTrans, , g_eIcon.Add)
    Call initToolbar(cmdEdit, imlTrans, , g_eIcon.Edit)
    Call initToolbar(cmdDelete, imlTrans, , g_eIcon.Delete)
    Call initToolbar(cmdRefresh, imlTrans, , g_eIcon.Refresh)
    Call initToolbar(cmdSearch, imlTrans, , g_eIcon.Search)
    Call initToolbar(cmdPrint, imlTrans, , g_eIcon.Print)
    Call initToolbar(cmdTools, imlTrans, , g_eIcon.Options)
'    Call initToolbar(cmdAbout, imlTrans)
    Call initToolbar(cmdUser, imlTrans)
    Call initToolbar(cmdDate, imlTrans)
    Call initToolbar(cmdHelp, imlTrans, , g_eIcon.Help)
    Call initToolbar(cmdFooter, imlTrans, 1)
    Call initToolbar(cmdNavigate(0), imlTrans, 1, g_eIcon.BOF, lv_Fill_Stretch)
    Call initToolbar(cmdNavigate(1), imlTrans, 1, g_eIcon.Previous, lv_Fill_Stretch)
    Call initToolbar(cmdNavigate(2), imlTrans, 1, g_eIcon.Next, lv_Fill_Stretch)
    Call initToolbar(cmdNavigate(3), imlTrans, 1, g_eIcon.EOF, lv_Fill_Stretch)
    
    '-- Recordsets.
    Set g_rsPatient = New ADODB.Recordset
    Set g_rsFee = New ADODB.Recordset
    Set g_rsUser = New ADODB.Recordset
    Set g_rsVisitation = New ADODB.Recordset
    Set g_rsToday = New ADODB.Recordset
    
    g_rsPatient.CursorLocation = adUseClient
    g_rsFee.CursorLocation = adUseClient
    g_rsUser.CursorLocation = adUseClient
    g_rsVisitation.CursorLocation = adUseClient
    g_rsToday.CursorLocation = adUseClient
    
    '-- Patient.
    g_sSQL = "SELECT * FROM tblPatient"
    g_rsPatient.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
    
    '-- Fee.
    g_sSQL = "SELECT * FROM tblFee"
    g_rsFee.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
    
    '-- User.
    g_sSQL = "SELECT * FROM tblUser"
    g_rsUser.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
    
    '-- Visitations.
    g_sSQL = "SELECT * FROM qryVisitation"
    g_rsVisitation.Open g_sSQL, g_cn, adOpenForwardOnly, adLockReadOnly
    
    '-- Today.
    g_sSQL = "SELECT * FROM qryVisitation WHERE VisitationDate = #" & g_sSysDate & "#"
    g_rsToday.Open g_sSQL, g_cn, adOpenForwardOnly, adLockReadOnly
    
    '-- Default.
    Call setPatient

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmMain", "Form_Load", True)
        Call Unload(Me)
    End If
    
End Sub

Private Sub Form_Resize()
On Error GoTo errHandler

    ucStatus.SizeGrip = (WindowState <> vbMaximized)
    pMenu.Move 20, cmdToolbar.Height + 20, ScaleWidth - 30, ScaleHeight - (30 + cmdToolbar.Height + ucStatus.Height)
    cmdDate.Left = ScaleWidth - cmdDate.Width
    ucDate_Sep.Left = cmdDate.Left
    cmdUser.Left = cmdDate.Left - cmdUser.Width
    ucUser_Sep.Left = cmdUser.Left
    cmdToolbar.Width = ScaleWidth - cmdToolbar.Left
    ucToolbar.Width = ScaleWidth + 20
    
errHandler: Err.Clear: Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set g_rsPatient = Nothing
    Set g_rsFee = Nothing
    Set g_rsVisitation = Nothing
    Set g_rsToday = Nothing
    Set m_rsCurrent = Nothing
    Set g_cn = Nothing
    
    If g_lLogOff = 1 Then
        frmLogin.Show
    Else
        If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub mnuPrint_Click()
    
End Sub

Private Sub mnuHelp_Item_Click(Index As Integer)
    Select Case Index
        Case 0
        Case 1
        Case 2
        Case 3
        Case 4  '-- About.
            frmAbout.Show vbModal
    End Select
End Sub

Private Sub mnuTools_Item_Click(Index As Integer)
    Select Case Index
        Case 0  '-- Backup.
            frmBackup.Show vbModal
        Case 1  '-- Restore.
            frmRestore.Show vbModal
        Case 3  '-- Calculator.
            Call OpenThisFile("Calc", "1", "", hWnd)
        Case 4  '-- Notepad.
            Call OpenThisFile("Notepad", "1", "", hWnd)
        Case 6  '-- Options.
            frmOptions.Show vbModal
        Case Else
    End Select
End Sub

Private Sub picData_Resize()
On Error GoTo errHandler

    dgData.Move 0, 0, picData.ScaleWidth, picData.ScaleHeight - cmdFooter.Height
    cmdFooter.Move 0, dgData.Height, picData.ScaleWidth, cmdFooter.Height
    cmdNavigate(0).Top = cmdFooter.Top
    cmdNavigate(1).Top = cmdFooter.Top
    cmdNavigate(2).Top = cmdFooter.Top
    cmdNavigate(3).Top = cmdFooter.Top

errHandler: Err.Clear: Resume Next
End Sub

Private Sub pMenu_PanelSelected(oPanel As PhantomPanel.pTab)
    Select Case oPanel.Key
        Case g_eMenu.Patient
            Call setPatient
        Case g_eMenu.Fee
            Call setFee
        Case g_eMenu.User
            Call setUser
        Case g_eMenu.Visitation
            Call setVisitation
        Case g_eMenu.Today
            Call setToday
        Case g_eMenu.Report
    End Select
    Call setToolbar
End Sub

