VERSION 5.00
Begin VB.UserControl ucVertical3DLine 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   210
   ScaleHeight     =   1590
   ScaleWidth      =   210
   Begin VB.Line LineB 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      X1              =   120
      X2              =   120
      Y1              =   0
      Y2              =   1560
   End
   Begin VB.Line LineA 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1560
   End
End
Attribute VB_Name = "ucVertical3DLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Resize()
    UserControl.Width = 90
    LineA.Y2 = UserControl.Height
    LineB.Y2 = UserControl.Height
    LineB.X1 = 15
    LineB.X2 = 15
End Sub
