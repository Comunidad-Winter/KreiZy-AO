VERSION 5.00
Begin VB.Form frmTelep 
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   Picture         =   "frmTelep.frx":0000
   ScaleHeight     =   4755
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image16 
      Height          =   255
      Left            =   2400
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Image Image15 
      Height          =   255
      Left            =   0
      Top             =   600
      Width           =   1095
   End
   Begin VB.Image Image14 
      Height          =   255
      Left            =   1560
      Top             =   600
      Width           =   1095
   End
   Begin VB.Image Image13 
      Height          =   255
      Left            =   3120
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Image Image12 
      Height          =   255
      Left            =   1560
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Image Image11 
      Height          =   255
      Left            =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Image Image10 
      Height          =   375
      Left            =   240
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Image Image9 
      Height          =   375
      Left            =   360
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Image Image8 
      Height          =   255
      Left            =   3000
      Top             =   600
      Width           =   1215
   End
   Begin VB.Image Image7 
      Height          =   375
      Left            =   2520
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   2400
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   2400
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   240
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   120
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3720
      Top             =   0
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   120
      Top             =   40879
      Width           =   855
   End
End
Attribute VB_Name = "frmTelep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image14_Click()
Call SendData("#[")
Unload Me
End Sub
 
Private Sub Image15_Click()
Call SendData("#%")
Unload Me
End Sub
 
Private Sub Image6_Click()
Call SendData("#¦")
End Sub

Private Sub Image8_Click()
Call SendData("#=")
Unload Me
End Sub
 
Private Sub Image11_Click()
Call SendData("#-")
Unload Me
End Sub
 
Private Sub Image12_Click()
Call SendData("#+")
Unload Me
End Sub
 
Private Sub Image13_Click()
Call SendData("#\")
Unload Me
End Sub

Private Sub Image3_Click()
Call SendData("#¿")
Unload Me
End Sub
 
Private Sub Image4_Click()
Call SendData("#_")
Unload Me
End Sub
 
Private Sub Image5_Click()
Call SendData("#ª")
Unload Me
End Sub
 
Private Sub Image10_Click()
Call SendData("#;")
Unload Me
End Sub
 
Private Sub Image7_Click()
Call SendData("#{")
Unload Me
End Sub
 
Private Sub Image16_Click()
Call SendData("#^")
Unload Me
End Sub

Private Sub Image2_Click()
Me.Visible = False
End Sub

Private Sub Image9_Click()
Call SendData("#|")
Unload Me
End Sub
