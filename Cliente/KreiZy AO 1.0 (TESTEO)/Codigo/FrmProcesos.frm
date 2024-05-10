VERSION 5.00
Begin VB.Form FrmProcesos 
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   3960
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "FrmProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
Unload Me
End Sub

