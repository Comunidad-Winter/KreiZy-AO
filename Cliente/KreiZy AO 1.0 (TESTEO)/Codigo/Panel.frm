VERSION 5.00
Begin VB.Form Panel 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Desactivar"
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Activar"
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   2370
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3960
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Activar"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Intervalo"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Auto lanzar"
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Intervalo"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Minima mana al empezar el poteo"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label label1 
      Caption         =   "Minima vida al empezar el poteo"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "Panel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()

If Text1.Text = "" Then
Label3.Caption = "La minima vida al empezar el poteo está mal"
Exit Sub
End If
If Text2.Text = "" Then
Label3.Caption = "La minima mana al empezar el poteo está mal"
Exit Sub
End If
If Text3.Text = "" Then
Label3.Caption = "El intervalo de poteo está mal"
Exit Sub
End If

If asd = True Then
frmMain.Timer4.Enabled = True
frmMain.Timer4.Interval = Text3.Text
command1.Caption = "Desactivar"
asd = False
UserMinVida = Text2.Text
UserMinMana = Text1.Text
Else
frmMain.Timer4.Enabled = False
frmMain.Timer4.Interval = Text3.Text
command1.Caption = "Activar"
asd = True
UserMinVida = Text2.Text
UserMinMana = Text1.Text
End If
End Sub

Private Sub Command2_Click()
If Text4.Text = "" Then
Label3.Caption = "El intervalo de lanzar está mal"
Exit Sub
End If



frmMain.Timer5.Enabled = True
frmMain.Timer5.Interval = Text4.Text
Command2.Caption = "Activar"

End Sub

Private Sub Command3_Click()
frmMain.Timer5.Enabled = False
frmMain.Timer5.Interval = Text4.Text
Command3.Caption = "Desactivar"
End Sub

