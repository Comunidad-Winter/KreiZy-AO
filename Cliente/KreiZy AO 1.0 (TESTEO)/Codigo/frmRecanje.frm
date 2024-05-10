VERSION 5.00
Begin VB.Form frmRecanje 
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command24 
      Caption         =   "Pantalon NEGRO"
      Height          =   495
      Left            =   3960
      TabIndex        =   25
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Pantalon VIOLETA"
      Height          =   495
      Left            =   3960
      TabIndex        =   24
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Pantalon AZUL"
      Height          =   495
      Left            =   3960
      TabIndex        =   23
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Pantalon ROJO"
      Height          =   495
      Left            =   3960
      TabIndex        =   22
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Túnica Angeilcal"
      Height          =   495
      Left            =   2760
      TabIndex        =   20
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Escudo Dinal +1"
      Height          =   495
      Left            =   2760
      TabIndex        =   19
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Daga de Hielo"
      Height          =   495
      Left            =   3960
      TabIndex        =   18
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Corona de Rey"
      Height          =   495
      Left            =   1440
      TabIndex        =   17
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Escudo de la Alianza"
      Height          =   495
      Left            =   2760
      TabIndex        =   16
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Escudo de león +1"
      Height          =   495
      Left            =   2760
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Flecha +3"
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Daga +5"
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Arco largo engarzado"
      Height          =   495
      Left            =   1440
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Arco de la Luz"
      Height          =   495
      Left            =   1440
      TabIndex        =   11
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Arco de las Sombras"
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Casco de Legionario"
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Espada Fantasmal"
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Corona"
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Espada de Neithan +2"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Túnica Durlock"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Armadura Thek"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Báculo Oscuro"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sombrero Infernal"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Túnica de Rey"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Atención: El recanjeo te da menos canje que el valor común del objeto ( 3 menos)"
      Height          =   375
      Left            =   480
      TabIndex        =   21
      Top             =   4920
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "¿Qué desea recanjear?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmRecanje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call SendData("/RECANJEO T1")
End Sub

Private Sub Command10_Click()
Call SendData("/RECANJEO T10")
End Sub

Private Sub Command11_Click()
Call SendData("/RECANJEO T11")
End Sub

Private Sub Command12_Click()
Call SendData("/RECANJEO T12")
End Sub

Private Sub Command13_Click()
Call SendData("/RECANJEO T13")
End Sub

Private Sub Command14_Click()
Call SendData("/RECANJEO T14")
End Sub

Private Sub Command15_Click()
Call SendData("/RECANJEO T15")
End Sub

Private Sub Command16_Click()
Call SendData("/RECANJEO T16")
End Sub

Private Sub Command17_Click()
Call SendData("/RECANJEO T17")
End Sub

Private Sub Command18_Click()
Call SendData("/RECANJEO T18")
End Sub

Private Sub Command19_Click()
Call SendData("/RECANJEO T19")
End Sub

Private Sub Command2_Click()
Call SendData("/RECANJEO T2")
End Sub

Private Sub Command20_Click()
Call SendData("/RECANJEO T20")
End Sub

Private Sub Command21_Click()
Call SendData("/RECANJEO T21")
End Sub

Private Sub Command22_Click()
Call SendData("/RECANJEO T22")
End Sub

Private Sub Command23_Click()
Call SendData("/RECANJEO T23")
End Sub

Private Sub Command24_Click()
Call SendData("/RECANJEO T24")
End Sub

Private Sub Command3_Click()
Call SendData("/RECANJEO T3")
End Sub

Private Sub Command4_Click()
Call SendData("/RECANJEO T4")
End Sub

Private Sub command5_Click()
Call SendData("/RECANJEO T5")
End Sub

Private Sub Command6_Click()
Call SendData("/RECANJEO T6")
End Sub

Private Sub command7_Click()
Call SendData("/RECANJEO T7")
End Sub

Private Sub Command8_Click()
Call SendData("/RECANJEO T8")
End Sub

Private Sub Command9_Click()
Call SendData("/RECANJEO T9")
End Sub
