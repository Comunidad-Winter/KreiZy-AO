VERSION 5.00
Begin VB.Form Guia 
   Caption         =   "Guía KreiZy AO 2.0"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Vip"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Castillo"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "Form1.frx":0000
      Top             =   600
      Width           =   6135
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Donaciones"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Recanjeo"
      Height          =   375
      Left            =   6480
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Entrenamiento"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clanes"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Honor"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Guia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = "¿Como se consigue el honor? -Matando usuarios (Al matar a un usuario superior a vos te da 20 puntos de honor, si es igual que vos, te da 15 de honor y si es menor a vos te da 10 de honor, esto se calcula todo con el ranking de honores                         -Si te matan perdes 10 puntos de honor                                                                        -Si jugas un reto apostas 50 puntos de honor.                                                                -La quest automatica te da 100 puntos de honor, el torneo te da 150 puntos de honor, y el deathmatch te da 100 puntos de honor."
End Sub

Private Sub Command2_Click()
Text1.Text = "Requisitos para fundarclan:                                                                                    -Nivel 45                                                                                                                    -750 puntos de honor                                                                                                   -1 de canjeo."
End Sub

Private Sub Command3_Click()
Text1.Text = "Al apretar el botón [EDITAR PJ], te sube 2 niveles y te pone a todos los skills en 100 además de darte por cada apretada de botón 1kk."
End Sub

Private Sub Command4_Click()
Text1.Text = "Al recanjear un item, te da varios puntos menos que el valor inicial del objeto rencajeado, esto se utliza con la tecla F8."
End Sub

Private Sub command5_Click()
Text1.Text = "Las donaciones serán activas 1 o 2 semanas después de la apertura del AO"
End Sub

Private Sub Command6_Click()
Text1.Text = "Al castillo de clanes se entra por la sala de teleports que te lleva a un laverinto y desde ahi hay un tp para ir al castillo del rey. Cada 30 min el servidor entrega automaticamente 1 punto de canjeo a cada uno de los integrantes del clan que tiene conquistado el castillo."
End Sub

Private Sub command7_Click()
Text1.Text = "El sistema de vip te da: 10 de vida, 40 de mana y túnica de vip.  Los requisitos son: 3k500 de honor y 20 de canje."
End Sub

Private Sub Form_Load()
Text1.Text = "Al apretar el botón [EDITAR PJ], te sube 2 niveles y te pone a todos los skills en 100 además de darte por cada apretada de botón 1kk."
End Sub

