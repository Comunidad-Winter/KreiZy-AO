VERSION 5.00
Begin VB.Form frmCanjes 
   BackColor       =   &H00000000&
   Caption         =   "Sistema de Canje"
   ClientHeight    =   4680
   ClientLeft      =   540
   ClientTop       =   765
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7035
   Begin VB.CommandButton Command2 
      Caption         =   "Arcos y Flechas"
      Height          =   495
      Left            =   5280
      TabIndex        =   13
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Escudos 
      Caption         =   "Escudos y Coronas"
      Height          =   495
      Left            =   3600
      TabIndex        =   12
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Armaduras 
      Caption         =   "Armaduras y Túnicas"
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Espadas 
      Caption         =   "Espadas y Baculos"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Canjear"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   3360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   540
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000004&
      Height          =   2205
      ItemData        =   "frmCanjes.frx":0000
      Left            =   120
      List            =   "frmCanjes.frx":0002
      TabIndex        =   0
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label lblPermisos 
      Height          =   975
      Left            =   3360
      TabIndex        =   8
      Top             =   3600
      Width           =   3495
   End
   Begin VB.Label lblStat 
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label lblPrecio 
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label lblNombre 
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clases Permitidas"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   3120
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stats:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   2640
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Precio:"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   2160
      Width           =   555
   End
End
Attribute VB_Name = "frmCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Armaduras_Click()
List1.Clear
List1.AddItem "Armadura Thek"
List1.AddItem "Túnica Durlock"
List1.AddItem "Túnica Angelical"
List1.AddItem "Tunica de Rey (Altos)"
List1.AddItem "Pantalon violeta"
List1.AddItem "Pantalon rojo"
List1.AddItem "Pantalon azul"
List1.AddItem "Pantalon negro"

    
End Sub

Private Sub Command1_Click()

If List1.Text = "Tunica de Rey (Altos)" Then Call SendData("/CANJEO T1")
If List1.Text = "Sombrero Infernal" Then Call SendData("/CANJEO T2")
If List1.Text = "Báculo de Mago Oscuro" Then Call SendData("/CANJEO T3")
If List1.Text = "Poción Roja GRANDE" Then Call SendData("/CANJEO T4")
If List1.Text = "Poción Azul GRANDE" Then Call SendData("/CANJEO T5")
If List1.Text = "Espada de Neithan + 2" Then Call SendData("/CANJEO T6")
If List1.Text = "Corona" Then Call SendData("/CANJEO T7")
If List1.Text = "Espada Fantasmal" Then Call SendData("/CANJEO T8")
If List1.Text = "Casco de Legionario" Then Call SendData("/CANJEO T9")
If List1.Text = "Arco de las Sombras" Then Call SendData("/CANJEO T10")
If List1.Text = "Arco de la Luz" Then Call SendData("/CANJEO T11")
If List1.Text = "Arco largo engarzado" Then Call SendData("/CANJEO T12")
If List1.Text = "Daga + 5" Then Call SendData("/CANJEO T13")
If List1.Text = "Flecha +3" Then Call SendData("/CANJEO T14")
If List1.Text = "Escudo de León + 1" Then Call SendData("/CANJEO T15")
If List1.Text = "Escudo de la Alianza" Then Call SendData("/CANJEO T16")
If List1.Text = "Corona de Rey" Then Call SendData("/CANJEO T17")
If List1.Text = "Daga de Hielo" Then Call SendData("/CANJEO T18")
If List1.Text = "Escudo Dinal +1" Then Call SendData("/CANJEO T19")
If List1.Text = "Túnica Angelical" Then Call SendData("/CANJEO T20")
If List1.Text = "Espada Ardiente" Then Call SendData("/CANJEO T21")
If List1.Text = "Armadura Thek" Then Call SendData("/CANJEO T22")
If List1.Text = "Túnica Durlock" Then Call SendData("/CANJEO T23")
If List1.Text = "Pantalon violeta" Then Call SendData("/CANJEO T24")
If List1.Text = "Pantalon rojo" Then Call SendData("/CANJEO T25")
If List1.Text = "Pantalon azul" Then Call SendData("/CANJEO T26")
If List1.Text = "Pantalon negro" Then Call SendData("/CANJEO T27")

End Sub

Private Sub Drive1_Change()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command6_Click()

End Sub

Private Sub Command2_Click()
List1.Clear
List1.AddItem "Arco de las Sombras"
List1.AddItem "Arco de la Luz"
List1.AddItem "Arco largo engarzado"
List1.AddItem "Flecha +3"


End Sub

Private Sub Escudos_Click()
List1.Clear
List1.AddItem "Escudo de León + 1"
List1.AddItem "Escudo de la Alianza"
List1.AddItem "Corona de Rey"
List1.AddItem "Escudo Dinal +1"
List1.AddItem "Sombrero Infernal"
List1.AddItem "Corona"
List1.AddItem "Casco de Legionario"

End Sub

Private Sub Espadas_Click()
List1.Clear
List1.AddItem "Báculo de Mago Oscuro"
List1.AddItem "Espada Ardiente"
List1.AddItem "Daga de Hielo"
List1.AddItem "Daga + 5"
List1.AddItem "Espada de Neithan + 2"
List1.AddItem "Espada Fantasmal"

End Sub


Private Sub list1_Click()
If List1.Text = "Arco de las Sombras" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16116.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "25 Puntos de Canje"
    lblStat.Caption = "Min: 10 / Max: 15"
    lblPermisos.Caption = "Cazador"
    End If
If List1.Text = "Arco de la Luz" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16114.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "20 Puntos de Canje"
    lblStat.Caption = "Min: 10 / Max: 16"
    lblPermisos.Caption = "Arquero"
    End If
If List1.Text = "Arco largo engarzado" Then
    Picture1.Picture = LoadPicture(DirGraficos & "1004.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 14 / Max: 17"
    lblPermisos.Caption = "Arquero y Cazador"
    End If
If List1.Text = "Flecha +3" Then
    Picture1.Picture = LoadPicture(DirGraficos & "748.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "10 Puntos de Canje"
    lblStat.Caption = "Min: 0 / Max: 0"
    lblPermisos.Caption = "Arquero y Cazador"
    End If


If List1.Text = "Sombrero Infernal" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16032.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 13 / Max: 15"
    lblPermisos.Caption = "Mago"
    End If
   
If List1.Text = "Casco de Legionario" Then
    Picture1.Picture = LoadPicture(DirGraficos & "2019.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "25 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 42"
    lblPermisos.Caption = "Paladín, Guerrero y Arquero"
    End If
    
    If List1.Text = "Escudo de León + 1" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16060.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 9 / Max: 14"
    lblPermisos.Caption = "Clerigo"
    End If
If List1.Text = "Escudo de la Alianza" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16068.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "25 Puntos de Canje"
    lblStat.Caption = "Min: 8 / Max: 14"
    lblPermisos.Caption = "Paladín y Guerrero"
    End If
If List1.Text = "Corona de Rey" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16100.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "25 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Todas menos Guerrero"
    End If
    If List1.Text = "Escudo Dinal +1" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16064.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 10 / Max: 12"
    lblPermisos.Caption = "Bardo"
    End If

If List1.Text = "Poción Roja GRANDE" Then
    Picture1.Picture = LoadPicture(DirGraficos & "535.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "12 Puntos de Canje"
    lblStat.Caption = "Min: 31 / Max: 32"
    lblPermisos.Caption = "Todas las Clases"
    End If
If List1.Text = "Poción Azul GRANDE" Then
    Picture1.Picture = LoadPicture(DirGraficos & "534.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "13 Puntos de Canje"
    lblStat.Caption = "Min: 31 / Max: 32"
    lblPermisos.Caption = "Todas las Clases"
    End If



    If List1.Text = "Daga + 5" Then
    Picture1.Picture = LoadPicture(DirGraficos & "3537.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "25 Puntos de Canje"
    lblStat.Caption = "Min: 9 / Max: 11"
    lblPermisos.Caption = "Bardo"
    End If
If List1.Text = "Báculo de Mago Oscuro" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16030.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "35 Puntos de Canje"
    lblStat.Caption = "Min: 0 / Max: 0"
    lblPermisos.Caption = "Mago"
    End If
     If List1.Text = "Espada de Neithan + 2" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16070.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "40 Puntos de Canje"
    lblStat.Caption = "Min: 21 / Max: 25"
    lblPermisos.Caption = "Guerrero"
    End If
    
If List1.Text = "Espada Fantasmal" Then
    Picture1.Picture = LoadPicture(DirGraficos & "9630.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "40 Puntos de Canje"
    lblStat.Caption = "Min: 20 / Max: 22"
    lblPermisos.Caption = "Paladín y Guerrero"
    End If
    
    If List1.Text = "Daga de Hielo" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16118.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "25 Puntos de Canje"
    lblStat.Caption = "Min: 10 / Max: 12"
    lblPermisos.Caption = "Asesino"
    End If
    
 If List1.Text = "Espada Ardiente" Then
    Picture1.Picture = LoadPicture(DirGraficos & "9629.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "35 Puntos de Canje"
    lblStat.Caption = "Min: 18 / Max: 22"
    lblPermisos.Caption = "Clerigo"
    End If
  
If List1.Text = "Tunica de Rey (Altos)" Then
    Picture1.Picture = LoadPicture(DirGraficos & "685.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "15 Puntos de Canje"
    lblStat.Caption = "Min: 30 / Max: 35"
    lblPermisos.Caption = "Todas las Clases"
    End If
    If List1.Text = "Túnica Angelical" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16112.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "40 Puntos de Canje"
    lblStat.Caption = "Min: 35 / Max: 40"
    lblPermisos.Caption = "Mago"
    End If
  If List1.Text = "Corona" Then
    Picture1.Picture = LoadPicture(DirGraficos & "2023.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "20 Puntos de Canje"
    lblStat.Caption = "Min: 11 / Max: 13"
    lblPermisos.Caption = "Mago, clero, bardo y paladin"
End If
  If List1.Text = "Armadura Thek" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16048.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Clerigo y paladin"
End If
  If List1.Text = "Túnica Durlock" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16181.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "40 Puntos de Canje"
    lblStat.Caption = "Min: 36 / Max: 41"
    lblPermisos.Caption = "Bardo"
End If
  If List1.Text = "Pantalon violeta" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16108.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "20 Puntos de Canje"
    lblStat.Caption = "Min: 31 / Max: 36"
    lblPermisos.Caption = "Todas las clases"
End If
  If List1.Text = "Pantalon rojo" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16104.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "20 Puntos de Canje"
    lblStat.Caption = "Min: 31 / Max: 36"
    lblPermisos.Caption = "Todas las clases"
End If
  If List1.Text = "Pantalon azul" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16106.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "20 Puntos de Canje"
    lblStat.Caption = "Min: 31 / Max: 36"
    lblPermisos.Caption = "Todas las clases"
End If
  If List1.Text = "Pantalon negro" Then
    Picture1.Picture = LoadPicture(DirGraficos & "16110.bmp")
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "20 Puntos de Canje"
    lblStat.Caption = "Min: 31 / Max: 36"
    lblPermisos.Caption = "Todas las clases"
End If


End Sub
