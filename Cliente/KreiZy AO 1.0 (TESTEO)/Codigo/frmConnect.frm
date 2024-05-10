VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmConectar 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Online"
   ClientHeight    =   9015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmConnect.frx":000C
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1440
      Top             =   960
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3840
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock estado 
      Left            =   1560
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtPass 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3840
      Width           =   3300
   End
   Begin VB.TextBox txtUser 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   4350
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2880
      Width           =   3300
   End
   Begin VB.Label Lblestad 
      Caption         =   "Label1"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Lblestado 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   1800
      Left            =   3840
      MouseIcon       =   "frmConnect.frx":3FC6B
      MousePointer    =   99  'Custom
      Top             =   360
      Width           =   4320
   End
   Begin VB.Image imgWeb 
      Height          =   660
      Left            =   7920
      MouseIcon       =   "frmConnect.frx":3FF75
      MousePointer    =   99  'Custom
      Top             =   7920
      Width           =   4080
   End
   Begin VB.Image imgGetPass 
      Height          =   375
      Left            =   4680
      MouseIcon       =   "frmConnect.frx":4027F
      MousePointer    =   99  'Custom
      Top             =   6120
      Width           =   2625
   End
   Begin VB.Image Image1 
      Height          =   435
      Index           =   0
      Left            =   4680
      MouseIcon       =   "frmConnect.frx":40589
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   2625
   End
   Begin VB.Image Image1 
      Height          =   420
      Index           =   1
      Left            =   4680
      MouseIcon       =   "frmConnect.frx":40893
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   435
      Index           =   2
      Left            =   4680
      MouseIcon       =   "frmConnect.frx":40B9D
      MousePointer    =   99  'Custom
      Top             =   6600
      Width           =   2625
   End
End
Attribute VB_Name = "frmConectar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FénixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'You can contact the original creator of Argentum Online at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar
Option Explicit
Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    Call PlayWaveDS(SND_CLICK)
            
    If frmPrincipal.Socket1.Connected Then frmPrincipal.Socket1.Disconnect
    
    If frmConectar.MousePointer = 11 Then
    frmConectar.MousePointer = 1
        Exit Sub
    End If
    
    
    UserName = txtUser.Text
    Dim aux As String
    aux = txtPass.Text
    UserPassword = MD5String(aux)
    If CheckUserData(False) = True Then
        frmPrincipal.Socket1.HostName = IPdelServidor
        frmPrincipal.Socket1.RemotePort = PuertoDelServidor
        
        EstadoLogin = Normal
        Me.MousePointer = 11
        frmPrincipal.Socket1.Connect
    End If
End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
    frmCargando.Show
    frmCargando.Refresh
    AddtoRichTextBox frmCargando.Status, "Cerrando ValZhatAO.", 255, 150, 50, 1, 0, 1
    
    Call SaveGameini
    frmConectar.MousePointer = 1
    frmPrincipal.MousePointer = 1
    prgRun = False
    
    AddtoRichTextBox frmCargando.Status, "Liberando recursos..."
    frmCargando.Refresh
    LiberarObjetosDX
    AddtoRichTextBox frmCargando.Status, "Hecho", 255, 150, 50, 1, 0, 1
    AddtoRichTextBox frmCargando.Status, "¡¡Gracias por jugar ValZhatAO!!", 255, 150, 50, 1, 0, 1
    frmCargando.Refresh
    Call UnloadAllForms
End If

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyI And Shift = vbCtrlMask Then
    
    

    
    
    


    
    
    KeyCode = 0
    Exit Sub
End If

End Sub

Private Sub Form_Load()
    If estado.State <> sckClosed Then
estado.Close
End If
estado.Connect "190.210.25.119", 7790
    If estado.State <> sckClosed Then
Winsock1.Close
End If
Winsock1.Connect "10.0.0.112", 10200

  
    EngineRun = False
    
    
 Dim j
 For Each j In Image1()
    j.Tag = "0"
 Next

 IntervaloPaso = 0.19
 IntervaloUsar = 0.14
 Picture = LoadPicture(DirGraficos & "conectar.jpg")


 
 
 
 
 
 

End Sub

Private Sub Image1_Click(Index As Integer)

CurServer = 0

Call PlayWaveDS(SND_CLICK)

Select Case Index
    Case 0

        If Musica = 0 Then
            CurMidi = DirMidi & "7.mid"
            LoopMidi = 1
            Call CargarMIDI(CurMidi)
            Call Play_Midi
        End If

       
        EstadoLogin = dados
        frmPrincipal.Socket1.HostName = IPdelServidor
        frmPrincipal.Socket1.RemotePort = PuertoDelServidor
        Me.MousePointer = 11
        frmPrincipal.Socket1.Connect
        
    Case 1
        
        If frmPrincipal.Socket1.Connected Then frmPrincipal.Socket1.Disconnect
        
        If frmConectar.MousePointer = 11 Then
        frmConectar.MousePointer = 1
            Exit Sub
        End If
        
        
        
        UserName = txtUser.Text
        Dim aux As String
        aux = txtPass.Text
        UserPassword = MD5String(aux)
        If CheckUserData(False) = True Then
            frmPrincipal.Socket1.HostName = IPdelServidor
            frmPrincipal.Socket1.RemotePort = PuertoDelServidor
            
            EstadoLogin = Normal
            Me.MousePointer = 11
            frmPrincipal.Socket1.Connect
        End If
        
Case 2
If frmPrincipal.Socket1.Connected Then frmPrincipal.Socket1.Disconnect
     
If frmConectar.MousePointer = 11 Then
frmConectar.MousePointer = 1
Exit Sub
End If
     
frmPrincipal.Socket1.HostName = IPdelServidor
frmPrincipal.Socket1.RemotePort = PuertoDelServidor
EstadoLogin = BorrarPj
Me.MousePointer = 11
frmPrincipal.Socket1.Connect

End Select

End Sub
Private Sub Image2_Click()

MsgBox "Created By ValZhatAO Team." & vbCrLf & "Copyright © 2009. Todos los derechos reservados." & vbCrLf & vbCrLf & "Web: http://www.ValZhatAO.ucoz.com" & vbCrLf & vbCrLf & "¡Gracias por Jugar nuestro Argentum Online!" & vbCrLf & "Staff ValZhatAO.", vbInformation, "Proyecto ValZhatAO"

End Sub
Private Sub imgGetPass_Click()
     
If frmPrincipal.Socket1.Connected Then frmPrincipal.Socket1.Disconnect
     
If frmConectar.MousePointer = 11 Then
frmConectar.MousePointer = 1
Exit Sub
End If
     
frmPrincipal.Socket1.HostName = IPdelServidor
frmPrincipal.Socket1.RemotePort = PuertoDelServidor
EstadoLogin = RecuperarPass
Me.MousePointer = 11
frmPrincipal.Socket1.Connect
    
End Sub
Private Sub imgWeb_Click()

Call ShellExecute(Me.hWnd, "open", "http://www.ValZhatAO.ucoz.com", "", "", 1)

End Sub
Private Sub ESTADO_Connect()
Lblestado.ForeColor = vbGreen
Lblestado.Caption = "Online"
End Sub
Private Sub ESTADO_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Lblestado.ForeColor = vbRed
Lblestado.Caption = "Offline"
End Sub

Private Sub Timer1_Timer()
    If estado.State <> sckClosed Then
estado.Close
End If
estado.Connect "190.210.25.119", 7790
End Sub

Private Sub winsock1_Connect()
Lblestad.ForeColor = vbGreen
Lblestad.Caption = "Online"
End Sub
Private Sub winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Lblestad.ForeColor = vbRed
Lblestad.Caption = "Offline"
End Sub

