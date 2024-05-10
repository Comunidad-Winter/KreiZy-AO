VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servidor FlamiusAO  ~ Argentum Online ~"
   ClientHeight    =   3540
   ClientLeft      =   1950
   ClientTop       =   1695
   ClientWidth     =   6615
   ControlBox      =   0   'False
   FillColor       =   &H80000004&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000007&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3540
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.Timer Tlimpiar 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   3840
      Top             =   720
   End
   Begin VB.Timer retos2vs2 
      Interval        =   60000
      Left            =   3360
      Top             =   720
   End
   Begin VB.Timer TimerMeditar 
      Interval        =   400
      Left            =   2880
      Top             =   360
   End
   Begin VB.Data ADODB 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox i 
      Height          =   3180
      ItemData        =   "frmMain.frx":1042
      Left            =   5160
      List            =   "frmMain.frx":1049
      TabIndex        =   8
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "Mensaje BroadCast >>"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Usuarios:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   2760
         Top             =   720
      End
      Begin VB.Timer Timer10 
         Interval        =   1
         Left            =   1920
         Top             =   0
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   255
         Left            =   4560
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.Timer Timer7 
         Interval        =   60000
         Left            =   1920
         Top             =   840
      End
      Begin VB.Timer Timer6 
         Interval        =   500
         Left            =   2280
         Top             =   240
      End
      Begin VB.Timer Timer3 
         Interval        =   60000
         Left            =   4200
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Interval        =   40000
         Left            =   2280
         Top             =   720
      End
      Begin VB.Timer TimerTrabaja 
         Interval        =   1000
         Left            =   3960
         Top             =   120
      End
      Begin VB.Timer CmdExec 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3120
         Tag             =   "S"
         Top             =   120
      End
      Begin VB.Timer UserTimer 
         Interval        =   1000
         Left            =   2760
         Top             =   -120
      End
      Begin VB.Timer TimerFatuo 
         Interval        =   2500
         Left            =   3600
         Top             =   120
      End
      Begin VB.Timer tRevisarCabs 
         Left            =   10000
         Top             =   480
      End
      Begin VB.Label CantUsuarios 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   105
      End
      Begin VB.Label lblCantUsers 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de Usuarios Online:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   2400
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mensaje BroadCast:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4935
      Begin VB.Timer Timer9 
         Interval        =   150
         Left            =   4080
         Top             =   0
      End
      Begin VB.Timer Timer8 
         Left            =   3000
         Top             =   0
      End
      Begin VB.Timer Timer5 
         Interval        =   1
         Left            =   3480
         Top             =   0
      End
      Begin VB.Timer Timer4 
         Interval        =   1
         Left            =   2400
         Top             =   0
      End
      Begin VB.TextBox BroadMsg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1275
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   360
         Width           =   4695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Enviar Mensaje BroadCast"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   4695
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   6480
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   3
      X1              =   0
      X2              =   5160
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label txStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   5520
      Width           =   45
   End
   Begin VB.Menu mnuControles 
      Caption         =   "&FlamiusAO"
      Begin VB.Menu mnuServidor 
         Caption         =   "Configuracion"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystray 
         Caption         =   "SysTray Servidor"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSeparador2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCerrar 
         Caption         =   "Cerrar"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mnuAcerca 
         Caption         =   "Acerca de ..."
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuMostrar 
         Caption         =   "&Mostrar Servidor"
      End
      Begin VB.Menu mnuSeparador3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Cerrar"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
   
Const NIM_ADD = 0
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Const WM_MBUTTONDOWN = &H207
Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Private Function setNOTIFYICONDATA(hwnd As Long, ID As Long, flags As Long, CallbackMessage As Long, Icon As Long, Tip As String) As NOTIFYICONDATA
    Dim nidTemp As NOTIFYICONDATA

    nidTemp.cbSize = Len(nidTemp)
    nidTemp.hwnd = hwnd
    nidTemp.uID = ID
    nidTemp.uFlags = flags
    nidTemp.uCallbackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = Tip & Chr$(0)

    setNOTIFYICONDATA = nidTemp
End Function

Private Sub CmdExec_Timer()
On Error Resume Next

#If UsarQueSocket = 1 Then
Dim i As Integer

For i = 1 To MaxUsers
    If UserList(i).ConnID <> -1 Then
        If Not UserList(i).CommandsBuffer.IsEmpty Then Call HandleData(i, UserList(i).CommandsBuffer.Pop)
    End If
Next i

#End If

End Sub
Private Sub cmdMore_Click()

If cmdMore.caption = "Mensaje BroadCast >>" Then
    Me.Height = 4395
    cmdMore.caption = "<< Ocultar"
Else
    Me.Height = 2070
    cmdMore.caption = "Mensaje BroadCast >>"
End If

End Sub

Private Sub Command1_Click()
Call SendData(ToAll, 0, 0, "!!" & BroadMsg.Text & ENDC)
End Sub
Public Sub InitMain(f As Byte)

If f Then
    Call mnuSystray_Click
Else: frmMain.Show
End If

End Sub

Private Sub Command2_Click()
frmConID.Show
End Sub

Private Sub Form_Load()


hay_Quest = False
ganogrupoAZUL = False
ganogrupoROJO = False
Dim LoopC As Integer
For LoopC = 1 To LastUser
UserList(LoopC).flags.Estaenlaquest = False

guerra = 11
Call mnuSystray_Click
Codifico = RandomNumber(1, 99)
Death_Cantidad = 0
yagano = 0
yaganoo = False
GuerraPremiociuda = False
GuerraPremiocrimi = False
UserList(LoopC).flags.GuerraFcrimi = False
UserList(LoopC).flags.GuerraFciuda = False
Hay_guerra = False
If UserList(LoopC).flags.Honor <= 0 Then
UserList(LoopC).flags.Honor = 0
End If
Next
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   
   If Not Visible Then
        Select Case X \ Screen.TwipsPerPixelX
                
            Case WM_LBUTTONDBLCLK
                WindowState = vbNormal
                Visible = True
                Dim hProcess As Long
                GetWindowThreadProcessId hwnd, hProcess
                AppActivate hProcess
            Case WM_RBUTTONUP
                hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf AppHook, App.hInstance, App.ThreadID)
                PopupMenu mnuPopUp, , , , mnuMostrar
                If hHook Then UnhookWindowsHookEx hHook: hHook = 0
        End Select
   End If
   
End Sub
Private Sub QuitarIconoSystray()
On Error Resume Next


Dim i As Integer
Dim nid As NOTIFYICONDATA

nid = setNOTIFYICONDATA(frmMain.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, vbNull, frmMain.Icon, "")

i = Shell_NotifyIconA(NIM_DELETE, nid)
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Call QuitarIconoSystray
#If UsarQueSocket = 1 Then
    Call LimpiaWsApi(frmMain.hwnd)
#Else
    Socket1.Cleanup
#End If

Call DescargaNpcsDat

Dim LoopC As Integer

For LoopC = 1 To MaxUsers
    If UserList(LoopC).ConnID <> -1 Then Call CloseSocket(LoopC)
Next


Call LogMain(" Server cerrado")
End

End Sub

Private Sub mnuCerrar_Click()

Call SaveGuildsNew

If MsgBox("Si cierra el servidor puede provocar la perdida de datos." & vbCrLf & vbCrLf & "¿Desea hacerlo de todas maneras?", vbYesNo + vbExclamation, "Advertencia") = vbYes Then Call ApagarSistema

End Sub
Private Sub mnusalir_Click()

Call mnuCerrar_Click

End Sub
Public Sub mnuMostrar_Click()
On Error Resume Next

WindowState = vbNormal
Form_MouseMove 0, 0, 7725, 0

End Sub
Private Sub mnuServidor_Click()

frmServidor.Visible = True

End Sub
Private Sub mnuSystray_Click()
Dim i As Integer
Dim S As String
Dim nid As NOTIFYICONDATA

S = "Servidor FlamiusAO"
nid = setNOTIFYICONDATA(frmMain.hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, frmMain.Icon, S)
i = Shell_NotifyIconA(NIM_ADD, nid)
    
If WindowState <> vbMinimized Then WindowState = vbMinimized
Visible = False

End Sub
Private Sub Socket1_Blocking(Status As Integer, Cancel As Integer)
Cancel = True
End Sub
Private Sub Socket2_Connect(Index As Integer)

Set UserList(Index).CommandsBuffer = New CColaArray

End Sub
Private Sub Socket2_Disconnect(Index As Integer)

If UserList(Index).flags.UserLogged And _
    UserList(Index).Counters.Saliendo = False Then
    Call Cerrar_Usuario(Index)
Else: Call CloseSocket(Index)
End If
UserList(Index).flags.Desconecto = UserList(Index).flags.Desconecto + 1

If UserList(Index).flags.Desconecto >= 25 Then
Call SendData(ToAdmins, 0, 0, "||" & UserList(Index).Name & " Tira el server, IP: " & UserList(Index).ip & FONTTYPE_TALK)
 End If
End Sub
Private Sub Socket2_Read(Index As Integer, DataLength As Integer, IsUrgent As Integer)


#If UsarQueSocket = 0 Then
On Error GoTo ErrorHandler
Dim LoopC As Integer
Dim RD As String
Dim rBuffer(1 To COMMAND_BUFFER_SIZE) As String
Dim CR As Integer
Dim tChar As String
Dim sChar As Integer
Dim eChar As Integer
Dim AUX$
Dim OrigCad As String
Dim LenRD As Long

Call Socket2(Index).Read(RD, DataLength)

OrigCad = RD
LenRD = Len(RD)

If LenRD = 0 Then
    UserList(Index).AntiCuelgue = UserList(Index).AntiCuelgue + 1
    If UserList(Index).AntiCuelgue >= 150 Then
        UserList(Index).AntiCuelgue = 0
        Call LogError("!!!! Detectado bucle infinito de eventos socket2_read. cerrando indice " & Index)
        Socket2(Index).Disconnect
        Call CloseSocket(Index)
        Exit Sub
    End If
Else
    UserList(Index).AntiCuelgue = 0
End If

If Len(UserList(Index).RDBuffer) > 0 Then
    RD = UserList(Index).RDBuffer & RD
    UserList(Index).RDBuffer = ""
End If

sChar = 1
For LoopC = 1 To LenRD

    tChar = Mid$(RD, LoopC, 1)

    If tChar = ENDC Then
        CR = CR + 1
        eChar = LoopC - sChar
        rBuffer(CR) = Mid$(RD, sChar, eChar)
        sChar = LoopC + 1
    End If
        
Next LoopC

If Len(RD) - (sChar - 1) <> 0 Then UserList(Index).RDBuffer = Mid$(RD, sChar, Len(RD))

For LoopC = 1 To CR
    If ClientsCommandsQueue = 1 Then
        If Len(rBuffer(LoopC)) > 0 Then If Not UserList(Index).CommandsBuffer.Push(rBuffer(LoopC)) Then Call Cerrar_Usuario(Index)
    Else
        If UserList(Index).ConnID <> -1 Then
          Call HandleData(Index, rBuffer(LoopC))
        Else
          Exit Sub
        End If
    End If
Next LoopC

Exit Sub

ErrorHandler:
    Call LogError("Error en Socket read. " & Err.Description & " Numero paquetes:" & UserList(Index).NumeroPaquetesPorMiliSec & " . Rdata:" & OrigCad)
    Call CloseSocket(Index)
#End If
End Sub


Private Sub retos2vs2_Timer()

If OPCDuelos.OCUP Then
    OPCDuelos.Tiempo = OPCDuelos.Tiempo - 1
    If OPCDuelos.Tiempo <= 0 Then
        UserList(OPCDuelos.J1).Reto.Received_Request = False
        UserList(OPCDuelos.J1).Reto.Send_Request = False
        UserList(OPCDuelos.J1).Reto.Retando_2 = False
       
        UserList(OPCDuelos.J2).Reto.Received_Request = False
        UserList(OPCDuelos.J2).Reto.Send_Request = False
        UserList(OPCDuelos.J2).Reto.Retando_2 = False
       
        UserList(OPCDuelos.J3).Reto.Received_Request = False
        UserList(OPCDuelos.J3).Reto.Send_Request = False
        UserList(OPCDuelos.J3).Reto.Retando_2 = False
       
        UserList(OPCDuelos.J4).Reto.Received_Request = False
        UserList(OPCDuelos.J4).Reto.Send_Request = False
        UserList(OPCDuelos.J4).Reto.Retando_2 = False
       
        Call WarpUserChar(OPCDuelos.J1, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y) 'los mando a ulla
        Call WarpUserChar(OPCDuelos.J2, ULLATHORPE.Map, ULLATHORPE.X + 1, ULLATHORPE.Y) 'los mando a ulla
        Call WarpUserChar(OPCDuelos.J3, ULLATHORPE.Map, ULLATHORPE.X - 1, ULLATHORPE.Y) 'los mando a ulla
        Call WarpUserChar(OPCDuelos.J4, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y + 1) 'los mando a ulla
       
        frmMain.retos2vs2.Enabled = False '> CUANDO CREEN EL TIMER, CAMBIENLEN EL NOMBRE.
        OPCDuelos.OCUP = False
        OPCDuelos.Tiempo = 0
    End If
End If

End Sub

Private Sub Timer1_Timer()

' esta es la variable que cree en el mod declaraciones, llamada Torneo (Public AutoTorneo As Integer)
AutoTorneo = AutoTorneo + 1
Select Case AutoTorneo
Case 84
Call SendData(ToAll, 0, 0, "||Torneo> En 10 minutos se realizará un torneo automatico." & FONTTYPE_GUILD)
Case 89
Call SendData(ToAll, 0, 0, "||Torneo> En 5 minutos se realizará un torneo automatico." & FONTTYPE_GUILD)
Case 93
Call SendData(ToAll, 0, 0, "||Torneo> En 1 minutos se realizará un torneo automatico." & FONTTYPE_GUILD)
Case 94
Call torneos_auto(RandomNumber(1, 3)) ' con esto se hace un random si el torneo sera de 2 a 32 participantes.
Case 96
If Torneo_Esperando = True Then
Call Torneoauto_Cancela
AutoTorneo = 2
Else
AutoTorneo = 2
End If
End Select
End Sub
 

Private Sub Timer10_Timer()
Dim i As Integer
For i = 1 To LastUser
If hay_Quest = False And UserList(i).flags.GrupoRojo Then
UserList(i).flags.GrupoAzul = False
UserList(i).flags.GrupoRojo = False
End If
If hay_Quest = False And UserList(i).flags.GrupoAzul Then
UserList(i).flags.GrupoAzul = False
UserList(i).flags.GrupoRojo = False
End If
If UserList(i).flags.Estaenlaquest = True And hay_Quest = False Then
UserList(i).flags.GrupoRojo = False
UserList(i).flags.GrupoAzul = False
Call WarpUserChar(i, 1, 50, 50, True)
UserList(i).flags.Estaenlaquest = False
End If
Next
End Sub

Private Sub Timer11_Timer()
Dim userindex As Integer
Dim aa As Integer
aaa = Death + Death - 1
 Call SendData(ToIndex, userindex, 0, "||AutoDeath> Fraude, no puedes reclamar un trofeo ya que hay mas de 1 usuario en el mapa" & Death_Cantidad & "!!~255~0~255~0~0" & FONTTYPE_INFO)

 If Hay_Death = False Or Death_Cantidad > aaa Then
 Call SendData(ToIndex, userindex, 0, "||AutoDeath> Fraude, no puedes reclamar un trofeo ya que hay mas de 1 usuario en el mapa" & Death_Cantidad & "!!~255~0~255~0~0" & FONTTYPE_INFO)
 Exit Sub
 End If

Death_Gano = Death_Gano + 1

If Death_Gano = 1 Then

 Call SendData(ToAll, 0, 0, "||AutoDeath> El deathmatch a terminado, el usuario ganador debe tipear /reclamar ~255~0~255~0~0" & FONTTYPE_INFO)
 Death_Gano = Death_Gano + 1
 
 End If
 
End Sub

Private Sub Timer2_Timer()
Dim userindex As Integer
Dim aa As Integer
aaa = Death + Death - 1
If Hay_Death = False Then Exit Sub
If Death_Cantidad = aaa Then
death_termina = death_termina + 1
End If
If death_termina = 1 Then

 Call SendData(ToAll, 0, 0, "||AutoDeath> El deathmatch a terminado, el usuario ganador debe tipear /reclamar ~255~0~255~0~0" & FONTTYPE_INFO)
 Death_Gano = Death_Gano + 1
 
 End If
 
End Sub

Private Sub Timer3_Timer()
If asdd = 31 Then
Call SendData(ToAll, 0, 0, "||WEB:  www.ValZhatAO.ucoz.com " & FONTTYPE_GUILD)
asdd = 0
End If
If asdd = 20 Then
Call SendData(ToAll, 0, 0, "|| Para preguntas/dudas/reporte de bugs/postulaciones/info/etc FORO: http://www.servidoresonline.org/f17-kreizy-ao " & FONTTYPE_GUILD)
End If
If asdd = 15 Then
Call SendData(ToAll, 0, 0, "||ValZhatAO@hotmail.com Agreganos al MSN!!" & FONTTYPE_GUILD)
End If
asdd = asdd + 1

End Sub

Public Sub GuerraAutomatica(userindex As Integer)

Dim aa As Integer
Dim LoopC As Integer

 
If Hay_guerra = False Then Exit Sub

       If yagano = 0 Then
       If UserList(userindex).flags.Enguerra = True Or Hay_guerra = True Then
           If UserList(userindex).Faccion.Bando = Real Or GuerraPremiociuda = True Then
 Call SendData(ToAll, 0, 0, "||GuerraFaccionaria> La guerra faccionaria la ganaron los ciudas! ~255~0~255~0~0" & FONTTYPE_INFO)
 Call SendData(ToAll, 0, 0, "||GuerraFaccionaria> Premios repartidos a los ganadores!(1 de canjeo) ~255~0~255~0~0" & FONTTYPE_INFO)
UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 1
GuerraPremiocrimi = False
GuerraPremiociuda = False
yagano = 1
ganociuda = 0
ganocrimi = 0
guerra_crimis = 0
guerra_ciudas = 0
Call WarpUserChar(userindex, 1, 50, 50)
Hay_guerra = False
UserList(userindex).flags.Enguerra = False
End If
End If
           End If
           
           If yagano = 0 Then
           
                    
If UserList(userindex).Faccion.Bando = Caos Or GuerraPremiocrimi = True Or UserList(userindex).flags.Enguerra = True Or Hay_guerra = True Then
 Call SendData(ToAll, 0, 0, "||GuerraFaccionaria> La guerra faccionaria la ganaron los crimis! ~255~0~255~0~0" & FONTTYPE_INFO)
 Call SendData(ToAll, 0, 0, "||GuerraFaccionaria> Premios repartidos a los ganadores!(1 de canjeo) ~255~0~255~0~0" & FONTTYPE_INFO)
UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 1
GuerraPremiocrimi = False
GuerraPremiociuda = False
yagano = 1
ganociuda = 0
ganocrimi = 0
guerra_crimis = 0
guerra_ciudas = 0
Call WarpUserChar(userindex, 1, 50, 50)
Hay_guerra = False
UserList(userindex).flags.Enguerra = False
End If
           End If
          
     

End Sub
Private Sub Timer4_Timer()
Dim Ganador As Integer

Dim i As Integer

For i = 1 To LastUser
 
If Hay_guerra = False Then Exit Sub

       If yagano = 0 Then
       If UserList(i).flags.Enguerra = True And Hay_guerra = True Then
           If UserList(i).flags.GuerraFciuda = True And GuerraPremiociuda = True Then
 Ganador = UserList(i).flags.GuerraFciuda
 Call SendData(ToAll, 0, 0, "||GuerraFaccionaria> La guerra faccionaria la ganaron los ciudas! ~255~0~255~0~0" & FONTTYPE_INFO)
 Call SendData(ToAll, 0, 0, "||GuerraFaccionaria> Premios repartidos a los ganadores!(1 de canjeo) ~255~0~255~0~0" & FONTTYPE_INFO)
UserList(i).flags.Canje = UserList(i).flags.Canje + 1
GuerraPremiocrimi = False
GuerraPremiociuda = False
yagano = 1
ganociuda = 0
ganocrimi = 0
guerra_crimis = 0
guerra_ciudas = 0
UserList(i).flags.GuerraFciuda = False
Call WarpUserChar(i, 1, 50, 50)
Hay_guerra = False
UserList(i).flags.Enguerra = False
End If
End If
           End If
           
           If yagano = 0 Then
           
                    
If UserList(i).flags.GuerraFcrimi = True And GuerraPremiocrimi = True And UserList(i).flags.Enguerra = True And Hay_guerra = True Then
 Call SendData(ToAll, 0, 0, "||GuerraFaccionaria> La guerra faccionaria la ganaron los crimis! ~255~0~255~0~0" & FONTTYPE_INFO)
 Call SendData(ToAll, 0, 0, "||GuerraFaccionaria> Premios repartidos a los ganadores!(1 de canjeo) ~255~0~255~0~0" & FONTTYPE_INFO)
UserList(i).flags.Canje = UserList(i).flags.Canje + 1
GuerraPremiocrimi = False
GuerraPremiociuda = False
yagano = 1
ganociuda = 0
ganocrimi = 0
guerra_crimis = 0
guerra_ciudas = 0
UserList(i).flags.GuerraFcrimi = False
Call WarpUserChar(i, 1, 50, 50)
Hay_guerra = False
UserList(i).flags.Enguerra = False
End If
           End If
         Next
End Sub

Private Sub Timer5_Timer()
Dim jaja As Integer
Dim tomopocion As Boolean
If tomopocion = True Then
jaja = jaja + 1
Call SendData(ToAdmins, 0, 0, "|| El Usuario paso un intervalo" & "~0~50~0~0~0")
End If
End Sub

Private Sub Timer6_Timer()
Dim userindex As Integer
Dim maximooo As Integer
maximooo = 7
If intervalovida > maximooo Then
intervalovida = 0
MsgBox "Los gms te están vigilandoo"

Else

intervalovida = 0
End If
End Sub

Private Sub Timer7_Timer()

If ii = 30 Then
ii = 0
Call DarPremioCastillos
End If
ii = ii + 1

End Sub

Private Sub Timer8_Timer()
If Anticheatt = True Then
Anticheatt = False
Else
Anticheatt = True
End If
End Sub

Private Sub Timer9_Timer()
Dim userindex As Integer

If pocionroja >= 3 Then
Call SendData(ToAdmins, 0, 0, "||Anticheat> El usuario " & UserList(userindex).Name & " pasó el intervalo de pociones." & "~0~50~0~0~0")
pocionroja = 0
Else
pocionroja = 0
End If

End Sub

Private Sub TimerFatuo_Timer()
On Error GoTo Error
Dim i As Integer
Dim d As Integer


For i = 1 To LastNPC
    If Npclist(i).flags.NPCActive And Npclist(i).Numero = 89 Then Npclist(i).CanAttack = 1
Next i

Exit Sub

Error:
    Call LogError("Error en TimerFatuo: " & Err.Description)
    For d = 1 To LastUser
   If UserList(i).flags.GrupoAzul = True And hay_Quest = False Then
   Call WarpUserChar(i, 1, 50, 50, True)
   End If
     If UserList(i).flags.GrupoRojo = True And hay_Quest = False Then
   Call WarpUserChar(i, 1, 50, 50, True)
   End If
   Next d
End Sub
Private Sub TimerMeditar_Timer()
Dim i As Integer

For i = 1 To LastUser
    If UserList(i).flags.Meditando Then Call TimerMedita(i)
Next

End Sub
Sub TimerMedita(userindex As Integer)
Dim Cant As Single

If TiempoTranscurrido(UserList(userindex).Counters.tInicioMeditar) >= TIEMPO_INICIOMEDITAR Then
    Cant = UserList(userindex).Counters.ManaAcumulado + UserList(userindex).Stats.MaxMAN * (1 + UserList(userindex).Stats.UserSkills(Meditar) * 0.05) / 100
    If Cant <= 0.75 Then
        UserList(userindex).Counters.ManaAcumulado = Cant
        Exit Sub
    Else
        Cant = Round(Cant)
        UserList(userindex).Counters.ManaAcumulado = 0
    End If
    Call AddtoVar(UserList(userindex).Stats.MinMAN, Cant, UserList(userindex).Stats.MaxMAN)
    Call SendData(ToIndex, userindex, 0, "MN" & Cant)
    Call SubirSkill(userindex, Meditar)
    If UserList(userindex).Stats.MinMAN >= UserList(userindex).Stats.MaxMAN Then
        Call SendData(ToIndex, userindex, 0, "D9")
        Call SendData(ToIndex, userindex, 0, "MEDOK")
        UserList(userindex).flags.Meditando = False
        UserList(userindex).Char.FX = 0
        UserList(userindex).Char.loops = 0
        Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & 0 & "," & 0)
    End If
End If

Call SendUserMANA(userindex)

End Sub
Private Sub TimerTrabaja_Timer()
Dim i As Integer
On Error GoTo Error

For i = 1 To LastUser
    If UserList(i).flags.Trabajando Then
        UserList(i).Counters.IdleCount = Timer
        
        Select Case UserList(i).flags.Trabajando
            Case Pesca
                Call DoPescar(i)
                    
            Case Talar
                Call DoTalar(i, ObjData(MapData(UserList(i).POS.Map, UserList(i).TrabajoPos.X, UserList(i).TrabajoPos.Y).OBJInfo.OBJIndex).ArbolElfico = 1)
    
            Case Mineria
                Call DoMineria(i, ObjData(MapData(UserList(i).POS.Map, UserList(i).TrabajoPos.X, UserList(i).TrabajoPos.Y).OBJInfo.OBJIndex).MineralIndex)
        End Select
    End If
Next
Exit Sub
Error:
    Call LogError("Error en TimerTrabaja: " & Err.Description)
    
End Sub

Public Sub Tlimpiar_Timer()
MinutosTLimpiar = MinutosTLimpiar + 1
If MinutosTLimpiar = 2 Then
Call SendData(ToAll, 0, 0, "||Se realizará una limpieza del Mundo en 30 segundos. Por favor recojan sus items." & FONTTYPE_FENIX)
End If
If MinutosTLimpiar = 3 Then
Call SendData(ToAll, 0, 0, "||Se realizará una limpieza del Mundo en 15 segundos. Por favor recojan sus items." & FONTTYPE_FENIX)
End If
If MinutosTLimpiar = 4 Then
Call LimpiarItemsMundo
MinutosTLimpiar = 1
End If
End Sub

Private Sub UserTimer_Timer()
On Error GoTo Error
Static Andaban As Boolean, Contador As Single
Dim Andan As Boolean, UI As Integer, i As Integer, XXN As Integer

'matute
If encuestas.activa = 1 Then
    encuestas.Tiempo = encuestas.Tiempo + 1
    If encuestas.Tiempo = 15 Then
        Call SendData(ToAll, 0, 0, "||Faltan 15 segundos para finalizar la encuesta." & FONTTYPE_TALK)
    ElseIf encuestas.Tiempo = 30 Then
        Call SendData(ToAll, 0, 0, "||RESULTADOS DE LA ENCUESTA:" & FONTTYPE_FENIX)
        Call SendData(ToAll, 0, 0, "||VOTOS POSITIVOS: " & encuestas.votosSI & " | VOTOS NEGATIVOS: " & encuestas.votosNP & FONTTYPE_TALK)
        If encuestas.votosNP < encuestas.votosSI Then
            Call SendData(ToAll, 0, 0, "||Opción ganadora: SI" & FONTTYPE_FENIX)
        ElseIf encuestas.votosSI < encuestas.votosNP Then
            Call SendData(ToAll, 0, 0, "||Opción ganadora: NO" & FONTTYPE_FENIX)
        ElseIf encuestas.votosNP = encuestas.votosSI Then
            Call SendData(ToAll, 0, 0, "||Opción ganadora: NINGUNA - EMPATE" & FONTTYPE_FENIX)
        End If
        encuestas.activa = 0
        encuestas.Tiempo = 0
        encuestas.votosNP = 0
        encuestas.votosSI = 0
        For XXN = 1 To LastUser
            If UserList(XXN).flags.votoencuesta = 1 Then UserList(XXN).flags.votoencuesta = 0
        Next XXN
    End If
    Exit Sub
End If
'//////'matute

If CuentaRegresiva Then
    CuentaRegresiva = CuentaRegresiva - 1
    
    If CuentaRegresiva = 0 Then
        Call SendData(ToMap, 0, GMCuenta, "||YA!!!" & FONTTYPE_FIGHT)
        Me.Enabled = False
    Else
        Call SendData(ToMap, 0, GMCuenta, "||" & CuentaRegresiva & "..." & FONTTYPE_INFO)
    End If
End If

For i = 1 To LastUser
    If UserList(i).ConnID <> -1 Then DayStats.Segundos = DayStats.Segundos + 1
Next

If TiempoTranscurrido(Contador) >= 10 Then
    Contador = Timer
    Andan = EstadisticasWeb.EstadisticasAndando()
    If Not Andaban And Andan Then Call InicializaEstadisticas
    Andaban = Andan
End If

For UI = 1 To LastUser
    If UserList(UI).flags.UserLogged And UserList(UI).ConnID <> -1 Then
        Call TimerPiquete(UI)
        If UserList(UI).flags.Protegido > 1 Then Call TimerProtEntro(UI)
        If UserList(UI).flags.Encarcelado Then Call TimerCarcel(UI)
        If UserList(UI).flags.Muerto = 0 Then
            If UserList(UI).flags.Paralizado Then Call TimerParalisis(UI)
            If UserList(UI).flags.BonusFlecha Then Call TimerFlecha(UI)
            If UserList(UI).flags.Ceguera = 1 Then Call TimerCeguera(UI)
            If UserList(UI).flags.Envenenado = 1 Then Call TimerVeneno(UI)
            If UserList(UI).flags.Envenenado = 2 Then Call TimerVenenoDoble(UI)
            If UserList(UI).flags.Estupidez = 1 Then Call TimerEstupidez(UI)
            If UserList(UI).flags.AdminInvisible = 0 And UserList(UI).flags.Invisible = 1 And UserList(UI).flags.Oculto = 0 Then Call TimerInvisibilidad(UI)
            If UserList(UI).flags.Desnudo = 1 Then Call TimerFrio(UI)
            If UserList(UI).flags.tomopocion Then Call TimerPocion(UI)
            If UserList(UI).flags.Transformado Then Call TimerTransformado(UI)
            If UserList(UI).NroMascotas Then Call TimerInvocacion(UI)
            If UserList(UI).flags.Oculto Then Call TimerOculto(UI)
            If UserList(UI).flags.Sacrificando Then Call TimerSacrificando(UI)
            
            Call TimerHyS(UI)
            Call TimerSanar(UI)
            Call TimerStamina(UI)
        End If
        If EnviarEstats Then
            Call SendUserStatsBox(UI)
            EnviarEstats = False
        End If
        Call TimerIdleCount(UI)
        If UserList(UI).Counters.Saliendo Then Call TimerSalir(UI)
    End If
Next

Exit Sub

Error:
    Call LogError("Error en UserTimer:" & Err.Description & " " & UI)
    
End Sub
Public Sub TimerOculto(userindex As Integer)
Dim ClaseBuena As Boolean

ClaseBuena = UserList(userindex).Clase = GUERRERO Or UserList(userindex).Clase = ARQUERO Or UserList(userindex).Clase = CAZADOR

If RandomNumber(1, 10 + UserList(userindex).Stats.UserSkills(Ocultarse) / 4 + 15 * Buleano(ClaseBuena) + 25 * Buleano(ClaseBuena And Not UserList(userindex).Clase = GUERRERO And UserList(userindex).Invent.ArmourEqpObjIndex = 360)) <= 5 Then
    UserList(userindex).flags.Oculto = 0
    UserList(userindex).flags.Invisible = 0
    Call SendData(ToMap, 0, UserList(userindex).POS.Map, ("V3" & UserList(userindex).Char.CharIndex & ",0"))
    Call SendData(ToIndex, userindex, 0, "V5")
End If

End Sub
Public Sub TimerStamina(userindex As Integer)

If UserList(userindex).Stats.MinSta < UserList(userindex).Stats.MaxSta And UserList(userindex).flags.Hambre = 0 And UserList(userindex).flags.Sed = 0 And UserList(userindex).flags.Desnudo = 0 Then
   If (Not UserList(userindex).flags.Descansar And TiempoTranscurrido(UserList(userindex).Counters.STACounter) >= StaminaIntervaloSinDescansar) Or _
   (UserList(userindex).flags.Descansar And TiempoTranscurrido(UserList(userindex).Counters.STACounter) >= StaminaIntervaloDescansar) Then
        UserList(userindex).Counters.STACounter = Timer
        UserList(userindex).Stats.MinSta = Minimo(UserList(userindex).Stats.MinSta + CInt(RandomNumber(5, Porcentaje(UserList(userindex).Stats.MaxSta, 15))), UserList(userindex).Stats.MaxSta)
        If TiempoTranscurrido(UserList(userindex).Counters.CartelStamina) >= 10 Then
            UserList(userindex).Counters.CartelStamina = Timer
            Call SendData(ToIndex, userindex, 0, "MV")
        End If
        EnviarEstats = True
    End If
End If

End Sub
Sub TimerTransformado(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Transformado) >= IntervaloInvisible Then
    Call DoTransformar(userindex)
End If

End Sub
Sub TimerInvisibilidad(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Invisibilidad) >= IntervaloInvisible Then
    Call SendData(ToIndex, userindex, 0, "V6")
    Call QuitarInvisible(userindex)
End If

End Sub
Sub TimerFlecha(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.BonusFlecha) >= 45 Then
    UserList(userindex).Counters.BonusFlecha = 0
    UserList(userindex).flags.BonusFlecha = False
    Call SendData(ToIndex, userindex, 0, "||Se acabó el efecto del Arco Encantado." & FONTTYPE_INFO)
End If

End Sub
Sub TimerPiquete(userindex As Integer)

If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y).trigger = 5 Then
    UserList(userindex).Counters.PiqueteC = UserList(userindex).Counters.PiqueteC + 1
    If UserList(userindex).Counters.PiqueteC Mod 5 = 0 Then Call SendData(ToIndex, userindex, 0, "9N")
    If UserList(userindex).Counters.PiqueteC >= 25 Then
        UserList(userindex).Counters.PiqueteC = 0
        Call Encarcelar(userindex, 3)
    End If
Else: UserList(userindex).Counters.PiqueteC = 0
End If

End Sub
Public Sub TimerProtEntro(userindex As Integer)
On Error GoTo Error

UserList(userindex).Counters.Protegido = UserList(userindex).Counters.Protegido - 1
If UserList(userindex).Counters.Protegido <= 0 Then UserList(userindex).flags.Protegido = 0

Exit Sub

Error:
    Call LogError("Error en TimerProtEntro" & " " & Err.Description)
End Sub
Sub TimerParalisis(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Paralisis) >= IntervaloParalizadoUsuario Then
    UserList(userindex).Counters.Paralisis = 0
    UserList(userindex).flags.Paralizado = 0
    Call SendData(ToIndex, userindex, 0, "P8")
End If

End Sub
Sub TimerCeguera(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Ceguera) >= IntervaloParalizadoUsuario / 2 Then
    UserList(userindex).Counters.Ceguera = 0
    UserList(userindex).flags.Ceguera = 0
    Call SendData(ToIndex, userindex, 0, "NSEGUE")
End If

End Sub
Sub TimerEstupidez(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Estupidez) >= IntervaloParalizadoUsuario Then
    UserList(userindex).Counters.Estupidez = 0
    UserList(userindex).flags.Estupidez = 0
    Call SendData(ToIndex, userindex, 0, "NESTUP")
End If

End Sub
Sub TimerCarcel(userindex As Integer)

Dim j As Byte

If TiempoTranscurrido(UserList(userindex).Counters.Pena) >= UserList(userindex).Counters.TiempoPena Then
    UserList(userindex).Counters.TiempoPena = 0
    UserList(userindex).flags.Encarcelado = 0
    UserList(userindex).Counters.Pena = 0
    If UserList(userindex).POS.Map = Prision.Map Then
        Call WarpUserChar(userindex, Libertad.Map, Libertad.X, Libertad.Y, True)
        Call SendData(ToIndex, userindex, 0, "4P")
    End If
End If

End Sub
Sub TimerVenenoDoble(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Veneno) >= 2 Then
    If TiempoTranscurrido(UserList(userindex).flags.EstasEnvenenado) >= 8 Then
        UserList(userindex).flags.Envenenado = 0
        UserList(userindex).flags.EstasEnvenenado = 0
        UserList(userindex).Counters.Veneno = 0
    Else
        Call SendData(ToIndex, userindex, 0, "1M")
        UserList(userindex).Counters.Veneno = Timer
        If Not UserList(userindex).flags.Quest Then
            UserList(userindex).Stats.MinHP = Maximo(0, UserList(userindex).Stats.MinHP - 25)
            If UserList(userindex).Stats.MinHP = 0 Then
                Call UserDie(userindex)
            Else: EnviarEstats = True
            End If
        End If
    End If
End If

End Sub
Sub UserSacrificado(userindex As Integer)
Dim MiObj As Obj

MiObj.OBJIndex = Gema
MiObj.Amount = UserList(userindex).Stats.ELV ^ 2

Call MakeObj(ToMap, userindex, UserList(userindex).POS.Map, MiObj, UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y)
Call UserDie(userindex)

UserList(UserList(userindex).flags.Sacrificador).flags.Sacrificado = 0
Call SendData(ToIndex, UserList(userindex).flags.Sacrificador, 0, "||Sacrificaste a " & UserList(userindex).Name & " por " & MiObj.Amount & " partes de la piedra filosofal." & FONTTYPE_INFO)
UserList(userindex).flags.Ban = 1
Call CloseSocket(userindex)

End Sub
Sub TimerSacrificando(userindex As Integer)

UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP - 10
UserList(UserList(userindex).flags.Sacrificador).Stats.MinMAN = Minimo(0, UserList(UserList(userindex).flags.Sacrificador).Stats.MinMAN - 50)
Call SendUserMANA(UserList(userindex).flags.Sacrificador)

If UserList(UserList(userindex).flags.Sacrificador).Stats.MinMAN = 0 Then Call CancelarSacrificio(userindex)
If UserList(userindex).Stats.MinHP <= 0 Then Call UserSacrificado(userindex)

EnviarEstats = True

End Sub
Sub TimerVeneno(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Veneno) >= IntervaloVeneno Then
    If TiempoTranscurrido(UserList(userindex).flags.EstasEnvenenado) >= IntervaloVeneno * 10 Then
        UserList(userindex).flags.Envenenado = 0
        UserList(userindex).flags.EstasEnvenenado = 0
        UserList(userindex).Counters.Veneno = 0
    Else
        Call SendData(ToIndex, userindex, 0, "1M")
        UserList(userindex).Counters.Veneno = Timer
        If Not UserList(userindex).flags.Quest Then
            UserList(userindex).Stats.MinHP = Maximo(0, UserList(userindex).Stats.MinHP - RandomNumber(1, 5))
            If UserList(userindex).Stats.MinHP = 0 Then
                Call UserDie(userindex)
            Else: EnviarEstats = True
            End If
        End If
    End If
End If

End Sub
Public Sub TimerFrio(userindex As Integer)

If UserList(userindex).flags.Privilegios > 1 Then Exit Sub

If TiempoTranscurrido(UserList(userindex).Counters.Frio) >= IntervaloFrio Then
    UserList(userindex).Counters.Frio = Timer
    If MapInfo(UserList(userindex).POS.Map).Terreno = Nieve Then
        If TiempoTranscurrido(UserList(userindex).Counters.CartelFrio) >= 5 Then
            UserList(userindex).Counters.CartelFrio = Timer
            Call SendData(ToIndex, userindex, 0, "1K")
        End If
        If Not UserList(userindex).flags.Quest Then
            UserList(userindex).Stats.MinHP = Maximo(0, UserList(userindex).Stats.MinHP - Porcentaje(UserList(userindex).Stats.MaxHP, 5))
            EnviarEstats = True
            If UserList(userindex).Stats.MinHP = 0 Then
                Call SendData(ToIndex, userindex, 0, "1L")
                Call UserDie(userindex)
            End If
        End If
    End If
    Call QuitarSta(userindex, Porcentaje(UserList(userindex).Stats.MaxSta, 5))
    If TiempoTranscurrido(UserList(userindex).Counters.CartelFrio) >= 10 Then
        UserList(userindex).Counters.CartelFrio = Timer
        Call SendData(ToIndex, userindex, 0, "FR")
    End If
    EnviarEstats = True
End If

End Sub
Sub TimerPocion(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).flags.DuracionEfecto) >= 35 Then
Call Parpa(userindex)
If TiempoTranscurrido(UserList(userindex).flags.DuracionEfecto) >= 45 Then
    UserList(userindex).flags.DuracionEfecto = 0
    UserList(userindex).flags.tomopocion = False
    UserList(userindex).Stats.UserAtributos(Agilidad) = UserList(userindex).Stats.UserAtributosBackUP(Agilidad)
    UserList(userindex).Stats.UserAtributos(fuerza) = UserList(userindex).Stats.UserAtributosBackUP(fuerza)
    Call UpdateFuerzaYAg(userindex)
End If
End If
End Sub
Public Sub TimerHyS(userindex As Integer)
Dim EnviaInfo As Boolean

If UserList(userindex).flags.Privilegios > 1 Or (UserList(userindex).Clase = TALADOR And UserList(userindex).Recompensas(1) = 2) Or UserList(userindex).flags.Quest Then Exit Sub

If TiempoTranscurrido(UserList(userindex).Counters.AGUACounter) >= IntervaloSed Then
    If UserList(userindex).flags.Sed = 0 Then
        UserList(userindex).Stats.MinAGU = UserList(userindex).Stats.MinAGU - 10
        If UserList(userindex).Stats.MinAGU <= 0 Then
            UserList(userindex).Stats.MinAGU = 0
            UserList(userindex).flags.Sed = 1
        End If
        EnviaInfo = True
    End If
    UserList(userindex).Counters.AGUACounter = Timer
End If

If TiempoTranscurrido(UserList(userindex).Counters.COMCounter) >= IntervaloHambre Then
    If UserList(userindex).flags.Hambre = 0 Then
        UserList(userindex).Counters.COMCounter = Timer
        UserList(userindex).Stats.MinHam = UserList(userindex).Stats.MinHam - 10
        If UserList(userindex).Stats.MinHam <= 0 Then
            UserList(userindex).Stats.MinHam = 0
            UserList(userindex).flags.Hambre = 1
        End If
        EnviaInfo = True
    End If
    UserList(userindex).Counters.COMCounter = Timer
End If

If EnviaInfo Then Call EnviarHambreYsed(userindex)

End Sub
Sub TimerSanar(userindex As Integer)

If (UserList(userindex).flags.Descansar And TiempoTranscurrido(UserList(userindex).Counters.HPCounter) >= SanaIntervaloDescansar) Or _
     (Not UserList(userindex).flags.Descansar And TiempoTranscurrido(UserList(userindex).Counters.HPCounter) >= SanaIntervaloSinDescansar) Then
    If (Not Lloviendo Or Not Intemperie(userindex)) And UserList(userindex).Stats.MinHP < UserList(userindex).Stats.MaxHP And UserList(userindex).flags.Hambre = 0 And UserList(userindex).flags.Sed = 0 Then
        If UserList(userindex).flags.Descansar Then
            UserList(userindex).Stats.MinHP = Minimo(UserList(userindex).Stats.MaxHP, UserList(userindex).Stats.MinHP + Porcentaje(UserList(userindex).Stats.MaxHP, 20))
            If UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MinHP And UserList(userindex).Stats.MaxSta = UserList(userindex).Stats.MinSta Then
                Call SendData(ToIndex, userindex, 0, "DOK")
                Call SendData(ToIndex, userindex, 0, "DN")
                UserList(userindex).flags.Descansar = False
            End If
        Else
            UserList(userindex).Stats.MinHP = Minimo(UserList(userindex).Stats.MaxHP, UserList(userindex).Stats.MinHP + Porcentaje(UserList(userindex).Stats.MaxHP, 5))
        End If
        Call SendData(ToIndex, userindex, 0, "1N")
        EnviarEstats = True
    End If
    UserList(userindex).Counters.HPCounter = Timer
End If
    
End Sub
Sub TimerInvocacion(userindex As Integer)
Dim i As Integer
Dim NpcIndex As Integer

If UserList(userindex).flags.Privilegios > 0 Or UserList(userindex).flags.Quest Then Exit Sub

For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(userindex).flags.Quest)
    If UserList(userindex).MascotasIndex(i) Then
        NpcIndex = UserList(userindex).MascotasIndex(i)
        If Npclist(NpcIndex).Contadores.TiempoExistencia > 0 And TiempoTranscurrido(Npclist(NpcIndex).Contadores.TiempoExistencia) >= IntervaloInvocacion + 10 * Buleano(Npclist(NpcIndex).Numero = 92) Then Call MuereNpc(NpcIndex, 0)
    End If
Next

End Sub
Public Sub TimerIdleCount(userindex As Integer)

If UserList(userindex).flags.Privilegios = 0 And UserList(userindex).flags.Trabajando = 0 And TiempoTranscurrido(UserList(userindex).Counters.IdleCount) >= IntervaloParaConexion And Not UserList(userindex).Counters.Saliendo Then
    Call SendData(ToIndex, userindex, 0, "!!Demasiado tiempo inactivo. Has sido desconectado..")
    Call SendData(ToIndex, userindex, 0, "FINOK")
    Call CloseSocket(userindex)
End If

End Sub
Sub TimerSalir(userindex As Integer)

If TiempoTranscurrido(UserList(userindex).Counters.Salir) >= IntervaloCerrarConexion Then
    Call SendData(ToIndex, userindex, 0, "FINOK")
    Call CloseSocket(userindex)
End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub
