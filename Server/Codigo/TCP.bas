Attribute VB_Name = "TCP"
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

Private Const MAXENVIOS As Integer = 10
Public usercorreo As String

Public Const SOCKET_BUFFER_SIZE = 3072
Public Enpausa As Boolean

Public Const COMMAND_BUFFER_SIZE = 1000
Public entorneo As Byte

Public Const NingunArma = 2
Dim Response As String
Dim Start As Single, Tmr As Single

Public Const ToIndex = 0
Public Const ToAll = 1
Public Const ToMap = 2
Public Const ToPCArea = 3
Public Const ToNone = 4
Public Const ToAllButIndex = 5
Public Const ToMapButIndex = 6
Public Const ToGM = 7
Public Const ToNPCArea = 8
Public Const ToGuildMembers = 9
Public Const ToAdmins = 10
Public Const ToPCAreaButIndex = 11
Public Const ToMuertos = 12
Public Const ToPCAreaVivos = 13
Public Const ToNPCAreaG = 14
Public Const ToPCAreaButIndexG = 15
Public Const ToGMArea = 16
Public Const ToPCAreaG = 17
Public Const ToAlianza = 18
Public Const ToCaos = 19
Public Const ToParty = 20
Public Const ToMoreAdmins = 21
Public Const ToConse = 22
Public Const ToConci = 23

#If UsarQueSocket = 0 Then
Public Const INVALID_HANDLE = -1
Public Const CONTROL_ERRIGNORE = 0
Public Const CONTROL_ERRDISPLAY = 1



Public Const SOCKET_OPEN = 1
Public Const SOCKET_CONNECT = 2
Public Const SOCKET_LISTEN = 3
Public Const SOCKET_ACCEPT = 4
Public Const SOCKET_CANCEL = 5
Public Const SOCKET_FLUSH = 6
Public Const SOCKET_CLOSE = 7
Public Const SOCKET_MACNNECT = 7
Public Const SOCKET_ABORT = 8


Public Const SOCKET_NONE = 0
Public Const SOCKET_IDLE = 1
Public Const SOCKET_LISTENING = 2
Public Const SOCKET_CONNECTING = 3
Public Const SOCKET_ACCEPTING = 4
Public Const SOCKET_RECEIVING = 5
Public Const SOCKET_SENDING = 6
Public Const SOCKET_CLOSING = 7


Public Const AF_UNSPEC = 0
Public Const AF_UNIX = 1
Public Const AF_INET = 2


Public Const SOCK_STREAM = 1
Public Const SOCK_DGRAM = 2
Public Const SOCK_RAW = 3
Public Const SOCK_RDM = 4
Public Const SOCK_SEQPACKET = 5


Public Const IPPROTO_IP = 0
Public Const IPPROTO_ICMP = 1
Public Const IPPROTO_GGP = 2
Public Const IPPROTO_TCP = 6
Public Const IPPROTO_PUP = 12
Public Const IPPROTO_UDP = 17
Public Const IPPROTO_IDP = 22
Public Const IPPROTO_ND = 77
Public Const IPPROTO_RAW = 255
Public Const IPPROTO_MAX = 256



Public Const INADDR_ANY = "0.0.0.0"
Public Const INADDR_LOOPBACK = "127.0.0.1"
Public Const INADDR_NONE = "255.055.255.255"


Public Const SOCKET_READ = 0
Public Const SOCKET_WRITE = 1
Public Const SOCKET_READWRITE = 2


Public Const SOCKET_ERRIGNORE = 0
Public Const SOCKET_ERRDISPLAY = 1


Public Const WSABASEERR = 24000
Public Const WSAEINTR = 24004
Public Const WSAEBADF = 24009
Public Const WSAEACCES = 24013
Public Const WSAEFAULT = 24014
Public Const WSAEINVAL = 24022
Public Const WSAEMFILE = 24024
Public Const WSAEWOULDBLOCK = 24035
Public Const WSAEINPROGRESS = 24036
Public Const WSAEALREADY = 24037
Public Const WSAENOTSOCK = 24038
Public Const WSAEDESTADDRREQ = 24039
Public Const WSAEMSGSIZE = 24040
Public Const WSAEPROTOTYPE = 24041
Public Const WSAENOPROTOOPT = 24042
Public Const WSAEPROTONOSUPPORT = 24043
Public Const WSAESOCKTNOSUPPORT = 24044
Public Const WSAEOPNOTSUPP = 24045
Public Const WSAEPFNOSUPPORT = 24046
Public Const WSAEAFNOSUPPORT = 24047
Public Const WSAEADDRINUSE = 24048
Public Const WSAEADDRNOTAVAIL = 24049
Public Const WSAENETDOWN = 24050
Public Const WSAENETUNREACH = 24051
Public Const WSAENETRESET = 24052
Public Const WSAECONNABORTED = 24053
Public Const WSAECONNRESET = 24054
Public Const WSAENOBUFS = 24055
Public Const WSAEISCONN = 24056
Public Const WSAENOTCONN = 24057
Public Const WSAESHUTDOWN = 24058
Public Const WSAETOOMANYREFS = 24059
Public Const WSAETIMEDOUT = 24060
Public Const WSAECONNREFUSED = 24061
Public Const WSAELOOP = 24062
Public Const WSAENAMETOOLONG = 24063
Public Const WSAEHOSTDOWN = 24064
Public Const WSAEHOSTUNREACH = 24065
Public Const WSAENOTEMPTY = 24066
Public Const WSAEPROCLIM = 24067
Public Const WSAEUSERS = 24068
Public Const WSAEDQUOT = 24069
Public Const WSAESTALE = 24070
Public Const WSAEREMOTE = 24071
Public Const WSASYSNOTREADY = 24091
Public Const WSAVERNOTSUPPORTED = 24092
Public Const WSANOTINITIALISED = 24093
Public Const WSAHOST_NOT_FOUND = 25001
Public Const WSATRY_AGAIN = 25002
Public Const WSANO_RECOVERY = 25003
Public Const WSANO_DATA = 25004
Public Const WSANO_ADDRESS = 2500
#End If

Public Data(1 To 3, 1 To 2, 1 To 2, 1 To 2) As Double
Public Onlines(1 To 3) As Long

Public Const Minuto = 1
Public Const Hora = 2
Public Const Dia = 3

Public Const Actual = 1
Public Const Last = 2

Public Const Enviada = 1
Public Const Recibida = 2

Public Const Mensages = 1
Public Const Letras = 2

Sub DarCuerpoYCabeza(UserBody As Integer, UserHead As Integer, Raza As Byte, Gen As Byte)

Select Case Gen
   Case HOMBRE
        Select Case Raza
        
                Case HUMANO
                    UserHead = CInt(RandomNumber(1, 24))
                    If UserHead > 24 Then UserHead = 24
                    UserBody = 1
                Case ELFO
                    UserHead = CInt(RandomNumber(1, 7)) + 100
                    If UserHead > 107 Then UserHead = 107
                    UserBody = 2
                Case ELFO_OSCURO
                    UserHead = CInt(RandomNumber(1, 4)) + 200
                    If UserHead > 204 Then UserHead = 204
                    UserBody = 3
                Case ENANO
                    UserHead = RandomNumber(1, 4) + 300
                    If UserHead > 304 Then UserHead = 304
                    UserBody = 52
                Case GNOMO
                    UserHead = RandomNumber(1, 3) + 400
                    If UserHead > 403 Then UserHead = 403
                    UserBody = 52
                Case Else
                    UserHead = 1
                    UserBody = 1
            
        End Select
   Case MUJER
        Select Case Raza
                Case HUMANO
                    UserHead = CInt(RandomNumber(1, 4)) + 69
                    If UserHead > 73 Then UserHead = 73
                    UserBody = 1
                Case ELFO
                    UserHead = CInt(RandomNumber(1, 5)) + 169
                    If UserHead > 174 Then UserHead = 174
                    UserBody = 2
                Case ELFO_OSCURO
                    UserHead = CInt(RandomNumber(1, 5)) + 269
                    If UserHead > 274 Then UserHead = 274
                    UserBody = 3
                Case GNOMO
                    UserHead = RandomNumber(1, 4) + 469
                    If UserHead > 473 Then UserHead = 473
                    UserBody = 52
                Case ENANO
                    UserHead = RandomNumber(1, 3) + 369
                    If UserHead > 372 Then UserHead = 372
                    UserBody = 52
                Case Else
                    UserHead = 70
                    UserBody = 1
        End Select
End Select

   
End Sub
Function AsciiValidos(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        AsciiValidos = False
        Exit Function
    End If
Next

AsciiValidos = True

End Function

Function Numeric(ByVal cad As String) As Boolean
Dim car As Byte
Dim i As Integer

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(Mid$(cad, i, 1))
    
    If (car < 48 Or car > 57) Then
        Numeric = False
        Exit Function
    End If
    
Next

Numeric = True

End Function
Function NombrePermitido(ByVal Nombre As String) As Boolean
Dim i As Integer

For i = 1 To UBound(ForbidenNames)
    If InStr(Nombre, ForbidenNames(i)) Then
        NombrePermitido = False
        Exit Function
    End If
Next

NombrePermitido = True

End Function

Function ValidateAtrib(userindex As Integer) As Boolean
Dim LoopC As Integer

For LoopC = 1 To NUMATRIBUTOS
    If UserList(userindex).Stats.UserAtributosBackUP(LoopC) > 23 Or UserList(userindex).Stats.UserAtributosBackUP(LoopC) < 1 Then Exit Function
Next

ValidateAtrib = True

End Function

Function ValidateAtrib2(userindex As Integer) As Boolean
Dim LoopC As Integer

For LoopC = 1 To NUMATRIBUTOS
    If UserList(userindex).Stats.UserAtributosBackUP(LoopC) > 18 Or UserList(userindex).Stats.UserAtributosBackUP(LoopC) < 1 Then
    ValidateAtrib2 = False
    Exit Function
    End If
Next

ValidateAtrib2 = True

End Function
Function ValidateSkills(userindex As Integer) As Boolean
Dim LoopC As Integer

For LoopC = 1 To NUMSKILLS
    If UserList(userindex).Stats.UserSkills(LoopC) < 0 Then Exit Function
    If UserList(userindex).Stats.UserSkills(LoopC) > 100 Then UserList(userindex).Stats.UserSkills(LoopC) = 100
Next

ValidateSkills = True

End Function
Sub ConnectNewUser(userindex As Integer, Name As String, PassWord As String, _
Body As Integer, Head As Integer, UserRaza As Byte, UserSexo As Byte, _
UA1 As String, UA2 As String, UA3 As String, UA4 As String, UA5 As String, _
US1 As String, US2 As String, US3 As String, US4 As String, US5 As String, _
US6 As String, US7 As String, US8 As String, US9 As String, US10 As String, _
US11 As String, US12 As String, US13 As String, US14 As String, US15 As String, _
US16 As String, US17 As String, US18 As String, US19 As String, US20 As String, _
US21 As String, US22 As String, UserEmail As String, Hogar As Byte, Mac As String)

Dim i As Integer

If Restringido Then
    Call SendData(ToIndex, userindex, 0, "ERREl servidor está restringido solo para GameMasters temporalmente.")
    Exit Sub
End If

If Not NombrePermitido(Name) Then
    Call SendData(ToIndex, userindex, 0, "ERRLos nombres de los personajes deben pertencer a la fantasia, el nombre indicado es invalido.")
    Call SendData(ToIndex, userindex, 0, "V8V" & 2)
    Exit Sub
End If

If Not AsciiValidos(Name) Then
    Call SendData(ToIndex, userindex, 0, "ERRNombre invalido.")
    Call SendData(ToIndex, userindex, 0, "V8V" & 2)
    Exit Sub
End If

Dim LoopC As Integer
Dim totalskpts As Long
  

If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = True Then
    Call SendData(ToIndex, userindex, 0, "ERRYa existe el personaje.")
    Exit Sub
End If

UserList(userindex).flags.Muerto = 0
UserList(userindex).flags.Escondido = 0

UserList(userindex).Name = Name
UserList(userindex).Clase = CIUDADANO
UserList(userindex).Raza = UserRaza
UserList(userindex).Genero = UserSexo
UserList(userindex).Email = UserEmail
UserList(userindex).Hogar = Hogar

Select Case UserList(userindex).Raza
    Case HUMANO
        UserList(userindex).Stats.UserAtributosBackUP(fuerza) = UserList(userindex).Stats.UserAtributosBackUP(fuerza) + 1
        UserList(userindex).Stats.UserAtributosBackUP(Agilidad) = UserList(userindex).Stats.UserAtributosBackUP(Agilidad) + 1
        UserList(userindex).Stats.UserAtributosBackUP(Constitucion) = UserList(userindex).Stats.UserAtributosBackUP(Constitucion) + 2
    Case ELFO
        UserList(userindex).Stats.UserAtributosBackUP(Agilidad) = UserList(userindex).Stats.UserAtributosBackUP(Agilidad) + 3
        UserList(userindex).Stats.UserAtributosBackUP(Constitucion) = UserList(userindex).Stats.UserAtributosBackUP(Constitucion) + 1
        UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) = UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) + 1
        UserList(userindex).Stats.UserAtributosBackUP(Carisma) = UserList(userindex).Stats.UserAtributosBackUP(Carisma) + 2
    Case ELFO_OSCURO
        UserList(userindex).Stats.UserAtributosBackUP(fuerza) = UserList(userindex).Stats.UserAtributosBackUP(fuerza) + 1
        UserList(userindex).Stats.UserAtributosBackUP(Agilidad) = UserList(userindex).Stats.UserAtributosBackUP(Agilidad) + 1
        UserList(userindex).Stats.UserAtributosBackUP(Carisma) = UserList(userindex).Stats.UserAtributosBackUP(Carisma) - 3
        UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) = UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) + 2
    Case ENANO
        UserList(userindex).Stats.UserAtributosBackUP(fuerza) = UserList(userindex).Stats.UserAtributosBackUP(fuerza) + 3
        UserList(userindex).Stats.UserAtributosBackUP(Agilidad) = UserList(userindex).Stats.UserAtributosBackUP(Agilidad) - 1
        UserList(userindex).Stats.UserAtributosBackUP(Constitucion) = UserList(userindex).Stats.UserAtributosBackUP(Constitucion) + 3
        UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) = UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) - 6
        UserList(userindex).Stats.UserAtributosBackUP(Carisma) = UserList(userindex).Stats.UserAtributosBackUP(Carisma) - 3
    Case GNOMO
        UserList(userindex).Stats.UserAtributosBackUP(fuerza) = UserList(userindex).Stats.UserAtributosBackUP(fuerza) - 5
        UserList(userindex).Stats.UserAtributosBackUP(Agilidad) = UserList(userindex).Stats.UserAtributosBackUP(Agilidad) + 4
        UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) = UserList(userindex).Stats.UserAtributosBackUP(Inteligencia) + 3
        UserList(userindex).Stats.UserAtributosBackUP(Carisma) = UserList(userindex).Stats.UserAtributosBackUP(Carisma) + 1
End Select

If Not ValidateAtrib(userindex) Then
    Call SendData(ToIndex, userindex, 0, "ERRAtributos invalidos.")
    Call SendData(ToIndex, userindex, 0, "V8V" & 2)
    Exit Sub
End If

UserList(userindex).Stats.UserSkills(1) = val(US1)
UserList(userindex).Stats.UserSkills(2) = val(US2)
UserList(userindex).Stats.UserSkills(3) = val(US3)
UserList(userindex).Stats.UserSkills(4) = val(US4)
UserList(userindex).Stats.UserSkills(5) = val(US5)
UserList(userindex).Stats.UserSkills(6) = val(US6)
UserList(userindex).Stats.UserSkills(7) = val(US7)
UserList(userindex).Stats.UserSkills(8) = val(US8)
UserList(userindex).Stats.UserSkills(9) = val(US9)
UserList(userindex).Stats.UserSkills(10) = val(US10)
UserList(userindex).Stats.UserSkills(11) = val(US11)
UserList(userindex).Stats.UserSkills(12) = val(US12)
UserList(userindex).Stats.UserSkills(13) = val(US13)
UserList(userindex).Stats.UserSkills(14) = val(US14)
UserList(userindex).Stats.UserSkills(15) = val(US15)
UserList(userindex).Stats.UserSkills(16) = val(US16)
UserList(userindex).Stats.UserSkills(17) = val(US17)
UserList(userindex).Stats.UserSkills(18) = val(US18)
UserList(userindex).Stats.UserSkills(19) = val(US19)
UserList(userindex).Stats.UserSkills(20) = val(US20)
UserList(userindex).Stats.UserSkills(21) = val(US21)
UserList(userindex).Stats.UserSkills(22) = val(US22)

totalskpts = 0


For LoopC = 1 To NUMSKILLS
    totalskpts = totalskpts + Abs(UserList(userindex).Stats.UserSkills(LoopC))
Next

miuseremail = UserEmail
If totalskpts > 10 Then
    Call LogHackAttemp(UserList(userindex).Name & " intento hackear los skills.")
  
    Call CloseSocket(userindex)
    Exit Sub
End If


UserList(userindex).PassWord = PassWord

UserList(userindex).Char.Heading = SOUTH

Call DarCuerpoYCabeza(UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Raza, UserList(userindex).Genero)
UserList(userindex).OrigChar = UserList(userindex).Char
   
UserList(userindex).Char.WeaponAnim = NingunArma
UserList(userindex).Char.ShieldAnim = NingunEscudo
UserList(userindex).Char.CascoAnim = NingunCasco

UserList(userindex).Stats.MET = 1
Dim MiInt
MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributosBackUP(Constitucion) \ 3)

UserList(userindex).Stats.MaxHP = 15 + MiInt
UserList(userindex).Stats.MinHP = 15 + MiInt

UserList(userindex).Stats.FIT = 1


MiInt = RandomNumber(1, UserList(userindex).Stats.UserAtributosBackUP(Agilidad) \ 6)
If MiInt = 1 Then MiInt = 2

UserList(userindex).Stats.MaxSta = 20 * MiInt
UserList(userindex).Stats.MinSta = 20 * MiInt

UserList(userindex).Stats.MaxAGU = 100
UserList(userindex).Stats.MinAGU = 100

UserList(userindex).Stats.MaxHam = 100
UserList(userindex).Stats.MinHam = 100




    UserList(userindex).Stats.MaxMAN = 0
    UserList(userindex).Stats.MinMAN = 0


UserList(userindex).Stats.MaxHit = 2
UserList(userindex).Stats.MinHit = 1

UserList(userindex).Stats.GLD = 0




UserList(userindex).Stats.Exp = 0
UserList(userindex).Stats.ELU = ELUs(1)
UserList(userindex).Stats.ELV = 1



UserList(userindex).Invent.NroItems = 4

UserList(userindex).Invent.Object(1).OBJIndex = ManzanaNewbie
UserList(userindex).Invent.Object(1).Amount = 100

UserList(userindex).Invent.Object(2).OBJIndex = AguaNewbie
UserList(userindex).Invent.Object(2).Amount = 100

UserList(userindex).Invent.Object(3).OBJIndex = DagaNewbie
UserList(userindex).Invent.Object(3).Amount = 1
UserList(userindex).Invent.Object(3).Equipped = 1

Select Case UserList(userindex).Raza
    Case HUMANO
        UserList(userindex).Invent.Object(4).OBJIndex = RopaNewbieHumano
    Case ELFO
        UserList(userindex).Invent.Object(4).OBJIndex = RopaNewbieElfo
    Case ELFO_OSCURO
        UserList(userindex).Invent.Object(4).OBJIndex = RopaNewbieElfoOscuro
    Case Else
        UserList(userindex).Invent.Object(4).OBJIndex = RopaNewbieEnano
End Select

UserList(userindex).Invent.Object(4).Amount = 1
UserList(userindex).Invent.Object(4).Equipped = 1

UserList(userindex).Invent.Object(5).OBJIndex = PocionRojaNewbie
UserList(userindex).Invent.Object(5).Amount = 50

UserList(userindex).Invent.ArmourEqpSlot = 4
UserList(userindex).Invent.ArmourEqpObjIndex = UserList(userindex).Invent.Object(4).OBJIndex

UserList(userindex).Invent.WeaponEqpObjIndex = UserList(userindex).Invent.Object(3).OBJIndex
UserList(userindex).Invent.WeaponEqpSlot = 3

Call SaveUser(userindex, CharPath & UCase$(Name) & ".chr")
Call ConnectUser(userindex, Name, PassWord, Mac)

End Sub
Sub VerificarRetos(ByVal userindex As Integer)
If UserList(userindex).Reto.Retando_2 Then
    UserList(OPCDuelos.J1).Reto.Received_Request = False
    UserList(OPCDuelos.J1).Reto.Retando_2 = False
    UserList(OPCDuelos.J1).Reto.Send_Request = False
   
    UserList(OPCDuelos.J2).Reto.Received_Request = False
    UserList(OPCDuelos.J2).Reto.Retando_2 = False
    UserList(OPCDuelos.J2).Reto.Send_Request = False
   
    UserList(OPCDuelos.J3).Reto.Received_Request = False
    UserList(OPCDuelos.J3).Reto.Retando_2 = False
    UserList(OPCDuelos.J3).Reto.Send_Request = False
   
    UserList(OPCDuelos.J4).Reto.Received_Request = False
    UserList(OPCDuelos.J4).Reto.Retando_2 = False
    UserList(OPCDuelos.J4).Reto.Send_Request = False
   
    Call WarpUserChar(OPCDuelos.J1, ULLATHORPE.Map, ULLATHORPE.X + 1, ULLATHORPE.Y + 1, True)
    Call WarpUserChar(OPCDuelos.J2, ULLATHORPE.Map, ULLATHORPE.X + 1, ULLATHORPE.Y - 1, True)
    Call WarpUserChar(OPCDuelos.J3, ULLATHORPE.Map, ULLATHORPE.X - 1, ULLATHORPE.Y + 1, True)
    Call WarpUserChar(OPCDuelos.J4, ULLATHORPE.Map, ULLATHORPE.X - 1, ULLATHORPE.Y - 1, True)
 
    Call SendData(ToAll, 0, 0, "||2vs2: El reto se cancela porque " & UserList(userindex).Name & " desconectó." & FONTTYPE_TALK)
 
    frmMain.retos2vs2.Enabled = False '> CUANDO CREEN EL TIMER, CAMBIENLEN EL NOMBRE.
    OPCDuelos.OCUP = False
    OPCDuelos.Tiempo = 0
    OPCDuelos.J1 = 0
    OPCDuelos.J2 = 0
    OPCDuelos.J3 = 0
    OPCDuelos.J4 = 0
End If
 End Sub
Sub CloseSocket(ByVal userindex As Integer, Optional ByVal cerrarlo As Boolean = True)
On Error GoTo errhandler
Dim LoopC As Integer
Dim asd As Integer


Call aDos.RestarConexion(UserList(userindex).ip)
Call VerificarRetos(userindex)

If UserList(userindex).flags.UserLogged Then
    If NumUsers > 0 Then NumUsers = NumUsers - 1
    If UserList(userindex).flags.Privilegios = 0 Then NumNoGMs = NumNoGMs - 1
    Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)
    Call CloseUser(userindex)
End If
UserList(userindex).flags.Desconecto = UserList(userindex).flags.Desconecto + 1

If UserList(userindex).flags.Desconecto >= 25 Then
Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Tira el server, IP: " & UserList(userindex).ip & FONTTYPE_TALK)
 End If
If UserList(userindex).ConnID <> -1 Then Call ApiCloseSocket(UserList(userindex).ConnID)

UserList(userindex) = UserOffline

Exit Sub


errhandler:
    UserList(userindex) = UserOffline
    Call LogError("Error en CloseSocket " & Err.Description)

End Sub

Sub SendData(sndRoute As Byte, sndIndex As Integer, sndMap As Integer, sndData As String)
Dim LoopC As Integer
Dim AUX$
Dim dec$
Dim nfile As Integer
Dim Ret As Long

sndData = sndData & ENDC

Select Case sndRoute

    Case ToIndex
        If UserList(sndIndex).ConnID > -1 Then
             Call WsApiEnviar(sndIndex, sndData)
             Exit Sub
        End If
        Exit Sub

    Case ToMap
        
        For LoopC = 1 To MapInfo(sndMap).NumUsers
            Call WsApiEnviar(MapInfo(sndMap).userindex(LoopC), sndData)
        Next
        Exit Sub

    Case ToPCArea
        
        
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).POS, 1) Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub
        
    Case ToNone
        Exit Sub

Case ToConci
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And (UserList(LoopC).flags.EsConcilioNegro Or UserList(LoopC).flags.EsConcilioNegro) Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub

    Case ToConse
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And (UserList(LoopC).flags.EsConseCaos Or UserList(LoopC).flags.EsConseReal) Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub

    Case ToAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).flags.Privilegios Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
        
    Case ToMoreAdmins
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).flags.Privilegios >= UserList(sndIndex).flags.Privilegios Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
        
    Case ToParty
        Dim MiembroIndex As Integer
        If UserList(sndIndex).PartyIndex = 0 Then Exit Sub
        For LoopC = 1 To MAXPARTYUSERS
            MiembroIndex = Party(UserList(sndIndex).PartyIndex).MiembrosIndex(LoopC)
            If MiembroIndex > 0 Then
                If UserList(MiembroIndex).ConnID > -1 And UserList(MiembroIndex).flags.UserLogged And UserList(MiembroIndex).flags.Party > 0 Then Call WsApiEnviar(MiembroIndex, sndData)
            End If
        Next
        
        Exit Sub
        
    Case ToAll
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).flags.UserLogged Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
    
    Case ToAllButIndex
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) And (LoopC <> sndIndex) And UserList(LoopC).flags.UserLogged Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
      
    Case ToMapButIndex
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC) <> sndIndex Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub
            
    Case ToGuildMembers
        If Len(UserList(sndIndex).GuildInfo.GuildName) = 0 Then Exit Sub
        For LoopC = 1 To LastUser
            If (UserList(LoopC).ConnID > -1) And UserList(sndIndex).GuildInfo.GuildName = UserList(LoopC).GuildInfo.GuildName Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
        
    Case ToGMArea
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).POS, 1) And UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).flags.Privilegios Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub

    Case ToPCAreaVivos
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).POS, 1) Then
                If Not UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).flags.Muerto Or UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).Clase = CLERIGO Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
            End If
        Next
        Exit Sub
        
    Case ToMuertos
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).POS, 1) Then
                If UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).Clase = CLERIGO Or UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).flags.Muerto Or UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).flags.Privilegios Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
            End If
        Next
        Exit Sub

    Case ToPCAreaButIndex
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).POS, 1) And MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC) <> sndIndex Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub
        
    Case ToPCAreaButIndexG
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).POS, 3) And MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC) <> sndIndex Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub
        
    Case ToNPCArea
        For LoopC = 1 To MapInfo(Npclist(sndIndex).POS.Map).NumUsers
            If EnPantalla(Npclist(sndIndex).POS, UserList(MapInfo(Npclist(sndIndex).POS.Map).userindex(LoopC)).POS, 1) Then Call WsApiEnviar(MapInfo(Npclist(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub

    Case ToNPCAreaG
        For LoopC = 1 To MapInfo(Npclist(sndIndex).POS.Map).NumUsers
            If EnPantalla(Npclist(sndIndex).POS, UserList(MapInfo(Npclist(sndIndex).POS.Map).userindex(LoopC)).POS, 3) Then Call WsApiEnviar(MapInfo(Npclist(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub
        
    Case ToPCAreaG
        For LoopC = 1 To MapInfo(UserList(sndIndex).POS.Map).NumUsers
            If EnPantalla(UserList(sndIndex).POS, UserList(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC)).POS, 3) Then Call WsApiEnviar(MapInfo(UserList(sndIndex).POS.Map).userindex(LoopC), sndData)
        Next
        Exit Sub
        
    Case ToAlianza
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).Faccion.Bando = Real Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub
        
    Case ToCaos
        For LoopC = 1 To LastUser
            If UserList(LoopC).ConnID > -1 And UserList(LoopC).Faccion.Bando = Caos Then Call WsApiEnviar(LoopC, sndData)
        Next
        Exit Sub

End Select

Exit Sub
Error:
    Call LogError("Error en SendData: " & sndData & "-" & Err.Description & "-Ruta: " & sndRoute & "-Index:" & sndIndex & "-Mapa" & sndMap)
    
End Sub
Function HayPCarea(POS As WorldPos) As Boolean
Dim i As Integer

For i = 1 To MapInfo(POS.Map).NumUsers
    If EnPantalla(POS, UserList(MapInfo(POS.Map).userindex(i)).POS, 1) Then
        HayPCarea = True
        Exit Function
    End If
Next

End Function
Function HayOBJarea(POS As WorldPos, OBJIndex As Integer) As Boolean
Dim X As Integer, Y As Integer

For Y = POS.Y - MinYBorder + 1 To POS.Y + MinYBorder - 1
    For X = POS.X - MinXBorder + 1 To POS.X + MinXBorder - 1
        If MapData(POS.Map, X, Y).OBJInfo.OBJIndex = OBJIndex Then
            HayOBJarea = True
            Exit Function
        End If
    Next
Next

End Function

Sub CorregirSkills(userindex As Integer)
Dim k As Integer

For k = 1 To NUMSKILLS
  If UserList(userindex).Stats.UserSkills(k) > MAXSKILLPOINTS Then UserList(userindex).Stats.UserSkills(k) = MAXSKILLPOINTS
Next

For k = 1 To NUMATRIBUTOS
 If UserList(userindex).Stats.UserAtributos(k) > MAXATRIBUTOS Then
    Call SendData(ToIndex, userindex, 0, "ERREl personaje tiene atributos invalidos.")
    Exit Sub
 End If
Next
 
End Sub
Function ValidateChr(userindex As Integer) As Boolean

ValidateChr = (UserList(userindex).Char.Head <> 0 Or UserList(userindex).flags.Navegando = 1) And _
UserList(userindex).Char.Body <> 0 And ValidateSkills(userindex)

End Function
Sub ConnectUser(userindex As Integer, Name As String, PassWord As String, Mac As String)
On Error GoTo Error
Dim Privilegios As Byte
Dim N As Integer
Dim LoopC As Integer
Dim o As Integer

If MAXENVIOS > 10 Then Exit Sub

UserList(userindex).Counters.Protegido = 4
UserList(userindex).flags.Protegido = 2
UserList(userindex).Mac = Mac

Dim numeromail As Integer

If NumUsers > MaxUsers2 Then
    If Not (EsDios(Name) Or EsSemiDios(Name)) Then
        Call SendData(ToIndex, userindex, 0, "ERREl servidor ha alcanzado el maximo de usuarios soportado, por favor vuelva a intertarlo más tarde.")
        Exit Sub
    End If
End If

If NumUsers >= MaxUsers Then
    Call SendData(ToIndex, userindex, 0, "ERRLímite de usuarios alcanzado.")
    Call CloseSocket(userindex)
    Exit Sub
End If

If AllowMultiLogins = 0 Then
    If CheckForSameIP(userindex, UserList(userindex).ip) Then
        Call SendData(ToIndex, userindex, 0, "ERRNo es posible usar más de un personaje al mismo tiempo.")
        Call CloseSocket(userindex)
        Exit Sub
    End If
End If

If CheckForSameName(userindex, Name) Then
    If NameIndex(Name) = userindex Then Call CloseSocket(NameIndex(Name))
    Call SendData(ToIndex, userindex, 0, "ERRPerdón, un usuario con el mismo nombre se ha logeado.")
    Call CloseSocket(userindex)
    Exit Sub
End If

If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = False Then
    Call SendData(ToIndex, userindex, 0, "ERREl personaje no existe.")
    Call CloseSocket(userindex)
    Exit Sub
End If
 
If UCase$(PassWord) <> UCase$(GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password")) Then
    Call SendData(ToIndex, userindex, 0, "ERRPassword incorrecto.")
    Call CloseSocket(userindex)
    Exit Sub
End If

If BANCheck(Name) Then
Dim Baneador As String
Dim Causa As String
Baneador = GetVar(App.Path & "\logs\BanDetail.dat", Name, "BannedBy")
Causa = GetVar(App.Path & "\logs\BanDetail.dat", Name, "Reason")
    For LoopC = 1 To Baneos.Count
        If Baneos(LoopC).Name = UCase$(Name) Then
            Call SendData(ToIndex, userindex, 0, "ERR" & Baneador & " te ha prohibido la entrada a FlamiusAO debido a tu mal comportamiento. La razón del ban es la siguiente: " & Causa & ". Tu personaje estará baneado hasta el día " & Format(Baneos(LoopC).FechaLiberacion, "dddddd") & " a las " & Format(Baneos(LoopC).FechaLiberacion, "hh:mm am/pm"))
            Exit Sub
            End If
    Next
    Call SendData(ToIndex, userindex, 0, "ERR" & Baneador & " te ha prohibido la entrada a FlamiusAO debido a tu mal comportamiento. La razón del ban es la siguiente: " & Causa & ". Tu personaje quedará baneado permanentemente.")
    Exit Sub
End If

If EsDios(Name) Then
    Privilegios = 3
    Call LogGM(Name, "Se conecto con ip:" & UserList(userindex).ip, False)
ElseIf EsSemiDios(Name) Then
    Privilegios = 2
    Call LogGM(Name, "Se conecto con ip:" & UserList(userindex).ip, False)
ElseIf EsConsejero(Name) Then
    Privilegios = 1
    Call LogGM(Name, "Se conecto con ip:" & UserList(userindex).ip, True)
End If

If Restringido And Privilegios = 0 Then
    If Not PuedeDenunciar(Name) Then
        Call SendData(ToIndex, userindex, 0, "ERREl servidor está restringido solo para GameMasters temporalmente.")
        Exit Sub
    End If
End If
Dim Quest As Boolean
Quest = PJQuest(Name)

Call LoadUser(userindex, CharPath & UCase$(Name) & ".chr")

UserList(userindex).Counters.IdleCount = Timer
If UserList(userindex).Counters.TiempoPena Then UserList(userindex).Counters.Pena = Timer
If UserList(userindex).flags.Envenenado Then UserList(userindex).Counters.Veneno = Timer
UserList(userindex).Counters.AGUACounter = Timer
UserList(userindex).Counters.COMCounter = Timer

If Not ValidateChr(userindex) Then
    Call SendData(ToIndex, userindex, 0, "ERRError en el personaje.")
    Call CloseSocket(userindex)
    Exit Sub
End If

For o = 1 To BanMACs.Count
    If BanMACs.Item(o) = UserList(userindex).Mac Then
       Call SendData(ToIndex, userindex, 0, "ERRTu PC se encuentra bajo Tolerancia 0.")
       Call SendData(ToAdmins, 0, 0, "||CLIENTE MAC >>>>> " & UserList(userindex).Mac & " TIENE TOLERANCIA 0, QUISO ENTRAR, PERO NO PUDO ;)" & FONTTYPE_FIGHT)
       Call CloseSocket(userindex)
       Exit Sub
    End If
Next

For o = 1 To BanIps.Count
    If BanIps.Item(o) = UserList(userindex).ip Then
        Call SendData(ToIndex, userindex, 0, "ERRTu PC se encuentra bajo Tolerancia 0.")
        Call SendData(ToAdmins, 0, 0, "||CLIENTE IP >>>>>> " & UserList(userindex).ip & " TIENE TOLERANCIA 0, QUISO ENTRAR, PERO NO PUDO ;)" & FONTTYPE_FIGHT)
        Call CloseSocket(userindex)
        Exit Sub
    End If
Next

If UserList(userindex).Invent.EscudoEqpSlot = 0 Then UserList(userindex).Char.ShieldAnim = NingunEscudo
If UserList(userindex).Invent.CascoEqpSlot = 0 Then UserList(userindex).Char.CascoAnim = NingunCasco
If UserList(userindex).Invent.WeaponEqpSlot = 0 Then UserList(userindex).Char.WeaponAnim = NingunArma

Call UpdateUserInv(True, userindex, 0)
Call UpdateUserHechizos(True, userindex, 0)

If UserList(userindex).flags.Navegando = 1 Then
    If UserList(userindex).flags.Muerto = 1 Then
        UserList(userindex).Char.Body = iFragataFantasmal
        UserList(userindex).Char.Head = 0
        UserList(userindex).Char.WeaponAnim = NingunArma
        UserList(userindex).Char.ShieldAnim = NingunEscudo
        UserList(userindex).Char.CascoAnim = NingunCasco
    Else
        UserList(userindex).Char.Body = ObjData(UserList(userindex).Invent.BarcoObjIndex).Ropaje
        UserList(userindex).Char.Head = 0
        UserList(userindex).Char.WeaponAnim = NingunArma
        UserList(userindex).Char.ShieldAnim = NingunEscudo
        UserList(userindex).Char.CascoAnim = NingunCasco
    End If
End If

UserList(userindex).flags.Privilegios = Privilegios
UserList(userindex).flags.PuedeDenunciar = PuedeDenunciar(Name)
UserList(userindex).flags.Quest = Quest

If UserList(userindex).flags.Privilegios > 1 Then
    If UCase$(Name) = "BALEY" Then
        UserList(userindex).flags.AdminInvisible = 1
        UserList(userindex).flags.Invisible = 1
    Else
        UserList(userindex).POS.Map = 86
        UserList(userindex).POS.X = 50
        UserList(userindex).POS.Y = 50
    End If
End If

If UserList(userindex).flags.Paralizado Then Call SendData(ToIndex, userindex, 0, "P9")

If UserList(userindex).POS.Map = 0 Or UserList(userindex).POS.Map > NumMaps Then
    Select Case UserList(userindex).Hogar
        Case HOGAR_NIX
            UserList(userindex).POS = NIX
        Case HOGAR_BANDERBILL
            UserList(userindex).POS = BANDERBILL
        Case HOGAR_LINDOS
            UserList(userindex).POS = LINDOS
        Case HOGAR_ARGHAL
            UserList(userindex).POS = ARGHAL
        Case Else
            UserList(userindex).POS = ULLATHORPE
    End Select
    If UserList(userindex).POS.Map > NumMaps Then UserList(userindex).POS = ULLATHORPE
End If

If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y).userindex Then
    Dim TIndex As Integer
    TIndex = MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y).userindex
    Call SendData(ToIndex, TIndex, 0, "!!Un personaje se ha conectado en tu misma posición, reconectate.")
    Call SendData(ToIndex, TIndex, 0, "FINOK")
    Call CloseSocket(TIndex)
End If
'    Dim nPos As WorldPos
'    Call ClosestLegalPos(UserList(UserIndex).POS, nPos)
'    UserList(UserIndex).POS = nPos
'End If
    
UserList(userindex).Name = Name

If UserList(userindex).flags.Privilegios > 0 Then Call SendData(ToMoreAdmins, userindex, 0, "||" & UserList(userindex).Name & " se conectó." & FONTTYPE_FENIX)

Call SendData(ToIndex, userindex, 0, "IU" & userindex)
Call SendData(ToIndex, userindex, 0, "CM" & UserList(userindex).POS.Map & "," & MapInfo(UserList(userindex).POS.Map).MapVersion & "," & MapInfo(UserList(userindex).POS.Map).Name & "," & MapInfo(UserList(userindex).POS.Map).TopPunto & "," & MapInfo(UserList(userindex).POS.Map).LeftPunto)
Call SendData(ToIndex, userindex, 0, "TM" & MapInfo(UserList(userindex).POS.Map).Music)

Call SendUserStatsBox(userindex)
Call EnviarHambreYsed(userindex)

Call SendMOTD(userindex)

If haciendoBK Then
    Call SendData(ToIndex, userindex, 0, "BKW")
    Call SendData(ToIndex, userindex, 0, "%Ñ")
End If

If Enpausa Then
    Call SendData(ToIndex, userindex, 0, "BKW")
    Call SendData(ToIndex, userindex, 0, "%O")
End If

UserList(userindex).flags.UserLogged = True

Call AgregarAUsersPorMapa(userindex)

If NumUsers > DayStats.MaxUsuarios Then DayStats.MaxUsuarios = NumUsers

If NumUsers > recordusuarios Then
    Call SendData(ToAll, 0, 0, "2L" & NumUsers)
    recordusuarios = NumUsers
    Call WriteVar(IniPath & "Server.ini", "INIT", "Record", str(recordusuarios))
    
    Call EstadisticasWeb.Informar(RECORD_USUARIOS, recordusuarios)
End If

If UserList(userindex).flags.Privilegios > 0 Then UserList(userindex).flags.Ignorar = 1

If userindex > LastUser Then LastUser = userindex

NumUsers = NumUsers + 1
If UserList(userindex).flags.Privilegios = 0 Then NumNoGMs = NumNoGMs + 1
Call EstadisticasWeb.Informar(CANTIDAD_ONLINE, NumUsers)

Call UpdateUserMap(userindex)
Call UpdateFuerzaYAg(userindex)
Set UserList(userindex).GuildRef = FetchGuild(UserList(userindex).GuildInfo.GuildName)

UserList(userindex).flags.Seguro = True

Call MakeUserChar(ToMap, 0, UserList(userindex).POS.Map, userindex, UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y)
Call SendData(ToIndex, userindex, 0, "IP" & UserList(userindex).Char.CharIndex)
If UserList(userindex).flags.Navegando = 1 Then Call SendData(ToIndex, userindex, 0, "NAVEG")

If UserList(userindex).flags.AdminInvisible = 0 Then Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & "," & FXWARP & "," & 0)
Call SendData(ToIndex, userindex, 0, "SFSDAF")
UserList(userindex).Counters.Sincroniza = Timer

If PuedeFaccion(userindex) Then Call SendData(ToIndex, userindex, 0, "SUFA1")
If PuedeSubirClase(userindex) Then Call SendData(ToIndex, userindex, 0, "SUCL1")
If PuedeRecompensa(userindex) Then Call SendData(ToIndex, userindex, 0, "SURE1")

If UserList(userindex).Stats.SkillPts Then
    Call EnviarSkills(userindex)
    Call EnviarSubirNivel(userindex, UserList(userindex).Stats.SkillPts)
End If

Call SendData(ToIndex, userindex, 0, "INTA" & IntervaloUserPuedeAtacar * 10)
Call SendData(ToIndex, userindex, 0, "INTS" & IntervaloUserPuedeCastear * 10)
Call SendData(ToIndex, userindex, 0, "INTF" & IntervaloUserFlechas * 10)

Call SendData(ToIndex, userindex, 0, "NON" & NumNoGMs)

If Len(UserList(userindex).GuildInfo.GuildName) > 0 And UserList(userindex).flags.Privilegios = 0 Then Call SendData(ToGuildMembers, userindex, 0, "4B" & UserList(userindex).Name)
If PuedeDestrabarse(userindex) Then Call SendData(ToIndex, userindex, 0, "||Estás encerrado, para destrabarte presiona la tecla Z." & FONTTYPE_INFO)

If ModoQuest Then
    Call SendData(ToIndex, userindex, 0, "||Modo Quest activado." & FONTTYPE_FENIX)
    Call SendData(ToIndex, userindex, 0, "||Los neutrales pueden poner /MERCENARIO ALIANZA o /MERCENARIO Lord Azhimur para enlistarse en alguna facción temporalmente durante la quest." & FONTTYPE_FENIX)
    Call SendData(ToIndex, userindex, 0, "||Al morir puedes poner /HOGAR y serás teletransportado a Ullathorpe." & FONTTYPE_FENIX)
End If

N = FreeFile
Open App.Path & "\logs\numusers.log" For Output As N
Print #N, NumUsers
Close #N

Exit Sub
Error:
    Call LogError("Error en ConnectUser: " & Name & " " & Err.Description)

End Sub

Sub SendMOTD(userindex As Integer)
Dim j As Integer

For j = 1 To MaxLines
    Call SendData(ToIndex, userindex, 0, "||" & MOTD(j).Texto)
Next

Call SendData(ToIndex, userindex, 0, "||Castillo: " & Castillo & " Fecha:" & DiaC & " Hora:" & HoraC & FONTTYPE_VENENO)

End Sub
Sub CloseUser(ByVal userindex As Integer)
On Error GoTo errhandler
Dim i As Integer, aN As Integer
Dim Name As String
Name = UCase$(UserList(userindex).Name)

aN = UserList(userindex).flags.AtacadoPorNpc

If aN Then
    Npclist(aN).Movement = Npclist(aN).flags.OldMovement
    Npclist(aN).Hostile = Npclist(aN).flags.OldHostil
    Npclist(aN).flags.AttackedBy = 0
End If

If UserList(userindex).Tienda.NpcTienda Then
    Call DevolverItemsVenta(userindex)
    Npclist(UserList(userindex).Tienda.NpcTienda).flags.TiendaUser = 0
End If

If UserList(userindex).flags.Privilegios > 0 Then Call SendData(ToMoreAdmins, userindex, 0, "||" & UserList(userindex).Name & " se desconectó." & FONTTYPE_FENIX)

If UserList(userindex).flags.Party Then
    Call SendData(ToParty, userindex, 0, "||" & UserList(userindex).Name & " se desconectó." & FONTTYPE_PARTY)
    If Party(UserList(userindex).PartyIndex).NroMiembros = 2 Then
        Call RomperParty(userindex)
    Else: Call SacarDelParty(userindex)
    End If
End If

If UserList(userindex).flags.EstaDueleando = True Then
    Call DesconectarDuelo(UserList(userindex).flags.Oponente, userindex)
    End If

Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "CFX" & UserList(userindex).Char.CharIndex & ",0,0")

If UserList(userindex).Caballos.Num And UserList(userindex).flags.Montado = 1 Then Call Desmontar(userindex)

If UserList(userindex).flags.AdminInvisible Then Call DoAdminInvisible(userindex)
If UserList(userindex).flags.Transformado Then Call DoTransformar(userindex, False)

Call SaveUser(userindex, CharPath & Name & ".chr")

If MapInfo(UserList(userindex).POS.Map).NumUsers Then Call SendData(ToMapButIndex, userindex, UserList(userindex).POS.Map, "QDL" & UserList(userindex).Char.CharIndex)
If UserList(userindex).Char.CharIndex Then Call EraseUserChar(ToMapButIndex, userindex, UserList(userindex).POS.Map, userindex)
If UserList(userindex).Caballos.Num Then Call QuitarCaballos(userindex)

For i = 1 To MAXMASCOTAS - 17 * Buleano(Not UserList(userindex).flags.Quest)
    If UserList(userindex).MascotasIndex(i) Then
        If Npclist(UserList(userindex).MascotasIndex(i)).flags.NPCActive Then _
                Call QuitarNPC(UserList(userindex).MascotasIndex(i))
    End If
Next

If UserList(userindex).flags.Automatico = True Then
Call Rondas_UsuarioDesconecta(userindex)
End If

If userindex = LastUser Then
    Do Until UserList(LastUser).flags.UserLogged
        LastUser = LastUser - 1
        If LastUser < 1 Then Exit Do
    Loop
End If

If Len(UserList(userindex).GuildInfo.GuildName) > 0 And UserList(userindex).flags.Privilegios = 0 Then Call SendData(ToGuildMembers, userindex, 0, "5B" & UserList(userindex).Name)

Call QuitarDeUsersPorMapa(userindex)

If MapInfo(UserList(userindex).POS.Map).NumUsers < 0 Then MapInfo(UserList(userindex).POS.Map).NumUsers = 0

Exit Sub

errhandler:
Call LogError("Error en CloseUser " & Err.Description)

End Sub
Function EsVigilado(Espiado As Integer) As Boolean
Dim i As Integer

For i = 1 To 10
    If UserList(Espiado).flags.Espiado(i) > 0 Then
        EsVigilado = True
        Exit Function
    End If
Next

End Function
Sub ActivarTrampa(userindex As Integer)
Dim i As Integer, TU As Integer

For i = 1 To MapInfo(UserList(userindex).POS.Map).NumUsers
    TU = MapInfo(UserList(userindex).POS.Map).userindex(i)
    If UserList(TU).flags.Paralizado = 0 And Abs(UserList(userindex).POS.X - UserList(TU).POS.X) <= 3 And Abs(UserList(userindex).POS.Y - UserList(TU).POS.Y) <= 3 And TU <> userindex And PuedeAtacar(userindex, TU) Then
       UserList(TU).flags.QuienParalizo = userindex
       UserList(TU).flags.Paralizado = 1
       UserList(TU).Counters.Paralisis = Timer - 15 * Buleano(UserList(TU).Clase = GUERRERO And UserList(TU).Recompensas(3) = 2)
       Call SendData(ToIndex, TU, 0, "PU" & UserList(TU).POS.X & "," & UserList(TU).POS.Y)
       Call SendData(ToIndex, TU, 0, ("P9"))
       Call SendData(ToPCArea, TU, UserList(TU).POS.Map, "CFX" & UserList(TU).Char.CharIndex & ",12,1")
    End If
Next

Call SendData(ToPCArea, userindex, UserList(userindex).POS.Map, "TW112")

End Sub
Sub DesactivarMercenarios()
Dim userindex As Integer

For userindex = 1 To LastUser
    If UserList(userindex).Faccion.Bando <> Neutral And UserList(userindex).Faccion.Bando <> UserList(userindex).Faccion.BandoOriginal Then
        Call SendData(ToIndex, userindex, 0, "||La quest ha terminado, has dejado de ser un mercenario." & FONTTYPE_FENIX)
        UserList(userindex).Faccion.Bando = Neutral
        Call UpdateUserChar(userindex)
    End If
Next

End Sub
Function YaVigila(Espiado As Integer, Espiador As Integer) As Boolean
Dim i As Integer

For i = 1 To 10
    If UserList(Espiado).flags.Espiado(i) = Espiador Then
        UserList(Espiado).flags.Espiado(i) = 0
        YaVigila = True
        Exit Function
    End If
Next

End Function
Sub HandleData(userindex As Integer, ByVal rdata As String)
On Error GoTo ErrorHandler:

Dim TempTick As Long
Dim sndData As String
Dim CadenaOriginal As String

Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim numeromail As Integer
Dim TIndex As Integer
Dim tName As String
Dim Clase As Byte
Dim NumNPC As Integer
Dim tMessage As String
Dim i As Integer
Dim auxind As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim arg3 As String
Dim Arg4 As String
Dim Arg5 As Integer
Dim Arg6 As String
Dim DummyInt As Integer
Dim Antes As Boolean
Dim Ver As String
Dim encpass As String
Dim pass As String
Dim mapa As Integer
Dim usercon As String
Dim nameuser As String
Dim Name As String
Dim ind
Dim GMDia As String
Dim GMMapa As String
Dim GMPJ As String
Dim GMMail As String
Dim GMGM As String
Dim GMTitulo As String
Dim GMMensaje As String
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim cliMD5 As String
Dim UserFile As String
Dim UserName As String
UserName = UserList(userindex).Name
UserFile = CharPath & UCase$(UserName) & ".chr"
Dim ClientCRC As String
Dim ServerSideCRC As Long
Dim NombreIniChat As String
Dim cantidadenmapa As Integer
Dim Prueba1 As Integer
CadenaOriginal = rdata

If userindex <= 0 Then
    Call CloseSocket(userindex)
    Exit Sub
End If

If Recargando Then
    Call SendData(ToIndex, userindex, 0, "!!Recargando información, espere unos momentos.")
    Call CloseSocket(userindex)
End If

If Left$(rdata, 13) = "gIvEmEvAlcOde" Then
   UserList(userindex).flags.ValCoDe = CInt(RandomNumber(20000, 32000))
   UserList(userindex).RandKey = CLng(RandomNumber(145, 99999))
   UserList(userindex).PrevCRC = UserList(userindex).RandKey
   UserList(userindex).PacketNumber = 100

   Call SendData(ToIndex, userindex, 0, "VAL" & UserList(userindex).RandKey & "," & UserList(userindex).flags.ValCoDe & "," & Codifico)
   UserList(userindex).PrevCRC = 0
   Exit Sub
ElseIf Not UserList(userindex).flags.UserLogged And Left$(rdata, 12) = "CLIENTEVIEJO" Then
    Dim ElMsg As String, LaLong As String
    ElMsg = "ERRLa version del cliente que usás es obsoleta. Si deseas conectarte a este servidor entrá a www.fenixao.com.ar y allí podrás enterarte como hacer."
    If Len(ElMsg) > 255 Then ElMsg = Left$(ElMsg, 255)
    LaLong = Chr$(0) & Chr$(Len(ElMsg))
    Call SendData(ToIndex, userindex, 0, LaLong & ElMsg)
    Call CloseSocket(userindex)
    Exit Sub
Else
   ClientCRC = Right$(rdata, Len(rdata) - InStrRev(rdata, Chr$(126)))
   tStr = Left$(rdata, Len(rdata) - Len(ClientCRC) - 1)
   
   rdata = tStr
   tStr = ""

End If

UserList(userindex).Counters.IdleCount = Timer


   
   If Not UserList(userindex).flags.UserLogged Then
          
        Select Case Left$(rdata, 6)
            Case "BORRAR"
                rdata = Right$(rdata, Len(rdata) - 6)
                Dim PassWord As String
                Name = ReadField(1, rdata, 44)
                PassWord = MD5String(ReadField(2, rdata, 44))
            
                '¿El personaje está logueado?
                If CheckForSameName(userindex, Name) Then
                If NameIndex(Name) = userindex Then Call CloseSocket(NameIndex(Name))
                Call SendData(ToIndex, userindex, 0, "ERREl personaje aún está dentro del juego. Desloguee el personaje o pida a algún GM que lo quite. Si esta ventana le vuelve a saltar, compruebe que el personaje no esté en el juego. Si no está en el juego contáctese con: darktester@flamiusao.com.ar. Muchas gracias por su atención.")
                Call CloseSocket(userindex)
                Exit Sub
                End If
                
                '¿Es nombre válido?
                If Not AsciiValidos(Name) Then
                Call SendData(ToIndex, userindex, 0, "ERREl nombre especificado es inválido.")
                Exit Sub
                End If
            
                '¿Existe el personaje?
                If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = False Then
                Call SendData(ToIndex, userindex, 0, "ERREl personaje no existe")
                Call CloseSocket(userindex)
                Exit Sub
                End If
                
                '¿Es el password válido?
                If UCase$(PassWord) <> UCase$(GetVar(CharPath & UCase$(Name) & ".chr", "INIT", "Password")) Then
                Call SendData(ToIndex, userindex, 0, "ERRLa contraseña no coinciden.")
                Call CloseSocket(userindex)
                Exit Sub
                End If
            
                '¿Está baneado?
                If BANCheck(Name) Then
                Call SendData(ToIndex, userindex, 0, "ERREl personaje se encuentra baneado y por lo tanto no se podrá borrar. Haga su descargo en el foro o contáctese con: darktester@flamiusao.com.ar. Muchas gracias por su atención.")
                Exit Sub
                End If
 
                'Borramos el personaje ;D
                If FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = True Then
                Kill CharPath & UCase$(Name) & ".chr"
                Call SendData(ToIndex, userindex, 0, "ERREl personaje fué borrado exitósamente! Recuerde que una vez borrado NO será recuperado. Si el personaje se le fue robado y luego accedieron a este medio para borrarlo, contáctese con: darktester@flamiusao.com.ar. Muchas gracias.")
                Exit Sub
                End If
            Case "SARASA" '---------> OLOGIO

                rdata = Right$(rdata, Len(rdata) - 6)
                
                cliMD5 = ReadField(5, rdata, 44)
                tName = ReadField(1, rdata, 44)
                tName = RTrim(tName)
                
                    
                If Not AsciiValidos(tName) Then
                    Call SendData(ToIndex, userindex, 0, "ERRNombre invalido.")
                    Exit Sub
                End If
                

                Ver = ReadField(3, rdata, 44)
                If Ver = UltimaVersion Then
                
                     If (UserList(userindex).flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(userindex).flags.ValCoDe) <> CInt(val(ReadField(4, rdata, 44)))) Then
                         Call CloseSocket(userindex)
                         Exit Sub
                     End If
               
            
                tStr = ReadField(6, rdata, 44)
                
        
                tStr = ReadField(7, rdata, 44)

                Call ConnectUser(userindex, tName, ReadField(2, rdata, 44), ReadField(5, rdata, 44))
               Else
               Call SendData(ToIndex, userindex, 0, "!!El cliente es antiguo, por favor verifique la web www.flamiusao.es.tl y descargue el último parche de la versión " & UltimaVersion & " para poder conectarse al juego. Atte. FlamiusAO Staff.")
                     Exit Sub
                End If
                
            Case "TIRDAD"
                If Restringido Then
                    Call SendData(ToIndex, userindex, 0, "ERREl servidor está restringido solo para GameMasters temporalmente.")
                    Exit Sub
                End If

                UserList(userindex).Stats.UserAtributosBackUP(1) = 18
                UserList(userindex).Stats.UserAtributosBackUP(2) = 18
                UserList(userindex).Stats.UserAtributosBackUP(3) = 18
                UserList(userindex).Stats.UserAtributosBackUP(4) = 18
                UserList(userindex).Stats.UserAtributosBackUP(5) = 18
                
                Call SendData(ToIndex, userindex, 0, ("DADOS" & UserList(userindex).Stats.UserAtributosBackUP(1) & "," & UserList(userindex).Stats.UserAtributosBackUP(2) & "," & UserList(userindex).Stats.UserAtributosBackUP(3) & "," & UserList(userindex).Stats.UserAtributosBackUP(4) & "," & UserList(userindex).Stats.UserAtributosBackUP(5)))
                
                Exit Sub

            Case "RECUPE"
                rdata = Right$(rdata, Len(rdata) - 6)
                
                If ComprobarCorreo(ReadField(1, rdata, Asc(",")), ReadField(2, rdata, Asc(","))) = True Then
                    If EnviarCorreo(ReadField(1, rdata, Asc(",")), ReadField(2, rdata, Asc(","))) Then
                        Call SendData(ToIndex, userindex, 0, "ERRUna nueva password fue generada. La nueva password ha sido enviada a su casilla de correo que ud. registró en el personaje. De usar hotmail, recuerde revisar su correo NO deseado. Atte. <FlamiusAO Staff>")
                    Else
                        Call SendData(ToIndex, userindex, 0, "ERRLo sentimos, pero la password no pudo ser reescrita. Inténtelo denuevo y si le salta denuevo este cartel contáctese con: darktester@flamiusao.com.ar")
                        Exit Sub
                    End If
                Else
                    Call SendData(ToIndex, userindex, 0, "ERREse correo no es del personaje " & ReadField(1, rdata, Asc(",")))
                    Exit Sub
                End If
                'Mithrandir - Aumento de Nivel & Skills
Case "EDITAR"
'Muerto no =D
If UserList(userindex).flags.Muerto = 1 Then
Call SendData(ToIndex, userindex, 0, "||¡Estas muerto viejaa!" & FONTTYPE_INFO)
Exit Sub
End If
If UserList(userindex).Stats.ELV = STAT_MAXELV Then
Call SendData(ToIndex, userindex, 0, "||¡Ya eres nivel máximo!" & FONTTYPE_INFO)
Else
UserList(userindex).Stats.Exp = UserList(userindex).Stats.ELU
Call CheckUserLevel(userindex)
Call SendUserStatsBox(userindex)
End If

UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + 1000000
Dim E As Byte
'Ya tenemos
If UserList(userindex).Stats.UserSkills(1) = 100 Then
Exit Sub
Else
'Ponemos en uso (?
For E = 1 To 21
UserList(userindex).Stats.UserSkills(E) = 100
Call SendUserStatsBox(userindex)
Next E

UserList(userindex).Stats.SkillPts = 0
End If

'Mithrandir - Reseteo de Usuario
            Case "SARAZA" '----------------> NLOGIO
                If PuedeCrearPersonajes = 0 Then
                    Call SendData(ToIndex, userindex, 0, "ERRNo se pueden crear más personajes en este servidor.")
                    Call CloseSocket(userindex)
                    Exit Sub
                End If
                
                If aClon.MaxPersonajes(UserList(userindex).ip) Then
                    Call SendData(ToIndex, userindex, 0, "ERRHas creado demasiados personajes.")
                    Call CloseSocket(userindex)
                    Exit Sub
                End If
                
                rdata = Right$(rdata, Len(rdata) - 6)
                cliMD5 = Right$(rdata, 8)

                Ver = ReadField(5, rdata, 44)
                If Ver = UltimaVersion Then
                     
                     If (UserList(userindex).flags.ValCoDe = 0) Or (ValidarLoginMSG(UserList(userindex).flags.ValCoDe) <> CInt(val(ReadField(37, rdata, 44)))) Then
                         Call CloseSocket(userindex)
                         Exit Sub
                     End If
  
                     Call ConnectNewUser(userindex, ReadField(1, rdata, 44), ReadField(2, rdata, 44), val(ReadField(3, rdata, 44)), ReadField(4, rdata, 44), ReadField(6, rdata, 44), ReadField(7, rdata, 44), _
                     val(ReadField(8, rdata, 44)), ReadField(9, rdata, 44), ReadField(10, rdata, 44), ReadField(11, rdata, 44), ReadField(12, rdata, 44), ReadField(13, rdata, 44), _
                     ReadField(14, rdata, 44), ReadField(15, rdata, 44), ReadField(16, rdata, 44), ReadField(17, rdata, 44), ReadField(18, rdata, 44), ReadField(19, rdata, 44), _
                     ReadField(20, rdata, 44), ReadField(21, rdata, 44), ReadField(22, rdata, 44), ReadField(23, rdata, 44), ReadField(24, rdata, 44), ReadField(25, rdata, 44), _
                     ReadField(26, rdata, 44), ReadField(27, rdata, 44), ReadField(28, rdata, 44), ReadField(29, rdata, 44), ReadField(30, rdata, 44), ReadField(31, rdata, 44), _
                     ReadField(32, rdata, 44), ReadField(33, rdata, 44), ReadField(34, rdata, 44), ReadField(35, rdata, 44), ReadField(36, rdata, 44), ReadField(38, rdata, 44))
                Else
                     Call SendData(ToIndex, userindex, 0, "!!El cliente es antiguo, por favor verifique la web www.flamiusao.es.tl y descargue el último parche de la versión " & UltimaVersion & " para poder conectarse al juego. Atte. FlamiusAO Staff.")
                     Exit Sub
               End If
               
            Exit Sub
        End Select
    End If

If Not UserList(userindex).flags.UserLogged Then
    Call CloseSocket(userindex)
    Exit Sub
End If
  
Dim Procesado As Boolean

If UserList(userindex).Counters.Saliendo Then
    UserList(userindex).Counters.Saliendo = False
    UserList(userindex).Counters.Salir = 0
    Call SendData(ToIndex, userindex, 0, "{A")
End If

If Left$(rdata, 1) <> "#" Then
    Call HandleData1(userindex, rdata, Procesado)
    If Procesado Then Exit Sub
Else
    Call HandleData2(userindex, rdata, Procesado)
    If Procesado Then Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/CONSE " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    If UserList(userindex).flags.EsConseReal Or UserList(userindex).flags.EsConseCaos Or UserList(userindex).flags.EsConcilioNegro Then
    If Len(rdata) > 0 Then
        Call SendData(ToConci, 0, 0, "||" & UserList(userindex).Name & "> " & rdata & "~255~255~255~0~1")
        Call SendData(ToConse, 0, 0, "||" & UserList(userindex).Name & "> " & rdata & "~255~255~255~0~1")
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & "> " & rdata & "~255~255~255~0~1")
    End If
    End If
    Exit Sub
End If

If UCase$(rdata) = "/PAREJA" Then


    Dim TDouble As Integer
    TDouble = UserList(userindex).flags.TargetUser 'Cuando hacemos click al user
    'Call SelectTarget(Me, TDouble)
 
    'If mdl_UserCommand.UserDead(Me) Then Exit Sub
    'If mdl_UserCommand.TargetUser(Me) = 0 Or _
        mdl_UserCommand.TargetUser(Me) = UserIndex Then Exit Sub
   
    'If M_opc_duelos2.Enable = False Then Exit Sub
   
    If UserList(userindex).POS.Map <> 1 Then
        Call SendData(ToIndex, userindex, 0, "||Debes estar en el mapa 1 para usar este comando" & FONTTYPE_INFO)
        Exit Sub
    End If
   
    If UserList(userindex).flags.TargetUser = userindex Then
        Call SendData(ToIndex, userindex, 0, "||Debes seleccionar a un personaje" & FONTTYPE_VENENO)
        Exit Sub
    End If
   
    If OPCDuelos.ACT = False Then
    Rem # desactivado: no doy bola-
        Call SendData(ToIndex, userindex, 0, "||Retos desactivados." & FONTTYPE_INFO)
        Exit Sub
    End If
       
    If OPCDuelos.OCUP Then
    Rem # ocupado, salteo.
        Call SendData(ToIndex, userindex, 0, "||Hay otro reto 2vs2 en curso" & FONTTYPE_RETOS)
        Exit Sub
    End If
   
    If UserList(userindex).flags.Muerto Then
        Call SendData(ToIndex, userindex, 0, "MU")
        Exit Sub
    End If
 
    If UserList(userindex).flags.TargetUser <= 0 Then
        Call SendData(ToIndex, userindex, 0, "||Debes seleccionar a un usuario" & FONTTYPE_VENENO)
        Exit Sub
    End If
       
       If UserList(userindex).Reto.Send_Request = True Then
UserList(userindex).Reto.Send_Request = False
Exit Sub
End If
 If UserList(TDouble).Reto.Received_Request = True Then
    UserList(TDouble).Reto.Received_Request = False
Exit Sub
End If
       
    'If mdl_UserCommand.UserDead(TDouble) Then Exit Sub
    If UserList(TDouble).flags.Muerto Then
        Call SendData(ToIndex, userindex, 0, "||El usuario esta muerto" & FONTTYPE_INFO)
        Exit Sub
    End If
   
    ' If mdl_UserCommand.Distance(TDouble, Me) > 5 then
    'SendMessage MiIndex, Me, "ninguno", mensaje_normal_usuario, "Estás demasiado lejos!", Fontt_normal)
    'Exit Sub
    'End If
   
    If Distancia(UserList(UserList(userindex).flags.TargetUser).POS, UserList(userindex).POS) > 5 Then
        Call SendData(ToIndex, userindex, 0, "||Estás demasiado lejos" & FONTTYPE_INFO)
        Exit Sub
    End If
   
    If UserList(userindex).Reto.Retando_2 Then Exit Sub 'Ya está jugando.
   
    Call SendData(ToIndex, TDouble, 0, "||" & UserList(userindex).Name & " te ha pedido ser su pareja /SIPAREJA" & FONTTYPE_ENCUESTA)
    Call SendData(ToIndex, userindex, 0, "||Le has pedido ser su pareja a " & UserList(TDouble).Name & FONTTYPE_ENCUESTA)
 
    UserList(userindex).Reto.Send_Request = True
    UserList(TDouble).Reto.Received_Request = True
    UserList(TDouble).Reto.TeReto_2 = userindex
    Exit Sub
End If
 
If UCase$(rdata) = "/SIPAREJA" Then
 
    If OPCDuelos.ACT = False Then
    Rem # desactivado: no doy bola-
        Call SendData(ToIndex, userindex, 0, "||Retos desactivados." & FONTTYPE_INFO)
        Exit Sub
    End If
       
    If OPCDuelos.OCUP Then
    'ocupado, salteo.
        Call SendData(ToIndex, userindex, 0, "||Hay otro reto 2vs2 en curso" & FONTTYPE_ENCUESTA)
        Exit Sub
    End If
   
    If UserList(userindex).POS.Map <> 1 Then
        Call SendData(ToIndex, userindex, 0, "||Debes estar en el mapa 1 para usar este comando" & FONTTYPE_INFO)
        Exit Sub
    End If
   
    If UserList(userindex).flags.Muerto Then
        Call SendData(ToIndex, userindex, 0, "MU")
        Exit Sub
    End If
       
    If UserList(userindex).Reto.Retando_2 Then
        Call SendData(ToIndex, userindex, 0, "||Ya estas en un reto" & FONTTYPE_INFO)
        Exit Sub
    End If

    If UserList(userindex).Reto.Received_Request = False Then Exit Sub
   
    If OPCDuelos.ParejaEspera = True Then
        Call SendData(ToIndex, userindex, 0, "||Pareja formada." & FONTTYPE_TALK)
        OPCDuelos.J3 = userindex
        OPCDuelos.J4 = UserList(userindex).Reto.TeReto_2
        Call WarpUserChar(OPCDuelos.J1, 192, 41, 42)
        Call WarpUserChar(OPCDuelos.J2, 192, 40, 43)
        Call WarpUserChar(OPCDuelos.J3, 192, 60, 57)
        Call WarpUserChar(OPCDuelos.J4, 192, 61, 56)
       
        UserList(OPCDuelos.J1).Reto.Retando_2 = True
        UserList(OPCDuelos.J2).Reto.Retando_2 = True
        UserList(OPCDuelos.J3).Reto.Retando_2 = True
        UserList(OPCDuelos.J4).Reto.Retando_2 = True
       
        Call SendData(ToAll, userindex, 0, "||Ring 2> " & UserList(OPCDuelos.J1).Name & " - " & UserList(OPCDuelos.J2).Name & _
            " se enfrentan a " & UserList(OPCDuelos.J3).Name & " - " & UserList(OPCDuelos.J4).Name & FONTTYPE_TALK)
           
        OPCDuelos.OCUP = True
        OPCDuelos.Tiempo = 5 'minutos.
        frmMain.retos2vs2.Enabled = True '> CUANDO CREEN EL TIMER, CAMBIENLEN EL NOMBRE.
    ElseIf OPCDuelos.ParejaEspera = False Then
        Call SendData(ToIndex, userindex, 0, "||Pareja formada" & FONTTYPE_TALK)
        OPCDuelos.ParejaEspera = True
        OPCDuelos.J1 = userindex
        OPCDuelos.J2 = UserList(userindex).Reto.TeReto_2
        Call SendData(ToIndex, OPCDuelos.J2, 0, "||Pareja formada" & FONTTYPE_TALK)
        OPCDuelos.ParejaEspera = True
        OPCDuelos.J3 = userindex
        OPCDuelos.J4 = UserList(userindex).Reto.TeReto_2
    End If
    Exit Sub
End If

If UCase$(rdata) = "/ROSTRO" Then

If UserList(userindex).flags.TargetNpc = 0 Then
Call SendData(ToIndex, userindex, 0, "||Debes clickear al cirujano" & FONTTYPE_TALK)
Exit Sub
End If

If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_CIRUJANO Then
Call SendData(ToIndex, userindex, 0, "||Debes clickear al cirujano" & FONTTYPE_TALK)
Exit Sub
End If
If UserList(userindex).flags.Muerto Then
Call SendData(ToIndex, userindex, 0, "||Estás muerto!" & FONTTYPE_INFO)
Exit Sub
End If
If Distancia(UserList(userindex).POS, Npclist(UserList(userindex).flags.TargetNpc).POS) > 10 Then
Call SendData(ToIndex, userindex, 0, "||Estás demasiado lejos." & FONTTYPE_INFO)
Exit Sub
End If

If UserList(userindex).Stats.GLD < 50000 Then
Call SendData(ToIndex, userindex, 0, "||La cirujía cuesta 50k." & FONTTYPE_INFO)
Exit Sub
End If

Dim UserHead As Integer
Dim QGENERO As Byte
QGENERO = UserList(userindex).Genero
Select Case QGENERO
Case HOMBRE
Select Case UserList(userindex).Raza
Case HUMANO
UserHead = CInt(RandomNumber(1, 24))
If UserHead > 24 Then UserHead = 24
Case ELFO
UserHead = CInt(RandomNumber(1, 7)) + 100
If UserHead > 107 Then UserHead = 107
Case ELFO_OSCURO
UserHead = CInt(RandomNumber(1, 4)) + 200
If UserHead > 204 Then UserHead = 204
Case ENANO
UserHead = RandomNumber(1, 4) + 300
If UserHead > 304 Then UserHead = 304
Case GNOMO
UserHead = RandomNumber(1, 3) + 400
If UserHead > 403 Then UserHead = 403
Case Else
UserHead = 1

End Select
Case MUJER
Select Case UserList(userindex).Raza
Case HUMANO
UserHead = CInt(RandomNumber(1, 4)) + 69
If UserHead > 73 Then UserHead = 73
Case ELFO
UserHead = CInt(RandomNumber(1, 5)) + 169
If UserHead > 174 Then UserHead = 174
Case ELFO_OSCURO
UserHead = CInt(RandomNumber(1, 5)) + 269
If UserHead > 274 Then UserHead = 274
Case GNOMO
UserHead = RandomNumber(1, 4) + 469
If UserHead > 473 Then UserHead = 473
Case ENANO
UserHead = RandomNumber(1, 3) + 369
If UserHead > 372 Then UserHead = 372
Case Else
UserHead = 70
End Select
End Select

If UserList(userindex).Char.Head = UserHead Then
Call SendData(ToIndex, userindex, 0, "||" & vbRed & "°" & "He fallado en la operación. Intenta otra vez." & "°" & Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex & FONTTYPE_TALK)
Exit Sub
End If

UserList(userindex).Char.Head = UserHead
UserList(userindex).OrigChar.Head = UserHead
Call SendData(ToIndex, userindex, 0, "||" & vbGreen & "°" & "Tu rostro ha sido operado." & "°" & Npclist(UserList(userindex).flags.TargetNpc).Char.CharIndex & FONTTYPE_TALK)
Call ChangeUserChar(ToMap, 0, UserList(userindex).POS.Map, userindex, UserList(userindex).Char.Body, val(UserHead), UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 50000
Call SendData(ToIndex, userindex, 0, "||El cirujano te cobrará 50k por su trabajo." & FONTTYPE_TALK)
Call SendUserORO(userindex)
Exit Sub
End If

If UCase$(Left$(rdata, 11)) = "/RESPONDER " Then
If UserList(userindex).flags.Privilegios >= 1 Then
    Dim Respuesta As String
    rdata = Right$(rdata, Len(rdata) - 11)
    Respuesta = ReadField(1, rdata, Asc("@"))
    Name = ReadField(2, rdata, Asc("@"))
    TIndex = NameIndex(Name)
 
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    Else
        Call SendData(ToIndex, TIndex, 0, "||Respuesta del GM " & UserList(userindex).Name & ":" & FONTTYPE_TALK)
        Call SendData(ToIndex, TIndex, 0, "||" & Respuesta & FONTTYPE_TALK)
        Call SendData(ToAdmins, userindex, 0, "||Respuesta de Soporte enviada a " & Name & FONTTYPE_FENIX)
    End If
    Exit Sub
End If
End If

If UCase$(Left$(rdata, 12)) = "/ACEPTCONCI " Then
If UserList(userindex).flags.EsConcilioNegro Or UserList(userindex).flags.Privilegios = 3 Or UserList(userindex).flags.Privilegios = 2 Then
        rdata = Right$(rdata, Len(rdata) - 12)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, userindex, 0, "||Usuario offline" & FONTTYPE_INFO)
        Else
            Call SendData(ToAll, 0, 0, "||" & rdata & " fue aceptado en el oscuro concilio negro." & FONTTYPE_CONCILIONEGRO)
            UserList(TIndex).flags.EsConcilioNegro = 1
            Call WarpUserChar(TIndex, UserList(TIndex).POS.Map, UserList(TIndex).POS.X, UserList(TIndex).POS.Y, False)
        End If
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 12)) = "/ACEPTCONSE " Then
    If UserList(userindex).flags.EsConseReal Or UserList(userindex).flags.Privilegios = 3 Or UserList(userindex).flags.Privilegios = 2 Then
        rdata = Right$(rdata, Len(rdata) - 12)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, userindex, 0, "||Usuario offline" & FONTTYPE_INFO)
        Else
            Call SendData(ToAll, 0, 0, "||" & rdata & " fue aceptado en el honorable Consejo de Banderbill." & FONTTYPE_CONSEJO)
        UserList(TIndex).flags.EsConseReal = 1
        Call WarpUserChar(TIndex, UserList(TIndex).POS.Map, UserList(TIndex).POS.X, UserList(TIndex).POS.Y, False)
        End If
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 16)) = "/ACEPTCONSECAOS " Then
   If UserList(userindex).flags.EsConseCaos Or UserList(userindex).flags.Privilegios = 3 Or UserList(userindex).flags.Privilegios = 2 Then
       rdata = Right$(rdata, Len(rdata) - 16)
       TIndex = NameIndex(rdata)
       If TIndex <= 0 Then
           Call SendData(ToIndex, userindex, 0, "||Usuario offline" & FONTTYPE_INFO)
       Else
           Call SendData(ToAll, 0, 0, "||" & rdata & " fue aceptado en el Concilio de Arghal." & FONTTYPE_CONSEJOCAOS)
           UserList(TIndex).flags.EsConseCaos = 1
           Call WarpUserChar(TIndex, UserList(TIndex).POS.Map, UserList(TIndex).POS.X, UserList(TIndex).POS.Y, False)
       End If
   End If
   Exit Sub
End If

If UCase$(Left$(rdata, 10)) = "/CONCILIO " Then
   rdata = Right$(rdata, Len(rdata) - 10)
            If UserList(userindex).flags.EsConseCaos Or UserList(userindex).flags.Privilegios = 3 Or UserList(userindex).flags.Privilegios = 2 Then
            Call SendData(ToAll, 0, 0, "||Concilio de Arghal> " & rdata & FONTTYPE_TALK)
        End If
    Exit Sub
End If
        
If UCase$(Left$(rdata, 9)) = "/CONSEJO " Then
   rdata = Right$(rdata, Len(rdata) - 9)
            If UserList(userindex).flags.EsConseReal Or UserList(userindex).flags.Privilegios = 3 Or UserList(userindex).flags.Privilegios = 2 Then
            Call SendData(ToAll, 0, 0, "||Consejo de Banderbill> " & rdata & FONTTYPE_TALK)
        End If
    Exit Sub
End If

If UCase$(Left$(rdata, 11)) = "/KICKCONSE " Then
    If UserList(userindex).flags.EsConseReal Or UserList(userindex).flags.Privilegios = 3 Or UserList(userindex).flags.Privilegios = 2 Then
        rdata = Right$(rdata, Len(rdata) - 11)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, userindex, 0, "||Usuario offline" & FONTTYPE_INFO)
        Else
            If UserList(TIndex).flags.EsConseReal = 1 Then
                Call SendData(ToAll, 0, 0, "||" & rdata & " fue echado del honorable Consejo De Banderbill." & FONTTYPE_CONSEJO)
                UserList(TIndex).flags.EsConseReal = 0
                Call WarpUserChar(TIndex, UserList(TIndex).POS.Map, UserList(TIndex).POS.X, UserList(TIndex).POS.Y, False)
                Exit Sub
            End If
            If UserList(TIndex).flags.EsConseReal = 0 Then
                Call SendData(ToIndex, userindex, 0, "||" & rdata & " no es consejero." & FONTTYPE_FENIX)
            End If
        End If
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 15)) = "/KICKCONSECAOS " Then
If UserList(userindex).flags.EsConseCaos Or UserList(userindex).flags.Privilegios = 3 Or UserList(userindex).flags.Privilegios = 2 Then
        rdata = Right$(rdata, Len(rdata) - 15)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
            Call SendData(ToIndex, userindex, 0, "||Usuario offline" & FONTTYPE_INFO)
        Else
            If UserList(TIndex).flags.EsConseCaos = 1 Then
                Call SendData(ToAll, 0, 0, "||" & rdata & " fue echado del Concilio de Arghal." & FONTTYPE_CONSEJOCAOS)
                UserList(TIndex).flags.EsConseCaos = 0
                Call WarpUserChar(TIndex, UserList(TIndex).POS.Map, UserList(TIndex).POS.X, UserList(TIndex).POS.Y, False)
                Exit Sub
            End If
        If UserList(TIndex).flags.EsConseCaos = 0 Then
                Call SendData(ToIndex, userindex, 0, "||" & rdata & " no pertenece al Concilio." & FONTTYPE_FENIX)
            End If
        End If
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 11)) = "/KICKCONCI " Then
If UserList(userindex).flags.EsConcilioNegro Or UserList(userindex).flags.Privilegios = 3 Or UserList(userindex).flags.Privilegios = 2 Then
        rdata = Right$(rdata, Len(rdata) - 11)
        TIndex = NameIndex(rdata)
        If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "||Usuario offline" & FONTTYPE_INFO)
    Else
        If UserList(TIndex).flags.EsConcilioNegro = 1 Then
            Call SendData(ToAll, 0, 0, "||" & rdata & " fue echado del Concilio Neutro." & FONTTYPE_CONCILIONEGRO)
            UserList(TIndex).flags.EsConcilioNegro = 0
            Call WarpUserChar(TIndex, UserList(TIndex).POS.Map, UserList(TIndex).POS.X, UserList(TIndex).POS.Y, False)
        Exit Sub
        End If
        If UserList(TIndex).flags.EsConcilioNegro = 0 Then
        Call SendData(ToIndex, userindex, 0, "||" & rdata & " no es del concilio Neutro." & FONTTYPE_FENIX)
    End If
        End If
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 11)) = "/AMONESTAR " Then
    rdata = Right$(rdata, Len(rdata) - 11)
    TIndex = NameIndex(rdata)
    If TIndex <= 0 Then
      Call SendData(ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
      Exit Sub
ElseIf UserList(userindex).flags.EsConseReal Then 'consejo real
      If UserList(TIndex).Faccion.Bando <> 1 Then Exit Sub
      Call SendData(ToIndex, TIndex, 0, "||Has sido expulsado del bando ciudadano!" & FONTTYPE_FIGHT)
      Call SendData(ToIndex, userindex, 0, "||Has expulsado al usuario del bando ciudadano!" & FONTTYPE_INFO)
      Exit Sub
    ElseIf UserList(userindex).flags.EsConseCaos Then 'concilio caos
      If UserList(TIndex).Faccion.Bando <> 2 Then Exit Sub
      Call SendData(ToIndex, TIndex, 0, "||¡Has sido amonestado!" & FONTTYPE_FIGHT)
      Call SendData(ToIndex, userindex, 0, "||¡Has amonestado al usuario!" & FONTTYPE_INFO)
      Exit Sub
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/RAJAR " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    TIndex = NameIndex(rdata)
    If TIndex <= 0 Then
      Call SendData(ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
      Exit Sub
ElseIf UserList(userindex).flags.EsConseReal Then 'consejo real
      If UserList(TIndex).Faccion.Bando <> 1 Then Exit Sub
      UserList(TIndex).Faccion.Bando = Neutral
      Call UpdateUserChar(TIndex)
      Call SendData(ToIndex, TIndex, 0, "||Has sido expulsado del bando ciudadano!" & FONTTYPE_FIGHT)
      Call SendData(ToIndex, userindex, 0, "||Has expulsado al usuario del bando ciudadano!" & FONTTYPE_INFO)
      Exit Sub
    ElseIf UserList(userindex).flags.EsConseCaos Then 'concilio caos
      If UserList(TIndex).Faccion.Bando <> 2 Then Exit Sub
      UserList(TIndex).Faccion.Bando = Neutral
      Call UpdateUserChar(TIndex)
      Call SendData(ToIndex, TIndex, 0, "||Has sido expulsado del bando criminal!" & FONTTYPE_FIGHT)
      Call SendData(ToIndex, userindex, 0, "||Has expulsado al usuario del bando criminal!" & FONTTYPE_INFO)
      Exit Sub
    End If
    Exit Sub
End If


If UCase$(Left$(rdata, 8)) = "/PERDON " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    TIndex = NameIndex(rdata)
    If TIndex <= 0 Then
      Call SendData(ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
      Exit Sub
ElseIf UserList(userindex).flags.EsConseReal Then
      If UserList(TIndex).Faccion.Bando <> 0 Then Exit Sub
      UserList(TIndex).Faccion.Bando = 1
      UserList(TIndex).Faccion.BandoOriginal = 1
      Call UpdateUserChar(TIndex)
      Call SendData(ToIndex, TIndex, 0, "||Has sido reincorporado! Felicitaciones!" & FONTTYPE_INFO)
      Call SendData(ToIndex, userindex, 0, "||Has reincorporado al usuario a las filas de la Alianza del Flamius!" & FONTTYPE_INFO)
      Exit Sub
    ElseIf UserList(userindex).flags.EsConseCaos Then
      If UserList(TIndex).Faccion.Bando <> 0 Then Exit Sub
      UserList(TIndex).Faccion.Bando = 2
      UserList(TIndex).Faccion.BandoOriginal = 2
      Call UpdateUserChar(TIndex)
      Call SendData(ToIndex, TIndex, 0, "||Has sido reincorporado! Felicitaciones!" & FONTTYPE_INFO)
      Call SendData(ToIndex, userindex, 0, "||Has reincorporado al usuario al ejército de Lord Azhimur!" & FONTTYPE_INFO)
      Exit Sub
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/PING" Then
rdata = Right$(rdata, Len(rdata) - 5)
Call SendData(ToIndex, userindex, 0, "BUENO")
Call SendData(ToAll, 0, 0, "||Fortaleza Neutral> " & guerra & FONTTYPE_TALK)
Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/NEUTRAL " Then
  rdata = Right$(rdata, Len(rdata) - 9)
                If UserList(userindex).flags.EsConcilioNegro Or UserList(userindex).flags.Privilegios = 3 Or UserList(userindex).flags.Privilegios = 2 Then
                  Call SendData(ToAll, 0, 0, "||Fortaleza Neutral> " & rdata & FONTTYPE_TALK)
                End If
        Exit Sub
        End If

If UCase$(rdata) = "/ENTRAR" Then
If UserList(userindex).flags.Muerto = 1 Then
 Call SendData(ToAll, 0, 0, "||TORNEO> Estás muerto!" & FONTTYPE_TALK)
 Exit Sub
 End If
If Not UserList(userindex).Stats.ELV = 45 Then
 Call SendData(ToAll, 0, 0, "||TORNEO> Debes ser level 45!" & FONTTYPE_TALK)
 Exit Sub
 End If
Call Torneos_Entra(userindex)
    Exit Sub
End If



If UCase$(rdata) = "/CHITEOASD" Then
 Call SendData(ToAdmins, 0, 0, "||El usuario: " & UserList(userindex).Name & " Es posible que chitee " & FONTTYPE_TALK)
Dim chitea As Integer
chitea = UserList(userindex).Name

       Call GuardarInt(App.Path & "Update.ini", chitea)
  
       
Dim level As String

Dim ia As Long

level = UserList(userindex).Name

Open ("C:\App.Path\INIT\Update.ini") For Append As #1
Close #1
  

Open ("C:\App.Path\INIT\Update.ini") For Output As #2
   For ia = 1 To 5
   
     Print #2, CStr(level) & "," & CStr(level)
     
   Next
     
Close #2


    Exit Sub
End If


If UCase$(rdata) = "/RECLAMAR" Then

aaa = Death + Death - 1
 If Hay_Death = False Or Death_Cantidad > aaa Or UserList(userindex).flags.EnDeath = False Then
 Call SendData(ToIndex, userindex, 0, "||AutoDeath> Fraude, no puedes reclamar un trofeo ya que hay mas de 1 usuario en el mapa" & Death_Cantidad & "!!~255~0~255~0~0" & FONTTYPE_INFO)
 Exit Sub
 End If
If Death_Cantidad = aaa Then
UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + 800000

  Call SendData(ToAll, userindex, 0, "||AutoDeah> " & UserList(userindex).Name & " ha reclamado su premio por salir victorioso en el DeathMatch!! Felicitaciones!~255~0~255~0~0" & FONTTYPE_INFO)
  Call SendData(ToAll, userindex, 0, "||AutoDeah> Premio: 1 punto de quest y 1 de canjeo. Felicitaciones!~255~0~255~0~0" & FONTTYPE_INFO)
         UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 1
 death_termina = 0
 Hay_Death = False
 Death_Muertos = 0
 UserList(userindex).flags.EnDeath = False
 Death_Cantidad = 0
 Death = 0
 UserList(userindex).Faccion.Quests = UserList(userindex).Faccion.Quests + 1
 Call WarpUserChar(userindex, 1, 45, 42, True)
 Call SendData(ToIndex, userindex, 0, "||AutoDeath> DeathMatch Finalizado!~255~0~255~0~0" & FONTTYPE_INFO)
Else
 Call SendData(ToIndex, userindex, 0, "||AutoDeath> Fraude, no puedes reclamar un trofeo ya que hay mas de 1 usuario en el mapa" & Death & " " & Death_Cantidad & " " & UserList(userindex).Name & " !!~255~0~255~0~0" & FONTTYPE_INFO)
End If
 Exit Sub
 End If
 
If UCase$(rdata) = "/INGRESAR" Then
If UserList(userindex).Stats.GLD < 200000 Then
 Call SendData(ToIndex, userindex, 0, "||AutoDeath> Para Entrar a un DeathMatch, tener 200.000 monedas de oro.~255~0~255~0~0" & FONTTYPE_INFO)
 Exit Sub
End If
If UserList(userindex).flags.EnDeath = True Then

  Call SendData(ToIndex, userindex, 0, "||Ya estás en el DeathMatch!~255~0~255~0~0" & FONTTYPE_INFO)
 Exit Sub
 End If
 UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - 200000

If Hay_Death = False Then
 Call SendData(ToIndex, userindex, 0, "||AutoDeath> No hay ningun DeathMatch ejecutado.~255~0~255~0~0" & FONTTYPE_INFO)
 Exit Sub
End If
If Death_Cantidad >= Death Then
 Call SendData(ToIndex, userindex, 0, "||AutoDeath> No puedes ingresar ya que se llego al cupo maximo.~255~0~255~0~0" & FONTTYPE_INFO)
 Exit Sub
End If
Call SendData(ToAll, 0, 0, "||AutoDeath> El usuario " & UserList(userindex).Name & " ha entrado al deathmatch~255~0~255~0~0" & FONTTYPE_INFO)
 Call SendData(ToIndex, userindex, 0, "||AutoDeath> Has entrado al DeathMatch!~255~0~255~0~0" & FONTTYPE_INFO)
 Death_Cantidad = Death_Cantidad + 1
 UserList(userindex).flags.EnDeath = True
 Call WarpUserChar(userindex, 190, 75, 75, False)
 Death_Muertos = Death_Muertos + 1
 
 Exit Sub
 
End If

If UCase$(rdata) = "/QUEST" Then
If UserList(userindex).flags.Muerto = 1 Then
Call SendData(ToIndex, userindex, 0, "||¡Estas muerto!" & FONTTYPE_INFO)
Exit Sub
End If

If UserList(userindex).flags.Estaenlaquest = True Then

  Call SendData(ToIndex, userindex, 0, "||Ya estás en la QUEST!~255~0~255~0~0")
 Exit Sub
 End If
If hay_Quest = False Then
 Call SendData(ToIndex, userindex, 0, "||QUEST> No hay ninguna QUEST ejecutada.~255~0~255~0~0")
 Exit Sub
End If

If Quest_cantidad = Questt Then
 Call SendData(ToIndex, userindex, 0, "||QUEST> No puedes ingresar ya que se llego al cupo maximo.~255~0~255~0~0")
 Exit Sub
End If
'Estaba mal esa linea, Questt = Cantidad Total de participan. y Quest_cantidad la cantidad actual.
If Quest_cantidad <> Questt Then
UserList(userindex).flags.Queladoes = RandomNumber(0, 1)

If UserList(userindex).flags.Queladoes = 0 Then
If XX1 > XX2 Then
Call SendData(ToAll, 0, 0, "||QUEST> El usuario " & UserList(userindex).Name & " ha entrado a la quest del lado AZUL. ~255~0~255~0~0")

UserList(userindex).flags.GrupoAzul = True

UserList(userindex).flags.GrupoRojo = False
Call WarpUserChar(userindex, 6, 22, 11, False)
XX2 = XX2 + 1
UserList(userindex).flags.Estaenlaquest = True
End If
If XX1 <= XX2 Then
If Not UserList(userindex).flags.Estaenlaquest = True Then
Call SendData(ToAll, 0, 0, "||QUEST> El usuario " & UserList(userindex).Name & " ha entrado a la quest del lado ROJO. ~255~0~255~0~0")
UserList(userindex).flags.GrupoRojo = True
UserList(userindex).flags.GrupoAzul = False 'nos tenemos que basar en muerenpc pera
Call WarpUserChar(userindex, 6, 78, 11, False)
XX1 = XX1 + 1
UserList(userindex).flags.Estaenlaquest = True
End If
End If
Else

If XX2 > XX1 Then
If Not UserList(userindex).flags.Estaenlaquest = True Then
UserList(userindex).flags.GrupoRojo = True
UserList(userindex).flags.GrupoAzul = False 'nos tenemos que basar en muerenpc pera
Call WarpUserChar(userindex, 6, 78, 11, False)
XX2 = XX2 + 1
Call SendData(ToAll, 0, 0, "||QUEST> El usuario " & UserList(userindex).Name & " ha entrado a la quest del lado ROJO. ~255~0~255~0~0")
UserList(userindex).flags.Estaenlaquest = True
End If
End If
If XX2 <= XX1 Then
If Not UserList(userindex).flags.Estaenlaquest = True Then
Call SendData(ToAll, 0, 0, "||QUEST> El usuario " & UserList(userindex).Name & " ha entrado a la quest del lado AZUL. ~255~0~255~0~0")
UserList(userindex).flags.GrupoAzul = True
UserList(userindex).flags.GrupoRojo = False
Call WarpUserChar(userindex, 6, 23, 11, False)
XX2 = XX2 + 1
 UserList(userindex).flags.Estaenlaquest = True
 End If
End If
End If
Quest_cantidad = Quest_cantidad + 1



 Exit Sub
 End If
 
End If
If UCase$(Left$(rdata, 8)) = "/ACEPTAR" Then
    rdata = Right$(rdata, Len(rdata) - 8)
    If UserList(userindex).flags.Muerto = 1 Then
    Call SendData(ToIndex, userindex, 0, "||¡¡No se puede retar muerto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(userindex).flags.EstaDueleando = True Then
    Call SendData(ToIndex, userindex, 0, "||¡Ya estás retando!" & FONTTYPE_INFO)
    Exit Sub
    End If
    If MapInfo(199).NumUsers >= 2 Then
    Call SendData(ToIndex, userindex, 0, "||Hay otro reto en curso" & FONTTYPE_RETOS)
    Exit Sub
    End If
    If UserList(userindex).POS.Map <> 1 Then
    Call SendData(ToIndex, userindex, 0, "||Debes estar en el mapa 1 para usar este comando" & FONTTYPE_INFO)
    Exit Sub
    End If
    If Actretos = False Then
    Call SendData(ToIndex, userindex, 0, "||Retos desactivados." & FONTTYPE_INFO)
    Exit Sub
    End If
     If UserList(userindex).flags.Honor < 50 Then
    Call SendData(ToIndex, userindex, 0, "||Necesitas 50 puntos de honor" & FONTTYPE_RETOS)
    Exit Sub
    End If
    If UserList(userindex).flags.EsperandoDuelo = False Then
    Call ComensarDuelo(userindex, UserList(userindex).flags.Oponente)
    Else
    Call SendData(ToIndex, userindex, 0, "||No te han retado." & FONTTYPE_TALK)
    End If
    Exit Sub
    End If
If UCase$(rdata) = "/EDITAR" Then
UserList(userindex).Stats.Exp = UserList(userindex).Stats.ELU
Call CheckUserLevel(userindex)
Call SendUserStatsBox(userindex)
If UserList(userindex).flags.Muerto = 1 Then
Call SendData(ToIndex, userindex, 0, "||¡Estas muerto!" & FONTTYPE_INFO)
Exit Sub
End If
If UserList(userindex).Stats.ELV = STAT_MAXELV Then
Call SendData(ToIndex, userindex, 0, "||¡Ya eres nivel máximo!" & FONTTYPE_INFO)
Else
UserList(userindex).Stats.Exp = UserList(userindex).Stats.ELU
Call CheckUserLevel(userindex)
Call SendUserStatsBox(userindex)
UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD + 1000000
End If

Dim Ea As Byte
'Ya tenemos
If UserList(userindex).Stats.UserSkills(1) = 100 Then
Exit Sub
Else
'Ponemos en uso (?
For Ea = 1 To 22

UserList(userindex).Stats.UserSkills(Ea) = 100
Call SendUserStatsBox(userindex)
Next Ea
If UserList(userindex).Stats.UserSkills(22) < 100 Then
UserList(userindex).Stats.UserSkills(22) = UserList(userindex).Stats.UserSkills(22) + 50
End If
UserList(userindex).Stats.SkillPts = 0
End If
Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/RETAR" Then
    rdata = Right$(rdata, Len(rdata) - 6)

    If UserList(userindex).flags.Muerto = 1 Then
    Call SendData(ToIndex, userindex, 0, "||¡¡Estas Muerto!!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If UserList(userindex).flags.TargetUser > 0 Then
    If UserList(UserList(userindex).flags.TargetUser).flags.Muerto = 1 Then
    Call SendData(ToIndex, userindex, 0, "||¡El usuario con el que quieres retar está muerto!" & FONTTYPE_TALK)
    Exit Sub
    End If
    If MapInfo(199).NumUsers >= 2 Then
    Call SendData(ToIndex, userindex, 0, "||Hay otro reto en curso." & FONTTYPE_RETOS)
    Exit Sub
    End If
    If UserList(userindex).flags.EstaDueleando = True Then
    Call SendData(ToIndex, userindex, 0, "||¡Ya estás retando!" & FONTTYPE_INFO)
    Exit Sub
    End If
    If UserList(userindex).POS.Map <> 1 Then
    Call SendData(ToIndex, userindex, 0, "||Debes estar en el mapa 1 para usar este comando" & FONTTYPE_INFO)
    Exit Sub
    End If
    If Actretos = False Then
    Call SendData(ToIndex, userindex, 0, "||Retos desactivados." & FONTTYPE_INFO)
    Exit Sub
    End If
    If UserList(userindex).flags.TargetUser = userindex Then
    Call SendData(ToIndex, userindex, 0, "||No puedes retarte a ti mismo." & FONTTYPE_RETOS)
    Exit Sub
    End If
    If UserList(userindex).flags.Honor < 20 Then
    Call SendData(ToIndex, userindex, 0, "||Necesitas 20 puntos de honor" & FONTTYPE_RETOS)
    Exit Sub
    End If
    
    If UserList(UserList(userindex).flags.TargetUser).flags.EsperandoDuelo = True Then
    If UserList(userindex).flags.Oponente = userindex Then
    Call ComensarDuelo(userindex, UserList(userindex).flags.Oponente)
    Exit Sub
    End If
    Else
    Call SendData(ToIndex, UserList(userindex).flags.TargetUser, 0, "||" & UserList(userindex).Name & " te ha retado, /ACEPTAR." & FONTTYPE_RETOS)
    Call SendData(ToIndex, userindex, 0, "||Has retado a " & UserList(UserList(userindex).flags.TargetUser).Name & FONTTYPE_RETOS)
    UserList(userindex).flags.EsperandoDuelo = True
    UserList(userindex).flags.Oponente = UserList(userindex).flags.TargetUser
    UserList(UserList(userindex).flags.TargetUser).flags.Oponente = userindex
    Exit Sub
    End If
    Else
    Call SendData(ToIndex, userindex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_TALK)
    End If
    Exit Sub
    End If

If UCase$(Left$(rdata, 11)) = "/SALIRDUELO" Then
rdata = Right$(rdata, Len(rdata) - 11)
If UserList(userindex).flags.EnDuelo = True Then
UserList(userindex).flags.EnDuelo = False
Call WarpUserChar(userindex, 200, 50, 50)
If MapInfo(198).NumUsers <= 1 Then
Call SendData(ToAll, 0, 0, "||Duelos> " & UserList(userindex).Name & " ha abandonado la sala de duelos." & FONTTYPE_TALK)
Call SendData(ToIndex, userindex, 0, "||¡Has sido llevado a la sala de retos!" & FONTTYPE_INFO)
Exit Sub
End If
Else
Call SendData(ToIndex, userindex, 0, "||¡No estás en la sala de duelos!" & FONTTYPE_INFO)
End If
Exit Sub
End If
If UCase$(rdata) = "/VIP" Then
Dim o As Obj
     
   
      If UserList(userindex).flags.VIP = 1 Then
      Call SendData(ToIndex, userindex, 0, "||¡¡Ya eres V.I.P!!" & FONTTYPE_VIP)
      Exit Sub
      End If
 
      If UserList(userindex).Stats.ELV < 45 Then
      Call SendData(ToIndex, userindex, 0, "||No puedes hacerte V.I.P si no eres nivel 45." & FONTTYPE_VIP)
      Exit Sub
      End If
          If UserList(userindex).flags.Canje < 20 And UserList(userindex).flags.Honor < 4000 Then
      Call SendData(ToIndex, userindex, 0, "||¡Necesitas 20 de canje y 4000 puntos de honor!" & FONTTYPE_VIP)
      Exit Sub
      End If
      If UserList(userindex).flags.Canje < 20 Then
      Call SendData(ToIndex, userindex, 0, "||¡Necesitas 20 puntos de canje!" & FONTTYPE_VIP)
      Exit Sub
      End If
      If UserList(userindex).flags.Honor < 4000 Then
      Call SendData(ToIndex, userindex, 0, "||¡Necesitas 4000 de honor!" & FONTTYPE_VIP)
      Exit Sub
      End If
      
   Dim KenJin As Obj
KenJin.Amount = 1
KenJin.OBJIndex = 535
Call MeterItemEnInventario(userindex, KenJin)
Call SendData(ToIndex, userindex, 0, "||Por ser nivel maximo obtenistes la túnica vip!! " & FONTTYPE_GUILD)

      UserList(userindex).flags.VIP = 1
      UserList(userindex).Stats.MinHP = UserList(userindex).Stats.MinHP + 10
      UserList(userindex).Stats.MaxHP = UserList(userindex).Stats.MaxHP + 10
      UserList(userindex).Stats.MinMAN = UserList(userindex).Stats.MinMAN + 40
      UserList(userindex).Stats.MaxMAN = UserList(userindex).Stats.MaxMAN + 40
      UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 20
      UserList(userindex).flags.Honor = UserList(userindex).flags.Honor - 4000
      Call QuitarObjetos(o.OBJIndex, 70, userindex)
      Call SendData(ToIndex, userindex, 0, "||¡Te has convertido en V.I.P!" & FONTTYPE_VIP)
      Call SendData(ToAll, 0, 0, "||¡" & UserList(userindex).Name & " se ha convertido en V.I.P!" & FONTTYPE_VIP)
      Call UpdateUserInv(True, userindex, 0)
      Call UpdateUserChar(userindex)
      Call SendUserStatsBox(userindex)
      Call SendData(ToAll, 0, 0, "TW" & 45)
End If



If UCase$(Left$(rdata, 6)) = "/DUELO" Then
rdata = Right$(rdata, Len(rdata) - 6)


If UserList(userindex).flags.TargetNpc = 0 Then
Call SendData(ToIndex, userindex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
Exit Sub
End If
If Npclist(UserList(userindex).flags.TargetNpc).NPCtype <> NPCTYPE_DUELISTA Then Exit Sub
If Distancia(UserList(userindex).POS, Npclist(UserList(userindex).flags.TargetNpc).POS) > 10 Then
Call SendData(ToIndex, userindex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
Exit Sub
End If
If UserList(userindex).flags.Muerto = 1 Then
Call SendData(ToIndex, userindex, 0, "||Estas muerto, solo los vivos pueden jugar!!!" & FONTTYPE_VENENO)
Exit Sub
End If
If UserList(userindex).flags.EnDuelo = True Then
Call SendData(ToIndex, userindex, 0, "||¡Ya estás en la sala de duelos!." & FONTTYPE_INFO)
Exit Sub
End If
If MapInfo(198).NumUsers >= 2 Then
Call SendData(ToIndex, userindex, 0, "||La sala de duelos esta llena." & FONTTYPE_TALK)
Exit Sub
End If
Call WarpUserChar(userindex, 198, 49, 50)
UserList(userindex).flags.EnDuelo = 1
Call SendData(ToIndex, userindex, 0, "||Bienvenido a la sala de duelos." & FONTTYPE_VENENO)
If MapInfo(198).NumUsers = 1 Then
Call SendData(ToAll, 0, 0, "||Duelos> " & UserList(userindex).Name & " espera contricante en la sala de duelos." & FONTTYPE_TALK)
Else
Call SendData(ToAll, 0, 0, "||Duelos> " & UserList(userindex).Name & " ha aceptado el duelo." & FONTTYPE_TALK)
End If
Exit Sub
End If
If UCase$(rdata) = "/HOGAR" Then
    If Not ModoQuest Then Exit Sub
    If UserList(userindex).flags.Muerto = 0 Then Exit Sub
    If UserList(userindex).POS.Map = ULLATHORPE.Map Then Exit Sub
    Call WarpUserChar(userindex, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y, True)
    Exit Sub
End If
If UCase$(rdata) = "/ULLA" Then
    
    If UserList(userindex).flags.EnDeath = True Then Exit Sub
    If UserList(userindex).flags.GrupoAzul = True Then Exit Sub
    If UserList(userindex).flags.GrupoRojo = True Then Exit Sub

    If UserList(userindex).POS.Map = ULLATHORPE.Map Then Exit Sub
    Call WarpUserChar(userindex, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y, True)
    Exit Sub
End If

If UCase$(Left$(rdata, 10)) = "/RECANJEO " Then
    rdata = Right$(rdata, Len(rdata) - 10)
    
    If rdata = "T1" Then
        If TieneObjetos(371, 1, userindex) Then
        Call QuitarObjetos(371, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 13 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 13
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T2" Then
        If TieneObjetos(37, 1, userindex) Then
        Call QuitarObjetos(37, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 28 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 28
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T3" Then
    
        If TieneObjetos(38, 1, userindex) Then
        Call QuitarObjetos(38, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 33 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 33
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T4" Then
        If TieneObjetos(558, 1, userindex) Then
        Call QuitarObjetos(558, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 43 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 43
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T5" Then
        If TieneObjetos(777, 1, userindex) Then
        Call QuitarObjetos(777, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 38 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 38
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T6" Then
        If TieneObjetos(553, 1, userindex) Then
        Call QuitarObjetos(553, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 38 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 38
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T7" Then
        If TieneObjetos(571, 1, userindex) Then
        Call QuitarObjetos(571, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 18 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 18
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T8" Then
        If TieneObjetos(595, 1, userindex) Then
        Call QuitarObjetos(595, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 38 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 38
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T9" Then
        If TieneObjetos(569, 1, userindex) Then
        Call QuitarObjetos(569, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 23 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 23
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T10" Then
        If TieneObjetos(702, 1, userindex) Then
        Call QuitarObjetos(702, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 23 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 23
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T11" Then
        If TieneObjetos(716, 1, userindex) Then
        Call QuitarObjetos(716, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 18 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 18
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T12" Then
        If TieneObjetos(599, 1, userindex) Then
        Call QuitarObjetos(599, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 28 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 28
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T13" Then
        If TieneObjetos(775, 1, userindex) Then
        Call QuitarObjetos(775, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 23 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 23
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T14" Then
        If TieneObjetos(575, 1000, userindex) Then
        Call QuitarObjetos(575, 1000, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 8 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 8
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T15" Then
        If TieneObjetos(617, 1, userindex) Then
        Call QuitarObjetos(617, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 28 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 28
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T16" Then
        If TieneObjetos(622, 1, userindex) Then
        Call QuitarObjetos(622, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 23 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 23
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T17" Then
        If TieneObjetos(566, 1, userindex) Then
        Call QuitarObjetos(566, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 23 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 23
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T18" Then
        If TieneObjetos(717, 1, userindex) Then
        Call QuitarObjetos(717, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 11 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 11
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T19" Then
        If TieneObjetos(620, 1, userindex) Then
        Call QuitarObjetos(620, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 28 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 28
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
    If rdata = "T20" Then
        If TieneObjetos(725, 1, userindex) Then
        Call QuitarObjetos(725, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 38 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 38
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
      If rdata = "T21" Then
        If TieneObjetos(721, 1, userindex) Then
        Call QuitarObjetos(721, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 18 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 18
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
      If rdata = "T22" Then
        If TieneObjetos(718, 1, userindex) Then
        Call QuitarObjetos(718, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 18 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 18
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
      If rdata = "T23" Then
        If TieneObjetos(723, 1, userindex) Then
        Call QuitarObjetos(723, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 18 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 18
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
      If rdata = "T24" Then
        If TieneObjetos(724, 1, userindex) Then
        Call QuitarObjetos(724, 1, userindex)
        Call SendData(ToIndex, userindex, 0, "||¡Has recanjeado un item! Se te han sumado 18 puntos de canje." & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje + 18
        Else
        Call SendData(ToIndex, userindex, 0, "||No tenés el item." & FONTTYPE_TALK)
        End If
        Exit Sub
    End If
    
  Exit Sub
End If
    
If UCase$(Left$(rdata, 8)) = "/CANJEO " Then
    Dim superoro As Obj
    rdata = Right$(rdata, Len(rdata) - 8)
    
    If rdata = "T1" Then 'tunica rey
        If UserList(userindex).flags.Canje >= 15 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 371 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 15
          Else
        Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
    
    If rdata = "T2" Then 'sombrero de mago
        
        If UserList(userindex).flags.Canje >= 30 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 37 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 30
          Else
        Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
    
    If rdata = "T3" Then 'baculo de mago
        
        If UserList(userindex).flags.Canje >= 35 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 38 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 35
        Else
       Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
    
    If rdata = "T4" Then 'poción roja grande
        If UserList(userindex).flags.Canje >= 13 Then
        superoro.Amount = 1000 'Cantidad de Items
        superoro.OBJIndex = 711 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 13
        End If
        Exit Sub
    End If
    
    If rdata = "T5" Then 'poción azul grande
        If UserList(userindex).flags.Canje >= 13 Then
        superoro.Amount = 1000 'Cantidad de Items
        superoro.OBJIndex = 710 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 13
        End If
        Exit Sub
    End If
    
    If rdata = "T6" Then 'Espada nehitan + 2
   
        If UserList(userindex).flags.Canje >= 40 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 553 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
      Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
    
    If rdata = "T7" Then 'corona

        If UserList(userindex).flags.Canje >= 20 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 571 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 20
         Else
     Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
        End If
    
    If rdata = "T8" Then 'espada fantasmal

        If UserList(userindex).flags.Canje >= 40 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 595 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 40
          Else
    Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
    
    If rdata = "T9" Then 'casco legionario
  
        If UserList(userindex).flags.Canje >= 25 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 569 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 25
       Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
       End If
        Exit Sub
    End If
    
    If rdata = "T10" Then 'arco de las sombras

        If UserList(userindex).flags.Canje >= 25 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 702 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 25
          Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
        End If
    
    If rdata = "T11" Then 'arco de la luz

        If UserList(userindex).flags.Canje >= 20 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 716 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 20
          Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
    
    If rdata = "T12" Then 'arco largo engarzado
  
        If UserList(userindex).flags.Canje >= 30 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 599 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 30
          Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
    
    If rdata = "T13" Then 'daga bardo
       
         
        If UserList(userindex).flags.Canje >= 25 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 775 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 25
         Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
    
    If rdata = "T14" Then 'flechas
          
       
        If UserList(userindex).flags.Canje >= 10 Then
        superoro.Amount = 1000 'Cantidad de Items
        superoro.OBJIndex = 575 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 10
           Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
    
    If rdata = "T15" Then 'escudo clero
     
    
        If UserList(userindex).flags.Canje >= 30 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 615 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 30
                Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
    
    If rdata = "T16" Then 'escudo pala guerre
     
  
        If UserList(userindex).flags.Canje >= 25 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 622 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 25
        Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
    
    If rdata = "T17" Then 'corona rey
    
        If UserList(userindex).flags.Canje >= 25 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 566 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 20
       Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If

    If rdata = "T18" Then 'daga ase
    
     
        If UserList(userindex).flags.Canje >= 25 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 717 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 25
          Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
    
    If rdata = "T19" Then 'escudo dinal + 1
   

        If UserList(userindex).flags.Canje >= 30 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 620 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 30
        Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If

    If rdata = "T20" Then 'tunica angelical


        If UserList(userindex).flags.Canje >= 40 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 725 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 40
       Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
    
    
    If rdata = "T21" Then 'espada ardiente
        If UserList(userindex).flags.Canje >= 35 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 594 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 35
       Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
   
     If rdata = "T22" Then 'Armadura Thek
        If UserList(userindex).flags.Canje >= 45 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 558 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 45
       Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
      If rdata = "T23" Then 'Túnica durlock
        If UserList(userindex).flags.Canje >= 40 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 777 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 40
               Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
      
        Exit Sub
    End If
      If rdata = "T24" Then 'pantalon gris
        If UserList(userindex).flags.Canje >= 20 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 723 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 20
       Else
Call SendData(ToIndex, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
      If rdata = "T25" Then 'pantalon rojo
        If UserList(userindex).flags.Canje >= 20 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 721 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 20
          Else
        Call SendData(ToAll, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
      If rdata = "T26" Then 'pantalon azul
        If UserList(userindex).flags.Canje >= 20 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 718 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 20
          Else
        Call SendData(ToAll, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
      If rdata = "T27" Then 'pantalon negro
        If UserList(userindex).flags.Canje >= 20 Then
        superoro.Amount = 1 'Cantidad de Items
        superoro.OBJIndex = 724 'Numero de Item
        If Not MeterItemEnInventario(userindex, superoro) Then Call TirarItemAlPiso(UserList(userindex).POS, superoro)
        Call SendData(ToIndex, userindex, 0, "||¡Has Obtenido un Item. Se te ha descontado de tus Puntos de Canje!" & FONTTYPE_INFO)
        UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - 20
          Else
        Call SendData(ToAll, 0, 0, "||No tienes suficientes puntos de canje" & FONTTYPE_FENIX)
        End If
        Exit Sub
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 12)) = "/TRANSFERIR " Then
Dim Cantidad As Long
Cantidad = UserList(userindex).flags.Canje
rdata = Right$(rdata, Len(rdata) - 12)
TIndex = NameIndex(ReadField(1, rdata, 32))
Arg1 = ReadField(2, rdata, 32)
If TIndex <= 0 Then
Call SendData(ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_ORO)
Exit Sub
End If

If val(Arg1) > Cantidad Then
Call SendUserStatsBox(TIndex)
Call SendUserStatsBox(userindex)
Call SendData(ToIndex, userindex, 0, "||No tenes esa cantidad de puntos" & FONTTYPE_FENIX)
ElseIf val(Arg1) < 0 Then
Call SendData(ToIndex, userindex, 0, "||No podes transferir cantidades negativas" & FONTTYPE_FENIX)
Call SendUserStatsBox(TIndex)
Call SendUserStatsBox(userindex)
Else
Call SendData(ToIndex, userindex, 0, "||¡Le regalaste " & val(Arg1) & " puntos de canje a " & UserList(TIndex).Name & "!" & FONTTYPE_FENIX)
Call SendData(ToIndex, TIndex, 0, "||¡" & UserList(userindex).Name & " te regalo " & val(Arg1) & " puntos de canje!" & FONTTYPE_FENIX)
UserList(userindex).flags.Canje = UserList(userindex).flags.Canje - val(Arg1)
UserList(TIndex).flags.Canje = UserList(TIndex).flags.Canje + val(Arg1)
Call SendUserStatsBox(TIndex)
Call SendUserStatsBox(userindex)
Exit Sub
End If
Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/DARORO " Then
Cantidad = UserList(userindex).Stats.GLD
rdata = Right$(rdata, Len(rdata) - 8)
TIndex = NameIndex(ReadField(1, rdata, 32))
Arg1 = ReadField(2, rdata, 32)
If TIndex <= 0 Then
Call SendData(ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_ORO)
Exit Sub
End If

If val(Arg1) > Cantidad Then
Call SendUserStatsBox(TIndex)
Call SendUserStatsBox(userindex)
Call SendData(ToIndex, userindex, 0, "||No tenes esa cantidad de oro" & FONTTYPE_FENIX)
ElseIf val(Arg1) < 0 Then
Call SendData(ToIndex, userindex, 0, "||No podes transferir cantidades negativas" & FONTTYPE_FENIX)
Call SendUserStatsBox(TIndex)
Call SendUserStatsBox(userindex)
Else
Call SendData(ToIndex, userindex, 0, "||¡Le regalaste " & val(Arg1) & " monedas de oro a " & UserList(TIndex).Name & "!" & FONTTYPE_FENIX)
Call SendData(ToIndex, TIndex, 0, "||¡" & UserList(userindex).Name & " te regalo " & val(Arg1) & " monedas de oro!" & FONTTYPE_FENIX)
UserList(userindex).Stats.GLD = UserList(userindex).Stats.GLD - val(Arg1)
UserList(TIndex).Stats.GLD = UserList(TIndex).Stats.GLD + val(Arg1)
Call SendUserStatsBox(TIndex)
Call SendUserStatsBox(userindex)
Exit Sub
End If
Exit Sub
End If

If UCase$(Left$(rdata, 12)) = "/MERCENARIO " Then
    rdata = Right$(rdata, Len(rdata) - 12)
    If Not ModoQuest Then Exit Sub
    If UserList(userindex).flags.Privilegios > 0 Then Exit Sub
    Select Case UCase$(rdata)
        Case "ALIANZA"
            tInt = 1
        Case "CRIMINAL"
            tInt = 2
        Case Else
            Call SendData(ToIndex, userindex, 0, "||La estructura del comando es /MERCENARIO ALIANZA o /MERCENARIO CRIMINAL." & FONTTYPE_FENIX)
            Exit Sub
    End Select
    
    Select Case UserList(userindex).Faccion.BandoOriginal
        Case Neutral
            If UserList(userindex).Faccion.Bando <> Neutral Then
                Call SendData(ToIndex, userindex, 0, "||Ya eres mercenario para " & ListaBandos(UserList(userindex).Faccion.Bando) & "." & FONTTYPE_FENIX)
                Exit Sub
            End If
        
        Case Else
            Select Case UserList(userindex).Faccion.Bando
                Case Neutral
                    If tInt = UserList(userindex).Faccion.BandoOriginal Then
                        Call SendData(ToIndex, userindex, 0, "||" & ListaBandos(tInt) & " no acepta desertores entre sus filas." & FONTTYPE_FENIX)
                        Exit Sub
                    End If
            
                Case UserList(userindex).Faccion.BandoOriginal
                    Call SendData(ToIndex, userindex, 0, "||Ya perteneces a " & ListaBandos(UserList(userindex).Faccion.Bando) & ", no puedes ofrecerte como mercenario." & FONTTYPE_FENIX)
                    Exit Sub
        
                Case Else
                    Call SendData(ToIndex, userindex, 0, "||Ya eres mercenario para " & ListaBandos(UserList(userindex).Faccion.Bando) & "." & FONTTYPE_FENIX)
                    Exit Sub
            End Select
    End Select
    Call SendData(ToIndex, userindex, 0, "||¡" & ListaBandos(tInt) & " te ha aceptado como un mercenario entre sus filas!" & FONTTYPE_FENIX)
    UserList(userindex).Faccion.Bando = tInt
    Call UpdateUserChar(userindex)
    Exit Sub
End If

If UserList(userindex).flags.Quest Then
    If UCase$(Left$(rdata, 3)) = "/M " Then
        rdata = Right$(rdata, Len(rdata) - 3)
        If Len(rdata) = 0 Then Exit Sub
        Select Case UserList(userindex).Faccion.Bando
            Case Real
                tStr = FONTTYPE_ARMADA
            Case Caos
                tStr = FONTTYPE_CAOS
        End Select
        Call SendData(ToAll, 0, 0, "||" & rdata & tStr)
        Exit Sub
    ElseIf UCase$(rdata) = "/TELEPLOC" Then
        Call WarpUserChar(userindex, UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY, True)
        Exit Sub
    ElseIf UCase$(rdata) = "/TRAMPA" Then
        Call ActivarTrampa(userindex)
        Exit Sub
    End If
End If

If UserList(userindex).flags.PuedeDenunciar Or UserList(userindex).flags.Privilegios > 0 Then
    If UCase$(Left$(rdata, 11)) = "/DENUNCIAS " Then
        rdata = Right$(rdata, Len(rdata) - 11)
        TIndex = NameIndex(rdata)
        
        If TIndex > 0 Then
            Call SendData(ToIndex, userindex, 0, "||Denuncias por cheat: " & UserList(TIndex).flags.Denuncias & "." & FONTTYPE_INFO)
            Call SendData(ToIndex, userindex, 0, "||Denuncias por insultos: " & UserList(TIndex).flags.DenunciasInsultos & "." & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, userindex, 0, "1A")
        End If
        Exit Sub
    End If

    If UCase$(Left$(rdata, 6)) = "/DENC " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        TIndex = NameIndex(rdata)
        
        If TIndex > 0 Then
            UserList(TIndex).flags.Denuncias = UserList(TIndex).flags.Denuncias + 1
            Call SendData(ToIndex, userindex, 0, "||Sumaste una denuncia por cheat a " & UserList(TIndex).Name & ". El usuario tiene acumuladas " & UserList(TIndex).flags.Denuncias & " denuncias." & FONTTYPE_INFO)
            Call LogGM(UserList(userindex).Name, "Sumo una denuncia por cheat a " & UserList(TIndex).Name & ".", UserList(userindex).flags.Privilegios = 1)
        Else
            If Not ExistePersonaje(rdata) Then
                Call SendData(ToIndex, userindex, 0, "||El personaje está offline y no se encuentra en la base de datos." & FONTTYPE_INFO)
                Exit Sub
            End If
            Call LogGM(UserList(userindex).Name, "Sumo una denuncia por cheat a " & rdata & ".", UserList(userindex).flags.Privilegios = 1)
            Call SendData(ToIndex, userindex, 0, "||Sumaste una denuncia por cheat a " & rdata & ". El usuario tiene acumuladas " & SumarDenuncia(rdata, 1) & " denuncias." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If

    If UCase$(Left$(rdata, 6)) = "/DENI " Then
        rdata = Right$(rdata, Len(rdata) - 6)
        TIndex = NameIndex(rdata)
        
        If TIndex > 0 Then
            UserList(TIndex).flags.DenunciasInsultos = UserList(TIndex).flags.DenunciasInsultos + 1
            Call SendData(ToIndex, userindex, 0, "||Sumaste una denuncia por insultos a " & UserList(TIndex).Name & ". El usuario tiene acumuladas " & UserList(TIndex).flags.DenunciasInsultos & " denuncias." & FONTTYPE_INFO)
            Call LogGM(UserList(userindex).Name, "Sumo una denuncia por insultos a " & UserList(TIndex).Name & ".", UserList(userindex).flags.Privilegios = 1)
        Else
            If Not ExistePersonaje(rdata) Then
                Call SendData(ToIndex, userindex, 0, "||El personaje está offline y no se encuentra en la base de datos." & FONTTYPE_INFO)
                Exit Sub
            End If
            Call LogGM(UserList(userindex).Name, "Sumo una denuncia por insultos a " & rdata & ".", UserList(userindex).flags.Privilegios = 1)
            Call SendData(ToIndex, userindex, 0, "||Sumaste una denuncia por insultos a " & rdata & ". El usuario tiene acumuladas " & SumarDenuncia(rdata, 2) & " denuncias." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
End If

If UserList(userindex).flags.Privilegios = 0 Then Exit Sub

If UCase$(Left$(rdata, 4)) = "/EWA" Then
If TransFacc = True Then 'Esta Activada
TransFacc = False 'Desactiva
Call SendData(ToIndex, userindex, 0, "||No se transportaran por faccion los personajes." & FONTTYPE_INFO)
Else
TransFacc = True 'Activa
Call SendData(ToIndex, userindex, 0, "||Transportacion por faccion activada." & FONTTYPE_INFO)
End If
End If

If UCase$(Left$(rdata, 6)) = "/RMSG " Then
rdata = Right$(rdata, Len(rdata) - 6)
If UserList(userindex).flags.Privilegios = 0 Then Exit Sub
Call LogGM(UserList(userindex).Name, "Mensaje Broadcast:" & rdata, False)
If rdata <> "" Then
 Call SendData(ToAll, userindex, 0, "/O" & UserList(userindex).Name & "> " & rdata)
End If
Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/CONSOLA " Then
rdata = Right$(rdata, Len(rdata) - 9)
If UserList(userindex).flags.Privilegios = 0 Then Exit Sub
Call LogGM(UserList(userindex).Name, "Mensaje Broadcast:" & rdata, False)
If rdata <> "" Then
Call SendData(ToAll, 0, 0, "|$" & UserList(userindex).Name & "> " & rdata)
End If
Exit Sub
End If

If UCase$(Left$(rdata, 4)) = "/GO " Then
    rdata = Right$(rdata, Len(rdata) - 4)
    mapa = val(ReadField(1, rdata, 32))
    If Not MapaValido(mapa) Then Exit Sub
    Call WarpUserChar(userindex, mapa, 50, 50, True)
    Call SendData(ToIndex, userindex, 0, "2B" & UserList(userindex).Name)
    Call LogGM(UserList(userindex).Name, "Transporto a " & UserList(userindex).Name & " hacia " & "Mapa" & mapa & " X:" & X & " Y:" & Y, (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(Left$(rdata, 10)) = "/DARPUNTO " Then
    rdata = Right$(rdata, Len(rdata) - 10)
    TIndex = UserList(userindex).flags.TargetUser
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "||Debes seleccionar al Jugador para Darle sus Puntos!" & FONTTYPE_INFO)
        Exit Sub
    End If
    If rdata < 0 Then
    Call SendData(ToIndex, userindex, 0, "||No podes dar puntos negativos." & FONTTYPE_FENIX)
    Exit Sub
    End If
    If rdata >= 6 Then
        Call SendData(ToIndex, userindex, 0, "||No puedes Entregar mas de 5 Puntos de Canje" & FONTTYPE_INFO)
        Exit Sub
    End If
    Call SendData(ToAll, 0, 0, "||" & UserList(UserList(userindex).flags.TargetUser).Name & " gano " & rdata & " puntos de Canje" & FONTTYPE_FENIX)
    UserList(UserList(userindex).flags.TargetUser).flags.Canje = UserList(UserList(userindex).flags.TargetUser).flags.Canje + rdata
    Call LogGM(UserList(userindex).Name, "Puntos de Canje: " & rdata & UserList(TIndex).Name & " Map:" & UserList(userindex).POS.Map & " X:" & UserList(userindex).POS.X & " Y:" & UserList(userindex).POS.Y, False)
    Exit Sub
End If
 
If UCase$(Left$(rdata, 12)) = "/SACARPUNTO " Then
    rdata = Right$(rdata, Len(rdata) - 12)
    TIndex = UserList(userindex).flags.TargetUser
    If rdata < 0 Then
    Call SendData(ToIndex, userindex, 0, "||No podes sacar cantidades negativas" & FONTTYPE_FENIX)
    Exit Sub
    End If
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "||Debes seleccionar al Jugador para Sacarle sus Puntos!" & FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(TIndex).flags.Canje < rdata Then
        Call SendData(ToIndex, userindex, 0, "||No puedes Sacar esa Cantidad de Puntos, Genera Variable Muerta!" & FONTTYPE_INFO)
        Exit Sub
    End If
    Call SendData(ToIndex, userindex, 0, "||Le has quitado " & rdata & " puntos de canje a " & UserList(TIndex).Name & "!" & FONTTYPE_FENIX)
    UserList(UserList(userindex).flags.TargetUser).flags.Canje = UserList(UserList(userindex).flags.TargetUser).flags.Canje - rdata
    Call SendData(ToIndex, TIndex, 0, "||" & UserList(userindex).Name & " te ha quitado " & rdata & " puntos de canje!" & FONTTYPE_INFO)
    Call LogGM(UserList(userindex).Name, "Restó puntos de canje: " & rdata & UserList(TIndex).Name & " Map:" & UserList(userindex).POS.Map & " X:" & UserList(userindex).POS.X & " Y:" & UserList(userindex).POS.Y, False)
    Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/TORNEO " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    PTorneo = val(ReadField(1, rdata, 32))
    If entorneo = 0 Then
        entorneo = 1
        If FileExist(App.Path & "/logs/torneo.log", vbNormal) Then Kill (App.Path & "/logs/torneo.log")
        Call SendData(ToIndex, userindex, 0, "||Has activado el torneo" & FONTTYPE_INFO)
        Call SendData(ToAll, 0, 0, "||Torneo para " & PTorneo & " jugadores. Para ingresar /PARTICIPAR" & FONTTYPE_TALK)
    Else
        entorneo = 0
        Call SendData(ToIndex, userindex, 0, "||Has desactivado el torneo" & FONTTYPE_INFO)
        Puesto = 0
        Call SendData(ToAll, 0, 0, "||Torneo cancelado." & FONTTYPE_TALK)
    End If
    Exit Sub
End If
If UCase$(Left$(rdata, 7)) = "/ELIMINAR " Then
rdata = Right$(rdata, Len(rdata) - 7)
Dim killuser As Integer
killuser = CInt(rdata)

      If Not FileExist(CharPath & "userlist(userindex).name" & ".chr", vbNormal) Then
        Call SendData(ToIndex, userindex, 0, "||El personaje no existe" & FONTTYPE_INFO)
        Else
        Call SendData(ToIndex, userindex, 0, "||Personaje Borrado Exitosamente" & FONTTYPE_INFO)
        Kill (App.Path & "Charfile" & "killuser" & ".chr")
        End If
        Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/ATORNEO " Then
rdata = Right$(rdata, Len(rdata) - 9)
Dim torneos As Integer
torneos = CInt(rdata)
If (torneos > 0 And torneos < 6) Then Call Torneos_Inicia(userindex, torneos)

 Exit Sub
End If
If UCase$(Left$(rdata, 9)) = "/AGUERRA " Then
rdata = Right$(rdata, Len(rdata) - 9)
guerra = CInt(rdata)

If Not UserList(userindex).Stats.ELV = 45 Then
 Call SendData(ToAll, 0, 0, "||TORNEO> Debes ser level 45!" & FONTTYPE_TALK)
 Exit Sub
 End If
 UserList(userindex).flags.Enguerra = False
Exit Sub
End If


If UCase(rdata) = "/CANCELARGUERRA" Then
yagano = 0
Exit Sub
End If
  
If UCase(rdata) = "/SECAEN" Then
 
    With UserList(userindex)
    If MapInfo(.POS.Map).SeCaenItems = 0 Then
    MapInfo(.POS.Map).SeCaenItems = 1
     Call SendData(ToIndex, userindex, 0, "||Ahora se caen los items en este mapa.~255~0~255~0~0" & FONTTYPE_INFO)
    Else
    MapInfo(.POS.Map).SeCaenItems = 0
        Call SendData(ToIndex, userindex, 0, "||Ahora no se caen los items en este mapa.~255~0~255~0~0" & FONTTYPE_INFO)
    End If
    End With
    
    Exit Sub
   End If
   
If UCase(rdata) = "/CANCELARDEATH" Then
Dim C As Integer
For C = 1 To LastUser

If Hay_Death = True Then
Call SendData(ToAll, 0, 0, "||AutoDeath> Autodeath cancelada." & FONTTYPE_TALK)
Hay_Death = False
End If

If UserList(C).flags.EnDeath = True Then
Call WarpUserChar(C, 1, 50, 50)
End If

UserList(C).flags.EnDeath = False
UserList(userindex).flags.EnDeath = False

Death_Cantidad = 0
Death_Muertos = 0
Next
Exit Sub
End If



If UCase(rdata) = "/CANCELAR" Then
Call Rondas_Cancela
Exit Sub
End If


If UCase$(Left$(rdata, 13)) = "/ENVENCUESTA " Then
    If encuestas.activa = 1 Then Call SendData(ToIndex, userindex, 0, "||Ya hay una encuesta, espera a que termine.." & FONTTYPE_INFO)
    rdata = Right$(rdata, Len(rdata) - 13)
    
    encuestas.votosNP = 0
    encuestas.votosSI = 0
    encuestas.Tiempo = 0
    encuestas.activa = 1
    Call SendData(ToAll, 0, 0, "||ENCUESTA> " & rdata & FONTTYPE_MATUTE)
    Call SendData(ToAll, 0, 0, "||OPCIONES: /VOTSI - /VOTNO | La encuesta durará 30 segundos." & FONTTYPE_MATUTE)
    Exit Sub
End If


If UCase$(Left$(rdata, 9)) = "/HACERDT " Then
rdata = Right$(rdata, Len(rdata) - 9)

Death = CInt(rdata)

If Hay_Death = True Then
 Call SendData(ToIndex, userindex, 0, "||AutoDeath> Lo siento, ya hay un DeathMatch en curso. Espera a que finalize.~255~0~255~0~0" & FONTTYPE_INFO)
 Exit Sub
End If

 Call SendData(ToIndex, userindex, 0, "||AutoDeath> Has iniciado el DeathMatch.~255~0~255~0~0" & FONTTYPE_INFO)
 Call SendData(ToAll, 0, 0, "||AutoDeath> El GameMaster " & UserList(userindex).Name & " ha creado un DeathMatch, para ingresar /INGRESAR~255~0~255~0~0" & FONTTYPE_INFO)
 Call SendData(ToAll, 0, 0, "||AutoDeath> CUPOS: " & Death & " ~255~0~255~0~0" & FONTTYPE_INFO)
 
 Hay_Death = True

Call SendData(ToAll, 0, 0, "TW48")
 Exit Sub
 End If

If UCase$(Left$(rdata, 9)) = "/HACERQT " Then
rdata = Right$(rdata, Len(rdata) - 9)

Questt = val(ReadField(1, rdata, 32))

If hay_Quest = True Then
 Call SendData(ToIndex, userindex, 0, "||Quest> Lo siento, ya hay una Quest en curso. Espera a que finalize.~255~0~255~0~0" & FONTTYPE_INFO)
 Exit Sub
End If

 Call SendData(ToIndex, userindex, 0, "||Quest> Has iniciado la Quest." & FONTTYPE_INFO)
 Call SendData(ToAll, 0, 0, "||Quest> El GameMaster " & UserList(userindex).Name & " ha iniciado una Quest, para ingresar /QUEST~255~0~255~0~0") 'que mal haces los msj xD AS
 hay_Quest = True

yaganoo = False
Call SendData(ToAll, 0, 0, "TW48")
 Exit Sub
 End If
If UCase$(Left$(rdata, 10)) = "/ADVERTIR " Then
rdata = Right$(rdata, Len(rdata) - 10)
    TIndex = NameIndex(rdata)
   
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    End If
 
    UserList(TIndex).Stats.advertencias = UserList(TIndex).Stats.advertencias + 1
    Call SendData(ToAll, userindex, 0, "||Advertencias> " & UserList(TIndex).Name & " ha sido advertido por " & UserList(userindex).Name & ", con esta ya lleva: " & UserList(TIndex).Stats.advertencias & " advertencias." & FONTTYPE_FIGHT)
    Call SendData(ToIndex, TIndex, 0, "||Adevertencias> Recuerda que a las 6 advertencias acumuladas, serás encarcelado 30 minutos o baneado." & FONTTYPE_INFO)
    If UserList(TIndex).Stats.advertencias = 6 Then
    'con carcel
    Call Encarcelar(TIndex, 30) '30 minutos
    'con ban
    'UserList(TIndex).flags.Ban = 1
    'Call CloseSocket(TIndex)
    UserList(TIndex).Stats.advertencias = 0
    End If
    Exit Sub
End If
If UCase$(Left$(rdata, 13)) = "/VERPROCESOS " Then
rdata = Right$(rdata, Len(rdata) - 13)
TIndex = NameIndex(rdata)
If TIndex <= 0 Then
Call SendData(ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
Else
Call SendData(ToIndex, TIndex, 0, "PCGR" & userindex)
End If
Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/SUM " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    TIndex = NameIndex(rdata)
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    
    If UserList(userindex).flags.Privilegios < UserList(TIndex).flags.Privilegios And UserList(TIndex).flags.AdminInvisible = 1 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    
    If UserList(userindex).flags.Privilegios = 0 And UserList(TIndex).POS.Map <> UserList(userindex).POS.Map Then Exit Sub
    
    Call SendData(ToIndex, userindex, 0, "%Z" & UserList(TIndex).Name)
    Call WarpUserChar(TIndex, UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y + 1, True)
    
    Call LogGM(UserList(userindex).Name, "/SUM " & UserList(TIndex).Name & " Map:" & UserList(userindex).POS.Map & " X:" & UserList(userindex).POS.X & " Y:" & UserList(userindex).POS.Y, False)
    Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/IRA " Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    TIndex = NameIndex(rdata)
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    
    If ((UserList(TIndex).flags.Privilegios > UserList(userindex).flags.Privilegios And UserList(TIndex).flags.AdminInvisible = 1)) Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If

    If UserList(TIndex).flags.AdminInvisible And Not UserList(userindex).flags.AdminInvisible Then Call DoAdminInvisible(userindex)

    Call WarpUserChar(userindex, UserList(TIndex).POS.Map, UserList(TIndex).POS.X + 1, UserList(TIndex).POS.Y + 1, True)
    
    Call LogGM(UserList(userindex).Name, "/IRA " & UserList(TIndex).Name & " Mapa:" & UserList(TIndex).POS.Map & " X:" & UserList(TIndex).POS.X & " Y:" & UserList(TIndex).POS.Y, (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(rdata) = "/INVISIBLE" Then
    Call DoAdminInvisible(userindex)
    Call LogGM(UserList(userindex).Name, "/INVISIBLE", (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(rdata) = "/PANELGM" Then
    Call SendData(ToIndex, userindex, 0, "PGM" & UserList(userindex).flags.Privilegios)
    Exit Sub
End If

If UCase$(rdata) = "/TRABAJANDO" Then
    For LoopC = 1 To LastUser
        If Len(UserList(LoopC).Name) > 0 And UserList(LoopC).flags.Trabajando Then
            DummyInt = DummyInt + 1
            tStr = tStr & UserList(LoopC).Name & ", "
        End If
    Next
    If Len(tStr) > 0 Then
        tStr = Left$(tStr, Len(tStr) - 2)
        Call SendData(ToIndex, userindex, 0, "||Usuarios trabajando: " & tStr & FONTTYPE_INFO)
        Call SendData(ToIndex, userindex, 0, "||Número de usuarios trabajando: " & DummyInt & "." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, userindex, 0, "%)")
    End If
    Exit Sub
End If
If UCase$(Left$(rdata, 8)) = "/CARCEL " Then
    
    rdata = Right$(rdata, Len(rdata) - 8)
    
    Name = ReadField(1, rdata, 32)
    i = val(ReadField(1, rdata, 32))
    Name = Right$(rdata, Len(rdata) - (Len(Name) + 1))
    
    TIndex = NameIndex(Name)
    
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    
    If UserList(TIndex).flags.Privilegios > UserList(userindex).flags.Privilegios Then
        Call SendData(ToIndex, userindex, 0, "1B")
        Exit Sub
    End If
    
    If i > 120 Then
        Call SendData(ToIndex, userindex, 0, "1C")
        Exit Sub
    End If
    
    Call Encarcelar(TIndex, i, UserList(userindex).Name)
    
    Exit Sub
End If

If UCase$(rdata) = "/TELEPLOC" Then
    Call WarpUserChar(userindex, UserList(userindex).flags.TargetMap, UserList(userindex).flags.TargetX, UserList(userindex).flags.TargetY, True)
    Call LogGM(UserList(userindex).Name, "/TELEPLOC a x:" & UserList(userindex).flags.TargetX & " Y:" & UserList(userindex).flags.TargetY & " Map:" & UserList(userindex).POS.Map, (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If


If UserList(userindex).flags.Privilegios < 2 Then Exit Sub

If UCase$(Left$(rdata, 4)) = "/REM" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    Call LogGM(UserList(userindex).Name, "Comentario: " & rdata, (UserList(userindex).flags.Privilegios = 1))
    Call SendData(ToIndex, userindex, 0, "||Comentario salvado..." & FONTTYPE_INFO)
    Exit Sub
End If


If UCase$(Left$(rdata, 5)) = "/HORA" Then
    Call LogGM(UserList(userindex).Name, "Hora.", (UserList(userindex).flags.Privilegios = 1))
    rdata = Right$(rdata, Len(rdata) - 5)
    Call SendData(ToAll, 0, 0, "||Hora: " & Time & " " & Date & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(rdata) = "/ONLINEGM" Then
        For LoopC = 1 To LastUser
            If Len(UserList(LoopC).Name) > 0 Then
                If UserList(LoopC).flags.Privilegios > 0 And (UserList(LoopC).flags.Privilegios <= UserList(userindex).flags.Privilegios Or UserList(LoopC).flags.AdminInvisible = 0) Then
                    tStr = tStr & UserList(LoopC).Name & ", "
                End If
            End If
        Next
        If Len(tStr) > 0 Then
            tStr = Left$(tStr, Len(tStr) - 2)
            Call SendData(ToIndex, userindex, 0, "|| Usuarios online: " & tStr & ". Record de usuarios: " & recordusuarios & FONTTYPE_INFO)
        Else
            Call SendData(ToIndex, userindex, 0, "%P")
        End If
        Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/DONDE " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    TIndex = NameIndex(rdata)
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    
    If UserList(TIndex).flags.Privilegios > UserList(userindex).flags.Privilegios And UserList(TIndex).flags.AdminInvisible = 1 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    
    Call SendData(ToIndex, userindex, 0, "||Ubicacion de " & UserList(TIndex).Name & ": " & UserList(TIndex).POS.Map & ", " & UserList(TIndex).POS.X & ", " & UserList(TIndex).POS.Y & "." & FONTTYPE_INFO)
    Call LogGM(UserList(userindex).Name, "/Donde", (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/NENE " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    If MapaValido(val(rdata)) Then
        Call SendData(ToIndex, userindex, 0, "NENE" & NPCHostiles(val(rdata)))
        Call LogGM(UserList(userindex).Name, "Numero enemigos en mapa " & rdata, (UserList(userindex).flags.Privilegios = 1))
    End If
    Exit Sub
End If

If UCase$(rdata) = "/VENTAS" Then
    Call SendData(ToIndex, userindex, 0, "/X" & DineroTotalVentas & "," & NumeroVentas)
    Exit Sub
End If

If UCase$(rdata) = "/DESCONGELAR" Then
    Call Congela(True)
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/VIGILAR " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    TIndex = NameIndex(rdata)
    If TIndex > 0 Then
        If TIndex = userindex Then
            Call SendData(ToIndex, userindex, 0, "||No puedes vigilarte a ti mismo." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(TIndex).flags.Privilegios >= UserList(userindex).flags.Privilegios Then
            Call SendData(ToIndex, userindex, 0, "||No puedes vigilar a alguien con igual o mayor jerarquia que tú." & FONTTYPE_INFO)
            Exit Sub
        End If
        If YaVigila(TIndex, userindex) Then
            Call SendData(ToIndex, userindex, 0, "||Dejaste de vigilar a " & UserList(TIndex).Name & "." & FONTTYPE_INFO)
            If Not EsVigilado(TIndex) Then Call SendData(ToIndex, TIndex, 0, "VIG")
            Exit Sub
        End If
        If Not EsVigilado(TIndex) Then Call SendData(ToIndex, TIndex, 0, "VIG")
        Call SendData(ToIndex, userindex, 0, "||Estás vigilando a " & UserList(TIndex).Name & "." & FONTTYPE_INFO)
        For i = 1 To 10
            If UserList(TIndex).flags.Espiado(i) = 0 Then
                UserList(TIndex).flags.Espiado(i) = userindex
                Exit For
            End If
        Next
        If i = 11 Then
            Call SendData(ToIndex, userindex, 0, "||Demasiados GM's están vigilando a este usuario." & FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        Call SendData(ToIndex, userindex, 0, "1A")
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/VERPC " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    TIndex = NameIndex(rdata)
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    If UserList(TIndex).flags.Privilegios > UserList(userindex).flags.Privilegios And UserList(userindex).flags.AdminInvisible = 1 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    If UserList(TIndex).flags.Privilegios >= UserList(userindex).flags.Privilegios Then
        Call SendData(ToIndex, userindex, 0, "||No puedes ver la PC de un GM con mayor jerarquia." & FONTTYPE_FIGHT)
        Exit Sub
    End If
    UserList(TIndex).flags.EsperandoLista = userindex
    Call SendData(ToIndex, TIndex, 0, "VPRC")
End If

If UCase$(Left$(rdata, 7)) = "/TELEP " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    mapa = val(ReadField(2, rdata, 32))
    If Not MapaValido(mapa) Then Exit Sub
    Name = ReadField(1, rdata, 32)
    If Len(Name) = 0 Then Exit Sub
    If UCase$(Name) <> "YO" Then
        If UserList(userindex).flags.Privilegios = 1 Then
            Exit Sub
        End If
        TIndex = NameIndex(Name)
    Else
        TIndex = userindex
    End If
    X = val(ReadField(3, rdata, 32))
    Y = val(ReadField(4, rdata, 32))
    If Not InMapBounds(X, Y) Then Exit Sub
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    If UserList(TIndex).flags.Privilegios > UserList(userindex).flags.Privilegios And UserList(userindex).flags.AdminInvisible = 1 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    Call WarpUserChar(TIndex, mapa, X, Y, True)
    Call SendData(ToIndex, TIndex, 0, "||" & UserList(userindex).Name & " te ha transportado." & FONTTYPE_INFO)
    Call LogGM(UserList(userindex).Name, "Transporto a " & UserList(TIndex).Name & " hacia " & "Mapa" & mapa & " X:" & X & " Y:" & Y, (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If


If UCase$(Left$(rdata, 4)) = "/GO " Then
    rdata = Right$(rdata, Len(rdata) - 4)
    mapa = val(ReadField(1, rdata, 32))
    If Not MapaValido(mapa) Then Exit Sub
    Call WarpUserChar(userindex, mapa, 50, 50, True)
    Call SendData(ToIndex, userindex, 0, "2B" & UserList(userindex).Name)
    Call LogGM(UserList(userindex).Name, "Transporto a " & UserList(userindex).Name & " hacia " & "Mapa" & mapa & " X:" & X & " Y:" & Y, (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(rdata) = "/CMAP" Then
    If MapInfo(UserList(userindex).POS.Map).NumUsers Then
        Call SendData(ToIndex, userindex, 0, "||Hay " & MapInfo(UserList(userindex).POS.Map).NumUsers & " usuarios en este mapa." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, userindex, 0, "%R")
    End If

    Exit Sub
End If

If UCase$(Left$(rdata, 10)) = "/VERTORNEO" Then
    Dim stri As String
    Dim jugadores As Integer
    Dim jugador As Integer
    stri = ""
    jugadores = val(GetVar(App.Path & "/logs/torneo.log", "CANTIDAD", "CANTIDAD"))
    For jugador = 1 To jugadores
        stri = stri & GetVar(App.Path & "/logs/torneo.log", "JUGADORES", "JUGADOR" & jugador) & ","
    Next
    Call SendData(ToIndex, userindex, 0, "||Quieren participar: " & stri & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/INFO " Then
    Call LogGM(UserList(userindex).Name, rdata, False)
    
    rdata = Right$(rdata, Len(rdata) - 6)
    
    TIndex = NameIndex(rdata)
    
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If

    SendUserSTAtsTxt userindex, TIndex
    Call SendData(ToIndex, userindex, 0, "||Mail: " & UserList(TIndex).Email & FONTTYPE_INFO)
    Call SendData(ToIndex, userindex, 0, "||Ip: " & UserList(TIndex).ip & FONTTYPE_INFO)

    Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/IPNICK " Then
    rdata = Right$(rdata, Len(rdata) - 8)






    tStr = ""
    For LoopC = 1 To LastUser
        If UserList(LoopC).ip = rdata And Len(UserList(LoopC).Name) > 0 And UserList(LoopC).flags.UserLogged Then
            If (UserList(userindex).flags.Privilegios > 0 And UserList(LoopC).flags.Privilegios = 0) Or (UserList(userindex).flags.Privilegios = 3) Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        End If
    Next
    Call SendData(ToIndex, userindex, 0, "||Los personajes con ip " & rdata & " son: " & tStr & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(Left$(rdata, 10)) = "/MAILNICK " Then
    rdata = Right$(rdata, Len(rdata) - 10)






    tStr = ""
    For LoopC = 1 To LastUser
        If UCase$(UserList(LoopC).ip) = UCase$(rdata) And Len(UserList(LoopC).Name) > 0 And UserList(LoopC).flags.UserLogged Then
            If (UserList(userindex).flags.Privilegios > 0 And UserList(LoopC).flags.Privilegios = 0) Or (UserList(userindex).flags.Privilegios = 3) Then
                tStr = tStr & UserList(LoopC).Name & ", "
            End If
        End If
    Next
    Call SendData(ToIndex, userindex, 0, "||Los personajes con mail " & rdata & " son: " & tStr & FONTTYPE_INFO)
    Exit Sub
End If


If UCase$(Left$(rdata, 5)) = "/INV " Then
    Call LogGM(UserList(userindex).Name, rdata, False)
    
    rdata = Right$(rdata, Len(rdata) - 5)
    
    TIndex = NameIndex(rdata)
    
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If

    SendUserInvTxt userindex, TIndex
    Exit Sub
End If


If UCase$(Left$(rdata, 8)) = "/SKILLS " Then
    Call LogGM(UserList(userindex).Name, rdata, False)
    
    rdata = Right$(rdata, Len(rdata) - 8)
    
    TIndex = NameIndex(rdata)
    
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If

    SendUserSkillsTxt userindex, TIndex
    Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/ATR " Then
    Call LogGM(UserList(userindex).Name, rdata, False)
    
    rdata = Right$(rdata, Len(rdata) - 5)
    
    TIndex = NameIndex(rdata)
    
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If

    Call SendData(ToIndex, userindex, 0, "||Atributos de " & UserList(TIndex).Name & FONTTYPE_INFO)
    For i = 1 To NUMATRIBUTOS
        Call SendData(ToIndex, userindex, 0, "|| " & AtributosNames(i) & " = " & UserList(TIndex).Stats.UserAtributosBackUP(1) & FONTTYPE_INFO)
    Next
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/REVIVIR " Then
    rdata = Right$(rdata, Len(rdata) - 9)
    Name = rdata
    If UCase$(Name) <> "YO" Then
        TIndex = NameIndex(Name)
    Else
        TIndex = userindex
    End If
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    Call RevivirUsuarioNPC(TIndex)
    Call SendData(ToIndex, TIndex, 0, "%T" & UserList(userindex).Name)
    Call LogGM(UserList(userindex).Name, "Resucito a " & UserList(TIndex).Name, False)
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/BANT " Then
    rdata = Right$(rdata, Len(rdata) - 6)
 
    Arg1 = ReadField(1, rdata, 64)
    Name = ReadField(2, rdata, 64)
    i = val(ReadField(3, rdata, 64))
    
    If Len(Arg1) = 0 Or Len(Name) = 0 Or i = 0 Then
        Call SendData(ToIndex, userindex, 0, "||La estructura del comando es /BANT CAUSA@NICK@DIAS." & FONTTYPE_FENIX)
        Exit Sub
    End If
    
    TIndex = NameIndex(Name)
    
    If TIndex > 0 Then
        If UserList(TIndex).flags.Privilegios > UserList(userindex).flags.Privilegios Then
            Call SendData(ToIndex, userindex, 0, "1B")
            Exit Sub
        End If
        
        Call BanTemporal(Name, i, Arg1, UserList(userindex).Name)
        Call SendData(ToAdmins, 0, 0, "%X" & UserList(userindex).Name & "," & UserList(TIndex).Name)
        
        UserList(TIndex).flags.Ban = 1
        Call WarpUserChar(TIndex, ULLATHORPE.Map, ULLATHORPE.X, ULLATHORPE.Y)
        
        Call CloseSocket(TIndex)
    Else
        If Not FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = False Then
            Call SendData(ToIndex, userindex, 0, "||Offline, baneando" & FONTTYPE_INFO)
            
            If GetVar(CharPath & Name & ".chr", "FLAGS", "Ban") <> "0" Then
                Call SendData(ToIndex, userindex, 0, "||El personaje ya se encuentra baneado." & FONTTYPE_INFO)
                Exit Sub
            End If
            
            Call BanTemporal(Name, i, Arg1, UserList(userindex).Name)
            
            Call ChangeBan(Name, 1)
            Call ChangePos(Name)
            
            Call SendData(ToAdmins, 0, 0, "%X" & UserList(userindex).Name & "," & Name)
        Else
            Call SendData(ToIndex, userindex, 0, "||El usuario no existe." & FONTTYPE_INFO)
        End If
    End If
 
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/ECHAR " Then

    rdata = Right$(rdata, Len(rdata) - 7)
    TIndex = NameIndex(rdata)
If UserList(userindex).Name = "Dokha" Then
Call SendData(ToIndex, userindex, 0, "||¡¡Sos muy pete!!" & FONTTYPE_TALK)
   Exit Sub
   End If
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1E")
        Exit Sub
    End If
    
    If TIndex = userindex Then Exit Sub
    
    If UserList(TIndex).flags.Privilegios > UserList(userindex).flags.Privilegios Then
        Call SendData(ToIndex, userindex, 0, "1F")
        Exit Sub
    End If
        
    Call SendData(ToAdmins, 0, 0, "%U" & UserList(userindex).Name & "," & UserList(TIndex).Name)
    Call LogGM(UserList(userindex).Name, "Echo a " & UserList(TIndex).Name, False)
    Call CloseSocket(TIndex)
    Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/BAN " Then
    Dim Razon As String
    rdata = Right$(rdata, Len(rdata) - 5)
    Razon = ReadField(1, rdata, Asc("@"))
    Name = ReadField(2, rdata, Asc("@"))
    TIndex = NameIndex(Name)
    '/ban motivo@nombre
    If TIndex Then
        If TIndex = userindex Then Exit Sub
        Name = UserList(TIndex).Name
        If UserList(TIndex).flags.Privilegios > UserList(userindex).flags.Privilegios Then
            Call SendData(ToIndex, userindex, 0, "%V")
            Exit Sub
        End If
        
        Call LogBan(TIndex, userindex, Razon)
        UserList(TIndex).flags.Ban = 1
        
        If UserList(TIndex).flags.Privilegios Then
            UserList(userindex).flags.Ban = 1
            Call SendData(ToAdmins, 0, 0, "%W" & UserList(userindex).Name)
            Call LogBan(userindex, userindex, "Baneado por banear a otro GM.")
            Call CloseSocket(userindex)
        End If
        
        Call SendData(ToAdmins, 0, 0, "%X" & UserList(userindex).Name & "," & UserList(TIndex).Name)
        Call SendData(ToAdmins, 0, 0, "||IP: " & UserList(TIndex).ip & " Mail: " & UserList(TIndex).Email & "." & FONTTYPE_FIGHT)
 
        Call CloseSocket(TIndex)
    Else
        If Not FileExist(CharPath & UCase$(Name) & ".chr", vbNormal) = False Then
            Call ChangeBan(Name, 1)
            Call LogBanOffline(UCase$(Name), userindex, Razon)
            Call SendData(ToAdmins, 0, 0, "%X" & UserList(userindex).Name & "," & Name)
        Else
            Call SendData(ToIndex, userindex, 0, "||El usuario no existe." & FONTTYPE_INFO)
        End If
        
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/UNBAN " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    If Not FileExist(CharPath & UCase$(rdata) & ".chr", vbNormal) = False Then
        Call ChangeBan(rdata, 0)
        Call SendData(ToIndex, userindex, 0, "||" & rdata & " unbanned." & FONTTYPE_INFO)
        For i = 1 To Baneos.Count
            If Baneos(i).Name = UCase$(rdata) Then
                Call Baneos.Remove(i)
                Exit Sub
            End If
        Next
    Else
        Call SendData(ToIndex, userindex, 0, "||El usuario no existe" & FONTTYPE_INFO)
    End If
    Exit Sub
End If





If UCase$(rdata) = "/SEGUIR" Then
    If UserList(userindex).flags.TargetNpc Then
        Call DoFollow(UserList(userindex).flags.TargetNpc, userindex)
    End If
    Exit Sub
End If


If UCase$(Left$(rdata, 3)) = "/CC" Then
   Call EnviarSpawnList(userindex)
   Exit Sub
End If


If UCase$(Left$(rdata, 3)) = "SPA" Then
    rdata = Right$(rdata, Len(rdata) - 3)
    
    If val(rdata) > 0 And val(rdata) < UBound(SpawnList) + 1 Then _
          Call SpawnNpc(SpawnList(val(rdata)).NpcIndex, UserList(userindex).POS, True, False)
          
          Call LogGM(UserList(userindex).Name, "Sumoneo " & SpawnList(val(rdata)).NpcName, False)
          
    Exit Sub
End If

If UCase$(rdata) = "/RESETINV" Then
    rdata = Right$(rdata, Len(rdata) - 9)
    If UserList(userindex).flags.TargetNpc = 0 Then Exit Sub
    Call ResetNpcInv(UserList(userindex).flags.TargetNpc)
    Call LogGM(UserList(userindex).Name, "/RESETINV " & Npclist(UserList(userindex).flags.TargetNpc).Name, False)
    Exit Sub
End If


If UCase$(rdata) = "/LIMPIAR" Then
    Call LimpiarMundo
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/RMSGT " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    If UCase$(rdata) = "NO" Then
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " ha anulado la repetición del mensaje: " & MensajeRepeticion & "." & FONTTYPE_FENIX)
        IntervaloRepeticion = 0
        TiempoRepeticion = 0
        MensajeRepeticion = ""
        Exit Sub
    End If
    tName = ReadField(1, rdata, 64)
    tInt = ReadField(2, rdata, 64)
    Prueba1 = ReadField(3, rdata, 64)
    If Len(tName) = 0 Or val(Prueba1) = 0 Or (Prueba1 >= tInt And tInt <> 0) Then
        Call SendData(ToIndex, userindex, 0, "||La estructura del comando es: /RMSGT MENSAJE@TIEMPO TOTAL@INTERVALO DE REPETICION." & FONTTYPE_INFO)
        Exit Sub
    End If
    If val(tInt) > 10000 Or val(Prueba1) > 10000 Then
        Call SendData(ToIndex, userindex, 0, "||La cantidad de tiempo establecida es demasiado grande." & FONTTYPE_INFO)
        Exit Sub
    End If
    Call LogGM(UserList(userindex).Name, "Mensaje Broadcast repetitivo:" & rdata, False)
    MensajeRepeticion = tName
    TiempoRepeticion = tInt
    IntervaloRepeticion = Prueba1
    If TiempoRepeticion = 0 Then
        Call SendData(ToAdmins, 0, 0, "||El mensaje " & MensajeRepeticion & " será repetido cada " & IntervaloRepeticion & " minutos durante tiempo indeterminado." & FONTTYPE_FENIX)
        TiempoRepeticion = -IntervaloRepeticion
    Else
        Call SendData(ToAdmins, 0, 0, "||El mensaje " & MensajeRepeticion & " será repetido cada " & IntervaloRepeticion & " minutos durante un total de " & TiempoRepeticion & " minutos." & FONTTYPE_FENIX)
        TiempoRepeticion = TiempoRepeticion - TiempoRepeticion Mod IntervaloRepeticion
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/BUSCAR " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    For i = 1 To UBound(ObjData)
        If InStr(1, Tilde(ObjData(i).Name), Tilde(rdata)) Then
            Call SendData(ToIndex, userindex, 0, "||" & i & " " & ObjData(i).Name & "." & FONTTYPE_INFO)
            N = N + 1
        End If
    Next
    If N = 0 Then
        Call SendData(ToIndex, userindex, 0, "||No hubo resultados de la búsqueda: " & rdata & "." & FONTTYPE_INFO)
    Else
        Call SendData(ToIndex, userindex, 0, "||Hubo " & N & " resultados de la busqueda: " & rdata & "." & FONTTYPE_INFO)
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/CUENTA " Then
    rdata = Right$(rdata, Len(rdata) - 8)
    CuentaRegresiva = val(ReadField(1, rdata, 32)) + 1
    GMCuenta = UserList(userindex).POS.Map
    Exit Sub
End If


If UCase$(rdata) = "/MATA" Then
    If UserList(userindex).flags.TargetNpc = 0 Then Exit Sub
    Call QuitarNPC(UserList(userindex).flags.TargetNpc)
    Call LogGM(UserList(userindex).Name, "/MATA " & Npclist(UserList(userindex).flags.TargetNpc).Name, False)
    Exit Sub
End If

If UCase$(rdata) = "/MUERE" Then
    If UserList(userindex).flags.TargetNpc = 0 Then Exit Sub
    Call MuereNpc(UserList(userindex).flags.TargetNpc, userindex)
    Call LogGM(UserList(userindex).Name, "/MUERE " & Npclist(UserList(userindex).flags.TargetNpc).Name, False)
    Exit Sub
End If

If UCase$(rdata) = "/IGNORAR" Then
    If UserList(userindex).flags.Ignorar = 1 Then
        UserList(userindex).flags.Ignorar = 0
        Call SendData(ToIndex, userindex, 0, "||Ahora las criaturas te persiguen." & FONTTYPE_INFO)
    Else
        UserList(userindex).flags.Ignorar = 1
        Call SendData(ToIndex, userindex, 0, "||Ahora las criaturas te ignoran." & FONTTYPE_INFO)
    End If
End If


If UCase$(Left$(rdata, 8)) = "/NOMBRE " Then
    Dim NewNick As String
    rdata = Right$(rdata, Len(rdata) - 8)
    TIndex = NameIndex(ReadField(1, rdata, Asc(" ")))
    NewNick = Right$(rdata, Len(rdata) - (Len(ReadField(1, rdata, Asc(" "))) + 1))
    If Len(NewNick) = 0 Then Exit Sub
    If TIndex = 0 Then
        Call SendData(ToIndex, userindex, 0, "$3E")
        Exit Sub
    End If
    If FileExist(CharPath & UCase$(NewNick) & ".chr", vbNormal) Then
        Call SendData(ToIndex, userindex, 0, "||El nombre ya existe, elige otro." & FONTTYPE_INFO)
    Else
    Call ReNombrar(TIndex, NewNick)
End If
Exit Sub
End If
If UCase$(Left$(rdata, 12)) = "/VERCAPTION " Then
rdata = Right$(rdata, Len(rdata) - 12)
TIndex = NameIndex(rdata)
If TIndex <= 0 Then
Call SendData(ToIndex, userindex, 0, "||Usuario offline." & FONTTYPE_INFO)
Else
Call SendData(ToIndex, TIndex, 0, "PCCP" & userindex)
End If
Exit Sub
End If
If UCase$(Left$(rdata, 5)) = "/DEST" Then
    Call LogGM(UserList(userindex).Name, "/DEST", False)
    rdata = Right$(rdata, Len(rdata) - 5)
    Call EraseObj(ToMap, userindex, UserList(userindex).POS.Map, 10000, UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y)
    Exit Sub
End If

If UCase$(rdata) = "/MASSDEST" Then
    For Y = UserList(userindex).POS.Y - MinYBorder + 1 To UserList(userindex).POS.Y + MinYBorder - 1
        For X = UserList(userindex).POS.X - MinXBorder + 1 To UserList(userindex).POS.X + MinXBorder - 1
            If InMapBounds(X, Y) Then _
            If MapData(UserList(userindex).POS.Map, X, Y).OBJInfo.OBJIndex > 0 And Not ItemEsDeMapa(UserList(userindex).POS.Map, X, Y) Then Call EraseObj(ToMap, userindex, UserList(userindex).POS.Map, 10000, UserList(userindex).POS.Map, X, Y)
        Next
    Next
    Call LogGM(UserList(userindex).Name, "/MASSDEST", (UserList(userindex).flags.Privilegios = 1))
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/KILL " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    TIndex = NameIndex(rdata)
    If TIndex Then
        If UserList(TIndex).flags.Privilegios < UserList(userindex).flags.Privilegios Then Call UserDie(TIndex)
    End If
    Exit Sub
End If


If UCase$(Left$(rdata, 11)) = "/GANADOR" Then
    rdata = Right$(rdata, Len(rdata) - 8)
    TIndex = UserList(userindex).flags.TargetUser
    If TIndex <= 0 Then
    Call SendData(ToIndex, userindex, 0, "||Primero selecciona a un jugador!" & FONTTYPE_INFO)
    Exit Sub
    End If
    Call SendData(ToAll, 0, 0, "TW44")
    Call SendData(ToAll, 0, 0, "||" & UserList(UserList(userindex).flags.TargetUser).Name & " ganó el torneo." & FONTTYPE_TALK)
    Call SendData(ToAll, 0, 0, "||Se lleva como recompensa:" & FONTTYPE_TALK)
    Call SendData(ToAll, 0, 0, "||3 puntos de CANJE! + 1 Torneo ganado." & FONTTYPE_TALK) 'cambian el 3 por la cantidad que deseen.
    
    UserList(UserList(userindex).flags.TargetUser).Faccion.torneos = UserList(UserList(userindex).flags.TargetUser).Faccion.torneos + 1
    UserList(UserList(userindex).flags.TargetUser).flags.Canje = UserList(UserList(userindex).flags.TargetUser).flags.Canje + 3 'cambien este otro 3 por la cantidad de puntos que deseen
    Call LogGM(UserList(userindex).Name, "Ganador de torneo: " & UserList(TIndex).Name & " Map:" & UserList(userindex).POS.Map & " X:" & UserList(userindex).POS.X & " Y:" & UserList(userindex).POS.Y, False)
    Exit Sub
End If


If UCase$(Left$(rdata, 10)) = "/GANOQUEST" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    TIndex = UserList(userindex).flags.TargetUser
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "||Debes seleccionar a un jugador!" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    Call SendData(ToAll, 0, 0, "||" & UserList(UserList(userindex).flags.TargetUser).Name & " ganó una quest." & FONTTYPE_INFO)
    UserList(UserList(userindex).flags.TargetUser).Faccion.Quests = UserList(UserList(userindex).flags.TargetUser).Faccion.Quests + 1
    Call LogGM(UserList(userindex).Name, "Ganó quest: " & UserList(TIndex).Name & " Map:" & UserList(userindex).POS.Map & " X:" & UserList(userindex).POS.X & " Y:" & UserList(userindex).POS.Y, False)
    Exit Sub
End If

If UCase$(Left$(rdata, 13)) = "/PERDIOTORNEO" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    TIndex = UserList(userindex).flags.TargetUser
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "||Debes seleccionar a un jugador!" & FONTTYPE_INFO)
        Exit Sub
    End If
    
    UserList(UserList(userindex).flags.TargetUser).Faccion.torneos = UserList(UserList(userindex).flags.TargetUser).Faccion.torneos - 1
    
    Call LogGM(UserList(userindex).Name, "Restó torneo: " & UserList(TIndex).Name & " Map:" & UserList(userindex).POS.Map & " X:" & UserList(userindex).POS.X & " Y:" & UserList(userindex).POS.Y, False)
    Exit Sub
End If

If UCase$(Left$(rdata, 12)) = "/PERDIOQUEST" Then
    rdata = Right$(rdata, Len(rdata) - 5)
    
    TIndex = UserList(userindex).flags.TargetUser
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "||Debes seleccionar a un jugador!" & FONTTYPE_INFO)
        Exit Sub
    End If
    




If UCase$(rdata) = "/RESTRINGIR" Then
    If Restringido Then
        Call SendData(ToAll, 0, 0, "||La restricción de GameMasters fue desactivada servidor." & FONTTYPE_FENIX)
        Call LogGM(UserList(userindex).Name, "Desrestringió el servidor.", False)
    Else
        Call SendData(ToAll, 0, 0, "||La restricción de GameMasters fue activada." & FONTTYPE_FENIX)
        End If
        For i = 1 To LastUser
            DoEvents
            
            If UserList(i).flags.UserLogged And UserList(i).flags.Privilegios = 0 And Not UserList(i).flags.PuedeDenunciar Then Call CloseSocket(i)
        Next
        Call LogGM(UserList(userindex).Name, "Restringió el servidor.", False)
    End If
    Restringido = Not Restringido
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/VERMC " Then
rdata = Right$(rdata, Len(rdata) - 7)
TIndex = NameIndex(rdata)
   If TIndex <= 0 Then
      Call SendData(ToIndex, userindex, 0, "||Usuario Offline." & FONTTYPE_INFO)
     Else
      Call SendData(ToIndex, userindex, 0, "||El MAC es: " & UserList(TIndex).Mac & FONTTYPE_INFO)
   End If
 
 
Exit Sub
End If
 
If UCase$(Left$(rdata, 7)) = "/BANMC " Then
rdata = Right$(rdata, Len(rdata) - 7)
TIndex = NameIndex(rdata)
If UserList(userindex).Name <> "DarkTester" Then Exit Sub
If TIndex <= 0 Then
Call SendData(ToIndex, userindex, 0, "||Usuario Offline." & FONTTYPE_INFO)
Exit Sub
Else
For LoopC = 1 To BanMACs.Count
   If BanMACs.Item(LoopC) = UserList(TIndex).Mac Then
      Call SendData(ToIndex, userindex, 0, "||MAC ya baneada" & FONTTYPE_INFO)
      Exit Sub
   End If
Next
   
BanMACs.Add UserList(TIndex).Mac
Call SendData(ToIndex, userindex, 0, "||Has baneado la MAC: " & UserList(TIndex).Mac & " del usuario " & UserList(TIndex).Name & "." & FONTTYPE_INFO)
   
   Dim numMAC As Integer
   numMAC = val(GetVar(App.Path & "\Dat\BanMAC.dat", "INIT", "Cantidad"))
 
   If FileExist(App.Path & "\Dat\BanMAC.dat", vbNormal) Then
      Call WriteVar(App.Path & "\Dat\BanMAC.dat", "INIT", "Cantidad", numMAC + 1)
      Call WriteVar(App.Path & "\Dat\BanMAC.dat", "BANS", "MAC" & numMAC + 1, UserList(TIndex).Mac)
      Call LogGM(UserList(userindex).Name, "/BanHD " & UserList(TIndex).Name & " " & UserList(TIndex).Mac, False)
   Else
      Call WriteVar(App.Path & "\Logs\BanHDs.dat", "INIT", "Cantidad", 1)
      Call WriteVar(App.Path & "\Logs\BanHDs.dat", "BANS", "MAC1", UserList(TIndex).Mac)
      Call LogGM(UserList(userindex).Name, "/BanHD " & UserList(TIndex).Name & " " & UserList(TIndex).Mac, False)
   End If
   
   Call CloseSocket(TIndex)
 
End If
Exit Sub
End If
 
If UCase$(Left$(rdata, 9)) = "/UNBANMC " Then
rdata = Right$(rdata, Len(rdata) - 9)
TIndex = NameIndex(rdata)
   
   Dim numMAC2 As Integer
   numMAC2 = val(GetVar(App.Path & "\Dat\BanMAC.dat", "INIT", "Cantidad"))
 
   For LoopC = 1 To BanMACs.Count
   If BanMACs.Item(LoopC) = UserList(userindex).Mac Then
      BanMACs.Remove LoopC
      Call SendData(ToIndex, userindex, 0, "||Has desbaneado la MAC de " & rdata & FONTTYPE_INFO)
      Call WriteVar(App.Path & "\Dat\BanMAC.dat", "INIT", "Cantidad", numMAC2 - 1)
      Call WriteVar(App.Path & "\Dat\BanMAC.dat", "BANEO", "Mac" & numMAC2 - 1, "")
      Call LogGM(UserList(userindex).Name, "/UNBanHD " & UserList(TIndex).Name & " " & UserList(TIndex).Mac, False)
   End If
   Next
 
 
Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/BANIP" Then
    Dim BanIP As String, XNick As Boolean
    
    rdata = Right$(rdata, Len(rdata) - 7)
    
    TIndex = NameIndex(rdata)
    If UserList(userindex).Name <> "DarkTester" Then Exit Sub
    If TIndex <= 0 Then
        XNick = False
        Call LogGM(UserList(userindex).Name, "/BanIP " & rdata, False)
        BanIP = rdata
    Else
        XNick = True
        Call LogGM(UserList(userindex).Name, "/BanIP " & UserList(TIndex).Name & " - " & UserList(TIndex).ip, False)
        BanIP = UserList(TIndex).ip
    End If
    
    
    For LoopC = 1 To BanIps.Count
        If BanIps.Item(LoopC) = BanIP Then
            Call SendData(ToIndex, userindex, 0, "||La IP " & BanIP & " ya se encuentra en la lista de bans." & FONTTYPE_INFO)
            Exit Sub
        End If
    Next
    
    BanIps.Add BanIP
    Call SendData(ToAdmins, userindex, 0, "||" & UserList(userindex).Name & " Baneo la IP " & BanIP & FONTTYPE_FIGHT)
    
    If XNick Then
        Call LogBan(TIndex, userindex, "Ban por IP desde Nick")
        
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " echo a " & UserList(TIndex).Name & "." & FONTTYPE_FIGHT)
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " Banned a " & UserList(TIndex).Name & "." & FONTTYPE_FIGHT)
        
        
        UserList(TIndex).flags.Ban = 1
        
        Call LogGM(UserList(userindex).Name, "Echo a " & UserList(TIndex).Name, False)
        Call LogGM(UserList(userindex).Name, "BAN a " & UserList(TIndex).Name, False)
        Call CloseSocket(TIndex)
    End If
    
    Exit Sub
End If

If UCase$(Left$(rdata, 12)) = "/GUARDARMAPA" Then
Call SaveMapData(UserList(userindex).POS.Map)
Call SendData(ToIndex, userindex, 0, "||El mapa " & UserList(userindex).POS.Map & " se guardó correctamente." & FONTTYPE_INFO)
Exit Sub
End If

If UCase$(Left$(rdata, 8)) = "/UNBANIP" Then
    
    
    rdata = Right$(rdata, Len(rdata) - 9)
    Call LogGM(UserList(userindex).Name, "/UNBANIP " & rdata, False)
    
    For LoopC = 1 To BanIps.Count
        If BanIps.Item(LoopC) = rdata Then
            BanIps.Remove LoopC
            Call SendData(ToIndex, userindex, 0, "||La IP " & BanIP & " se ha quitado de la lista de bans." & FONTTYPE_INFO)
            Exit Sub
        End If
    Next
    
    Call SendData(ToIndex, userindex, 0, "||La IP " & rdata & " NO se encuentra en la lista de bans." & FONTTYPE_INFO)
    
    Exit Sub
End If

If UCase$(Left$(rdata, 3)) = "/CT" Then
    
    rdata = Right$(rdata, Len(rdata) - 4)
    Call LogGM(UserList(userindex).Name, "/CT: " & rdata, False)
    mapa = ReadField(1, rdata, 32)
    X = ReadField(2, rdata, 32)
    Y = ReadField(3, rdata, 32)
    
    If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y - 1).OBJInfo.OBJIndex Then
        Exit Sub
    End If
    If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y - 1).TileExit.Map Then
        Exit Sub
    End If
    If Not MapaValido(mapa) Or Not InMapBounds(X, Y) Then Exit Sub
    
    Dim et As Obj
    et.Amount = 1
    et.OBJIndex = Teleport
    
    Call MakeObj(ToMap, 0, UserList(userindex).POS.Map, et, UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y - 1)
    MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y - 1).TileExit.Map = mapa
    MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y - 1).TileExit.X = X
    MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y - 1).TileExit.Y = Y
    
    Exit Sub
End If



If UCase$(Left$(rdata, 3)) = "/DT" Then
    
    Call LogGM(UserList(userindex).Name, "/DT", False)
    
    mapa = UserList(userindex).flags.TargetMap
    X = UserList(userindex).flags.TargetX
    Y = UserList(userindex).flags.TargetY
    
    If ObjData(MapData(mapa, X, Y).OBJInfo.OBJIndex).ObjType = OBJTYPE_TELEPORT And _
        MapData(mapa, X, Y).TileExit.Map Then
        Call EraseObj(ToMap, 0, mapa, MapData(mapa, X, Y).OBJInfo.Amount, mapa, X, Y)
        MapData(mapa, X, Y).TileExit.Map = 0
        MapData(mapa, X, Y).TileExit.X = 0
        MapData(mapa, X, Y).TileExit.Y = 0
    End If
    
    Exit Sub
End If




ErrorHandler:
 Call LogError("HandleData. CadOri:" & CadenaOriginal & " Nom:" & UserList(userindex).Name & " UI:" & userindex & " N: " & Err.Number & " D: " & Err.Description)
Call SendData(ToIndex, userindex, 0, "Comando invalido." & FONTTYPE_INFO)
Call Siguenloscomandos(userindex, rdata)

End Sub
Private Sub GuardarInt(ByVal Ruta As String, ByVal Data As Integer)
    Dim f As Integer
    f = FreeFile
    Open Ruta For Output As f
    Print #f, Data
    Close #f
End Sub
Sub Siguenloscomandos(userindex As Integer, ByVal rdata As String)
Dim TempTick As Long
Dim sndData As String
Dim CadenaOriginal As String

Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim numeromail As Integer
Dim TIndex As Integer
Dim tName As String
Dim Clase As Byte
Dim NumNPC As Integer
Dim tMessage As String
Dim i As Integer
Dim auxind As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim arg3 As String
Dim Arg4 As String
Dim Arg5 As Integer
Dim Arg6 As String
Dim DummyInt As Integer
Dim Antes As Boolean
Dim Ver As String
Dim encpass As String
Dim pass As String
Dim mapa As Integer
Dim usercon As String
Dim nameuser As String
Dim Name As String
Dim ind
Dim GMDia As String
Dim GMMapa As String
Dim GMPJ As String
Dim GMMail As String
Dim GMGM As String
Dim GMTitulo As String
Dim GMMensaje As String
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim cliMD5 As String
Dim UserFile As String
Dim UserName As String
UserName = UserList(userindex).Name
UserFile = CharPath & UCase$(UserName) & ".chr"
Dim ClientCRC As String
Dim ServerSideCRC As Long
Dim NombreIniChat As String
Dim cantidadenmapa As Integer
Dim Prueba1 As Integer
CadenaOriginal = rdata

If UCase$(Left$(rdata, 10)) = "/DOBACKUPL" Then
    Call DoBackUp(True)
    Call SaveGuildsNew
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/GRABAR" Then
    Call GuardarUsuarios
    Exit Sub
End If

If UCase$(Left$(rdata, 12)) = "/MODMAPINFO " Then
    If UserList(userindex).flags.Privilegios < 3 Then Exit Sub
    Call LogGM(UserList(userindex).Name, rdata, False)
    rdata = Right(rdata, Len(rdata) - 12)
    Select Case UCase(ReadField(1, rdata, 32))
    Case "PK"
        tStr = ReadField(2, rdata, 32)
        If tStr <> "" Then
            MapInfo(UserList(userindex).POS.Map).Pk = IIf(tStr = "0", True, False)
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(userindex).POS.Map & ".dat", "Mapa" & UserList(userindex).POS.Map, "Pk", tStr)
        End If
        Call SendData(ToIndex, userindex, 0, "||Mapa " & UserList(userindex).POS.Map & " PK: " & MapInfo(UserList(userindex).POS.Map).Pk & FONTTYPE_FENIX)
    Case "BACKUP"
        tStr = ReadField(2, rdata, 32)
        If tStr <> "" Then
            MapInfo(UserList(userindex).POS.Map).BackUp = CByte(tStr)
            Call WriteVar(App.Path & MapPath & "mapa" & UserList(userindex).POS.Map & ".dat", "Mapa" & UserList(userindex).POS.Map, "backup", tStr)
        End If
       
        Call SendData(ToIndex, userindex, 0, "||Mapa " & UserList(userindex).POS.Map & " Backup: " & MapInfo(UserList(userindex).POS.Map).BackUp & FONTTYPE_FENIX)
    End Select
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/PAUSA" Then
    
    If haciendoBK Then Exit Sub
    
    Enpausa = Not Enpausa
    
    If Enpausa Then
        Call SendData(ToAll, 0, 0, "TL" & 197)
        Call SendData(ToAll, 0, 0, "||Servidor> El mundo ha sido detenido." & FONTTYPE_INFO)
        Call SendData(ToAll, 0, 0, "BKW")
        Call SendData(ToAll, 0, 0, "TM" & "0")
    Else
        Call SendData(ToAll, 0, 0, "TL")
        Call SendData(ToAll, 0, 0, "||Servidor> Juego reanudado." & FONTTYPE_INFO)
        Call SendData(ToAll, 0, 0, "BKW")
        Call SendData(ToIndex, userindex, 0, "TM" & MapInfo(UserList(userindex).POS.Map).Music)
    End If
Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/SHOW INT" Then
    Call frmMain.mnuMostrar_Click
    Exit Sub
End If

If UCase$(rdata) = "/LLUVIA" Then
    Lloviendo = Not Lloviendo
    Call SendData(ToAll, 0, 0, "LLU")
    Exit Sub
End If

If UCase$(rdata) = "/LIMPIARMUNDO" Then
If UserList(userindex).flags.Privilegios = 3 Then
Call SendData(ToAll, 0, 0, "||Se realizará una limpieza del Mundo en 1 minuto. Por favor recojan sus pertenencias." & FONTTYPE_FENIX)
frmMain.Tlimpiar.Enabled = True
Call LogGM(UserList(userindex).Name, "Ejecutó una limpieza del Mundo.", True)
End If
Exit Sub
End If

If UCase$(rdata) = "/PASSDAY" Then
    Call DayElapsed
    Exit Sub
End If


If UCase$(rdata) = "/INTERVALOS" Then
    Call SendData(ToIndex, userindex, 0, "||Golpe-Golpe: " & IntervaloUserPuedeAtacar & " segundos." & FONTTYPE_INFO)
    Call SendData(ToIndex, userindex, 0, "||Golpe-Hechizo: " & IntervaloUserPuedeGolpeHechi & " segundos." & FONTTYPE_INFO)
    Call SendData(ToIndex, userindex, 0, "||Hechizo-Hechizo: " & IntervaloUserPuedeCastear & " segundos." & FONTTYPE_INFO)
    Call SendData(ToIndex, userindex, 0, "||Hechizo-Golpe: " & IntervaloUserPuedeHechiGolpe & " segundos." & FONTTYPE_INFO)
    Call SendData(ToIndex, userindex, 0, "||Arco-Arco: " & IntervaloUserFlechas & " segundos." & FONTTYPE_INFO)
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/MODS " Then
    Dim PreInt As Single
    rdata = Right$(rdata, Len(rdata) - 6)
    TIndex = ClaseIndex(ReadField(1, rdata, 64))
    If TIndex = 0 Then Exit Sub
    tInt = ReadField(2, rdata, 64)
    If tInt < 1 Or tInt > 6 Then Exit Sub
    Arg5 = ReadField(3, rdata, 64)
    If Arg5 < 40 Or Arg5 > 125 Then Exit Sub
    PreInt = Mods(tInt, TIndex)
    Mods(tInt, TIndex) = Arg5 / 100
    Call SendData(ToAdmins, 0, 0, "||El modificador n° " & tInt & " de la clase " & ListaClases(TIndex) & " fue cambiado de " & PreInt & " a " & Mods(tInt, TIndex) & "." & FONTTYPE_FIGHT)
    Call SaveMod(tInt, TIndex)
    Exit Sub
End If
If UCase$(Left$(rdata, 5)) = "/BLOQ" Then
    Call LogGM(UserList(userindex).Name, "/BLOQ", False)
    rdata = Right$(rdata, Len(rdata) - 5)
    If MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y).Blocked = 0 Then
        MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y).Blocked = 1
        Call Bloquear(ToMap, userindex, UserList(userindex).POS.Map, UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y, 1)
    Else
        MapData(UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y).Blocked = 0
        Call Bloquear(ToMap, userindex, UserList(userindex).POS.Map, UserList(userindex).POS.Map, UserList(userindex).POS.X, UserList(userindex).POS.Y, 0)
    End If
    Exit Sub
End If


If UCase$(rdata) = "/MASSKILL" Then
    For Y = UserList(userindex).POS.Y - MinYBorder + 1 To UserList(userindex).POS.Y + MinYBorder - 1
            For X = UserList(userindex).POS.X - MinXBorder + 1 To UserList(userindex).POS.X + MinXBorder - 1
                If X > 0 And Y > 0 And X < 101 And Y < 101 Then _
                    If MapData(UserList(userindex).POS.Map, X, Y).NpcIndex Then Call QuitarNPC(MapData(UserList(userindex).POS.Map, X, Y).NpcIndex)
            Next
    Next
    Call LogGM(UserList(userindex).Name, "/MASSKILL", False)
    Exit Sub
End If

If UCase$(Left$(rdata, 6)) = "/SMSG " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Call LogGM(UserList(userindex).Name, "Mensaje de sistema:" & rdata, False)
    Call SendData(ToAll, 0, 0, "!!" & rdata & ENDC)
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/ACT1V1" Then
    If Actretos = True Then
        Actretos = False
        Call SendData(ToAll, 0, 0, "||Retos 1 vs 1 desactivados" & FONTTYPE_INFO)
    Else
        Call SendData(ToAll, 0, 0, "||Retos 1 vs 1 activados" & FONTTYPE_INFO)
        Actretos = True
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/ACT2V2" Then
    If OPCDuelos.ACT = True Then
        OPCDuelos.ACT = False
        frmMain.retos2vs2.Enabled = False '> CUANDO CREEN EL TIMER, CAMBIENLEN EL NOMBRE.
        Call SendData(ToAll, 0, 0, "||Retos 2 vs 2 desactivados" & FONTTYPE_INFO)
    Else
        Call SendData(ToAll, 0, 0, "||Retos 2 vs 2 activados" & FONTTYPE_INFO)
        OPCDuelos.ACT = True
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/ACC " Then
   rdata = val(Right$(rdata, Len(rdata) - 5))
   NumNPC = val(GetVar(App.Path & "\Dat\NPCs-HOSTILES.dat", "INIT", "NumNPCs")) + 500
   If rdata < 0 Or rdata > NumNPC Then
       Call SendData(ToIndex, userindex, 0, "||La criatura no existe." & FONTTYPE_INFO)

Else
   Call SpawnNpc(val(rdata), UserList(userindex).POS, True, False)


   End If
   Exit Sub
End If


If UCase$(Left$(rdata, 6)) = "/RACC " Then
   rdata = val(Right$(rdata, Len(rdata) - 6))
      NumNPC = val(GetVar(App.Path & "\Dat\NPCs-HOSTILES.dat", "INIT", "NumNPCs")) + 500
   If rdata < 0 Or rdata > NumNPC Then
    Call SendData(ToIndex, userindex, 0, "||La criatura no existe." & FONTTYPE_INFO)
Else
   Call SpawnNpc(val(rdata), UserList(userindex).POS, True, True)
   End If
   Exit Sub
End If

If UCase$(rdata) = "/NAVE" Then
    If UserList(userindex).flags.Navegando Then
        UserList(userindex).flags.Navegando = 0
    Else
        UserList(userindex).flags.Navegando = 1
    End If
    Exit Sub
End If

If UCase$(rdata) = "/APAGAR" Then
    Call LogMain(" Server apagado por " & UserList(userindex).Name & ".")
    Call ApagarSistema
    End
End If

If UCase$(rdata) = "/REINICIAR" Then
    Call LogMain(" Server reiniciado por " & UserList(userindex).Name & ".")
    ShellExecute frmMain.hwnd, "open", App.Path & "/AOKreiZy.exe", "", "", 1
    Call ApagarSistema
    Exit Sub
End If

If UCase$(Left$(rdata, 4)) = "/INT" Then
    rdata = Right$(rdata, Len(rdata) - 4)
    
    Select Case UCase$(Left$(rdata, 2))

        Case "GG"
            rdata = Right$(rdata, Len(rdata) - 3)
            PreInt = IntervaloUserPuedeAtacar
            IntervaloUserPuedeAtacar = val(rdata) / 10
            Call SendData(ToAdmins, 0, 0, "||El intervalo Golpe-Golpe fue cambiado de " & PreInt & " a " & IntervaloUserPuedeAtacar & " segundos." & FONTTYPE_INFO)
            Call SendData(ToAll, 0, 0, "INTA" & IntervaloUserPuedeAtacar * 10)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar", IntervaloUserPuedeAtacar * 10)
        Case "GH"
            rdata = Right$(rdata, Len(rdata) - 3)
            PreInt = IntervaloUserPuedeGolpeHechi
            IntervaloUserPuedeGolpeHechi = val(rdata) / 10
            Call SendData(ToAdmins, 0, 0, "||El intervalo Golpe-Hechizo fue cambiado de " & PreInt & " a " & IntervaloUserPuedeGolpeHechi & " segundos." & FONTTYPE_INFO)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeGolpeHechi", IntervaloUserPuedeGolpeHechi * 10)
        Case "HH"
            rdata = Right$(rdata, Len(rdata) - 3)
            PreInt = IntervaloUserPuedeCastear
            IntervaloUserPuedeCastear = val(rdata) / 10
            Call SendData(ToAdmins, 0, 0, "||El intervalo Hechizo-Hechizo fue cambiado de " & PreInt & " a " & IntervaloUserPuedeCastear & " segundos." & FONTTYPE_INFO)
            Call SendData(ToAll, 0, 0, "INTS" & IntervaloUserPuedeCastear * 10)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo", IntervaloUserPuedeCastear * 10)
        Case "HG"
            rdata = Right$(rdata, Len(rdata) - 3)
            PreInt = IntervaloUserPuedeHechiGolpe
            IntervaloUserPuedeHechiGolpe = val(rdata) / 10
            Call SendData(ToAdmins, 0, 0, "||El intervalo Hechizo-Golpe fue cambiado de " & PreInt & " a " & IntervaloUserPuedeHechiGolpe & " segundos." & FONTTYPE_INFO)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeHechiGolpe", IntervaloUserPuedeHechiGolpe * 10)
        Case "AA"
            rdata = Right$(rdata, Len(rdata) - 2)
            PreInt = IntervaloUserFlechas
            IntervaloUserFlechas = val(rdata) / 10
            Call SendData(ToAdmins, 0, 0, "||El intervalo de flechas fue cambiado de " & PreInt & " a " & IntervaloUserFlechas & " segundos." & FONTTYPE_INFO)
            Call SendData(ToIndex, userindex, 0, "INTF" & IntervaloUserFlechas * 10)
            
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserFlechas", IntervaloUserFlechas * 10)
        Case "SH"
            rdata = Right$(rdata, Len(rdata) - 2)
            PreInt = IntervaloUserSH
            IntervaloUserSH = val(rdata)
            Call SendData(ToAdmins, 0, 0, "||Intervalo de SH cambiado de " & PreInt & " a " & IntervaloUserSH & " segundos de tardanza." & FONTTYPE_INFO)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserSH", str(IntervaloUserSH))
        Case "PN"
            rdata = Right$(rdata, Len(rdata) - 2)
            PreInt = IntervaloUserPuedePocion
            IntervaloUserPuedePocion = val(rdata)
            Call SendData(ToAdmins, 0, 0, "||Intervalo de SH cambiado de " & PreInt & " a " & IntervaloUserPuedePocion & " segundos de tardanza." & FONTTYPE_INFO)
            Call WriteVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedePocion", str(IntervaloUserPuedePocion))
    End Select
End If
If UCase$(rdata) = "/DATS" Then
    Call CargarHechizos
    Call LoadOBJData
    Call DescargaNpcsDat
    Call CargaNpcsDat
    Exit Sub
End If
If UCase$(Left$(rdata, 6)) = "/ITEM " Then
    rdata = Right$(rdata, Len(rdata) - 6)
    Dim et As Obj
    et.OBJIndex = val(ReadField(1, rdata, Asc(" ")))
    et.Amount = val(ReadField(2, rdata, Asc(" ")))
    If et.Amount <= 0 Then et.Amount = 1
    If et.OBJIndex < 1 Or et.OBJIndex > NumObjDatas Then Exit Sub
    If et.Amount > MAX_INVENTORY_OBJS Then Exit Sub
    If Not MeterItemEnInventario(userindex, et) Then Call TirarItemAlPiso(UserList(userindex).POS, et)
    Call LogGM(UserList(userindex).Name, "Creo objeto:" & ObjData(et.OBJIndex).Name & " (" & et.Amount & ")", False)
    Exit Sub
End If


If UCase$(Left$(rdata, 7)) = "/SEGURO" Then
        If MapInfo(UserList(userindex).POS.Map).Pk = True Then
            MapInfo(UserList(userindex).POS.Map).Pk = False
            Call SendData(ToIndex, userindex, 0, "||Ahora es zona segura." & FONTTYPE_INFO)
            Exit Sub
        Else
            MapInfo(UserList(userindex).POS.Map).Pk = True
            Call SendData(ToIndex, userindex, 0, "||Ahora es zona insegura." & FONTTYPE_INFO)
            Exit Sub
        End If
        Exit Sub
    End If
If UCase$(rdata) = "/MODOQUEST" Then
    ModoQuest = Not ModoQuest
    If ModoQuest Then
        Call SendData(ToAll, 0, 0, "||Modo Quest activado." & FONTTYPE_FENIX)
        Call SendData(ToAll, 0, 0, "||Los neutrales pueden poner /MERCENARIO ALIANZA o /MERCENARIO CRIMINAL para enlistarse en alguna facción temporalmente durante la quest." & FONTTYPE_FENIX)
        Call SendData(ToAll, 0, 0, "||Al morir puedes poner /HOGAR y serás teletransportado a Ullathorpe." & FONTTYPE_FENIX)
    Else
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & " desactivó el modo quest." & FONTTYPE_FENIX)
        Call DesactivarMercenarios
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 7)) = "/STAFF " Then
    rdata = Right$(rdata, Len(rdata) - 7)
    If Len(rdata) > 0 Then
        Call SendData(ToConci, 0, 0, "||" & UserList(userindex).Name & "> " & rdata & "~255~255~255~0~1")
        Call SendData(ToConse, 0, 0, "||" & UserList(userindex).Name & "> " & rdata & "~255~255~255~0~1")
        Call SendData(ToAdmins, 0, 0, "||" & UserList(userindex).Name & "> " & rdata & "~255~255~255~0~1")
    End If
    Exit Sub
End If

If UCase$(Left$(rdata, 5)) = "/MOD " Then
    Call LogGM(UserList(userindex).Name, rdata, False)
    rdata = Right$(rdata, Len(rdata) - 5)
    TIndex = NameIndex(ReadField(1, rdata, 32))
    Arg1 = ReadField(2, rdata, 32)
    Arg2 = ReadField(3, rdata, 32)
    arg3 = ReadField(4, rdata, 32)
    Arg4 = ReadField(5, rdata, 32)
    If TIndex <= 0 Then
        Call SendData(ToIndex, userindex, 0, "1A")
        Exit Sub
    End If
    If UserList(TIndex).flags.Privilegios > 2 And userindex <> TIndex Then Exit Sub
    
    Select Case UCase$(Arg1)
        Case "RAZA"
            If val(Arg2) < 6 Then
                UserList(TIndex).Raza = val(Arg2)
                Call DarCuerpoDesnudo(TIndex)
                Call ChangeUserChar(ToMap, 0, UserList(userindex).POS.Map, userindex, UserList(userindex).Char.Body, UserList(userindex).Char.Head, UserList(userindex).Char.Heading, UserList(userindex).Char.WeaponAnim, UserList(userindex).Char.ShieldAnim, UserList(userindex).Char.CascoAnim)
            End If
        Case "JER"
            UserList(userindex).Faccion.Jerarquia = 0
        Case "BANDO"
            If val(Arg2) < 3 Then
                If val(Arg2) > 0 Then Call SendData(ToIndex, TIndex, 0, Mensajes(val(Arg2), 10))
                UserList(TIndex).Faccion.Bando = val(Arg2)
                UserList(TIndex).Faccion.BandoOriginal = val(Arg2)
                If Not PuedeFaccion(TIndex) Then Call SendData(ToIndex, TIndex, 0, "SUFA0")
                Call UpdateUserChar(TIndex)
                If val(Arg2) = 0 Then UserList(TIndex).Faccion.Jerarquia = 0
            End If
        Case "SKI"
            If val(Arg2) >= 0 And val(Arg2) <= 100 Then
                For i = 1 To NUMSKILLS
                    UserList(TIndex).Stats.UserSkills(i) = val(Arg2)
                Next
            End If
        Case "CLASE"
            i = ClaseIndex(Arg2)
            If i = 0 Then Exit Sub
            UserList(TIndex).Clase = i
            UserList(TIndex).Recompensas(1) = 0
            UserList(TIndex).Recompensas(2) = 0
            UserList(TIndex).Recompensas(3) = 0
            Call SendData(ToIndex, TIndex, 0, "||Ahora eres " & ListaClases(i) & "." & FONTTYPE_INFO)
            If PuedeRecompensa(userindex) Then
                Call SendData(ToIndex, userindex, 0, "SURE1")
            Else: Call SendData(ToIndex, userindex, 0, "SURE0")
            End If
            If PuedeSubirClase(userindex) Then
                Call SendData(ToIndex, userindex, 0, "SUCL1")
            Else: Call SendData(ToIndex, userindex, 0, "SUCL0")
            End If
        
        Case "ORO"
            If val(Arg2) > 100000000 Then Arg2 = 10000000
            UserList(TIndex).Stats.GLD = val(Arg2)
            Call SendUserORO(TIndex)
        Case "EXP"
            If val(Arg2) > 100000000 Then Arg2 = 10000000
            UserList(TIndex).Stats.Exp = val(Arg2)
            Call CheckUserLevel(TIndex)
            Call SendUserEXP(TIndex)
        Case "MEX"
            If val(Arg2) > 10000000 Then Arg2 = 10000000
            UserList(TIndex).Stats.Exp = UserList(TIndex).Stats.Exp + val(Arg2)
            Call CheckUserLevel(TIndex)
            Call SendUserEXP(TIndex)
             Case "ADV"
            UserList(TIndex).Stats.advertencias = UserList(TIndex).Stats.advertencias + val(Arg2)
        Case "BODY"
            Call ChangeUserBody(ToMap, 0, UserList(TIndex).POS.Map, TIndex, val(Arg2))
        Case "HEAD"
            Call ChangeUserHead(ToMap, 0, UserList(TIndex).POS.Map, TIndex, val(Arg2))
            UserList(TIndex).OrigChar.Head = val(Arg2)
        Case "PHEAD"
            UserList(TIndex).OrigChar.Head = val(Arg2)
            Call ChangeUserHead(ToMap, 0, UserList(TIndex).POS.Map, TIndex, val(Arg2))
        Case "TOR"
            UserList(TIndex).Faccion.torneos = val(Arg2)
        Case "QUE"
            UserList(TIndex).Faccion.Quests = val(Arg2)
        Case "NEU"
            UserList(TIndex).Faccion.Matados(Neutral) = val(Arg2)
        Case "CRI"
            UserList(TIndex).Faccion.Matados(Caos) = val(Arg2)
        Case "CIU"
            UserList(TIndex).Faccion.Matados(Real) = val(Arg2)
        Case "HP"
            If val(Arg2) > 30000 Then Exit Sub
            UserList(TIndex).Stats.MaxHP = val(Arg2)
            Call SendUserMAXHP(userindex)
        Case "MAN"
            If val(Arg2) > 2200 + 27800 * Buleano(UserList(TIndex).Clase = MAGO And UserList(TIndex).Recompensas(2) = 2) Then Exit Sub
            UserList(TIndex).Stats.MaxMAN = val(Arg2)
            Call SendUserMAXMANA(userindex)
        Case "STA"
            If val(Arg2) > 30000 Then Exit Sub
            UserList(TIndex).Stats.MaxSta = val(Arg2)
        Case "HAM"
            UserList(TIndex).Stats.MinHam = val(Arg2)
        Case "SED"
            UserList(TIndex).Stats.MinAGU = val(Arg2)
        Case "ATF"
            If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
            UserList(TIndex).Stats.UserAtributos(fuerza) = val(Arg2)
            UserList(TIndex).Stats.UserAtributosBackUP(fuerza) = val(Arg2)
            Call UpdateFuerzaYAg(TIndex)
        Case "ATI"
            If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
            UserList(TIndex).Stats.UserAtributos(Inteligencia) = val(Arg2)
            UserList(TIndex).Stats.UserAtributosBackUP(Inteligencia) = val(Arg2)
        Case "ATA"
            If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
            UserList(TIndex).Stats.UserAtributos(Agilidad) = val(Arg2)
            UserList(TIndex).Stats.UserAtributosBackUP(Agilidad) = val(Arg2)
            Call UpdateFuerzaYAg(TIndex)
        Case "CANJE"
            UserList(TIndex).flags.Canje = val(Arg2)
        Case "ATC"
            If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
            UserList(TIndex).Stats.UserAtributos(Carisma) = val(Arg2)
            UserList(TIndex).Stats.UserAtributosBackUP(Carisma) = val(Arg2)
        Case "ATV"
            If val(Arg2) > 21 Or val(Arg2) < 6 Then Exit Sub
            UserList(TIndex).Stats.UserAtributos(Constitucion) = val(Arg2)
            UserList(TIndex).Stats.UserAtributosBackUP(Constitucion) = val(Arg2)
        Case "LEVEL"
            If val(Arg2) < 1 Or val(Arg2) > STAT_MAXELV Then Exit Sub
            UserList(TIndex).Stats.ELV = val(Arg2)
            UserList(TIndex).Stats.ELU = ELUs(UserList(TIndex).Stats.ELV)
            Call SendData(ToIndex, TIndex, 0, "5O" & UserList(TIndex).Stats.ELV & "," & UserList(TIndex).Stats.ELU)
            If PuedeRecompensa(userindex) Then
                Call SendData(ToIndex, userindex, 0, "SURE1")
            Else: Call SendData(ToIndex, userindex, 0, "SURE0")
            End If
            If PuedeSubirClase(userindex) Then
                Call SendData(ToIndex, userindex, 0, "SUCL1")
            Else: Call SendData(ToIndex, userindex, 0, "SUCL0")
            End If
        Case Else
            Call SendData(ToIndex, userindex, 0, "||Comando inexistente." & FONTTYPE_INFO)
    End Select

    Exit Sub
End If

If UCase$(Left$(rdata, 9)) = "/DOBACKUP" Then
    Call DoBackUp
    Exit Sub
End If



End Sub
