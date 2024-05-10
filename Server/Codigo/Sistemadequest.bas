Attribute VB_Name = "Sistemadequest"

Option Explicit
Public Const NPCRey As Integer = 639 ' aca tienen q pner el npc del bichos.dat o como se llame
'Public Const DaPuntosHonor As Integer = 500
Public Const CastilloMap As Byte = 4 ' aca el mapa donde respawnea
Public Const CastilloX As Byte = 50 ' aca x donde respawnea
Public Const CastilloY As Byte = 50 'aca Y donde respawnea
Public GolpesRey As Byte
Public HayRey As Byte
 
 
Public Sub MuereReyAZUL(ByVal NpcIndex As Integer, UserIndex As Integer)
    Dim assd As Boolean
    

Dim LoopC As Integer
Dim addd As Integer
Dim grupoazuul As Boolean
Dim gruporoojo As Boolean
For LoopC = 1 To LastUser

DiaC = Date
HoraC = Time

If yaganoo = False Then

Call SendData(ToAll, 0, 0, "||La quest la ha ganado el equipo azul." & FONTTYPE_FENIX & ENDC)
Call SendData(ToAll, 0, 0, "||PREMIO. 150 puntos de honor y 1 punto de canjeo" & FONTTYPE_FENIX & ENDC)
yaganoo = True
Call WriteVar(IniPath & "quest.ini", "INIT", "Castillo", "AZUL")
Call WriteVar(IniPath & "quest.ini", "INIT", "DiaC", Date)
Call WriteVar(IniPath & "quest.ini", "INIT", "HoraC", Time)
Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
End If

If UserList(LoopC).flags.GrupoAzul = True Then

UserList(LoopC).flags.Honor = UserList(LoopC).flags.Honor + 150
UserList(LoopC).flags.Canje = UserList(LoopC).flags.Canje + 1
hay_Quest = False
Quest_cantidad = 0
Questt = 0
XX1 = 0
XX2 = 0
Call WarpUserChar(LoopC, 1, 50, 50, True)
End If




Next LoopC
 
 Call QuitarNPC(NpcIndex)


If UserIndex > 0 Then Call SubirSkill(UserIndex, Supervivencia, 1) 'cambiar el 1 por la dificultad de que suba supervivencia

Exit Sub

End Sub
 
Public Sub MuereReyROJO(ByVal NpcIndex As Integer, UserIndex As Integer)
    Dim assd As Boolean
    

Dim LoopC As Integer
Dim addd As Integer
Dim grupoazuul As Boolean
Dim gruporoojo As Boolean
For LoopC = 1 To LastUser

DiaC = Date
HoraC = Time




 If yaganoo = False Then
 
Call SendData(ToAll, 0, 0, "||La quest la ha ganado el equipo rojo." & FONTTYPE_FENIX & ENDC)
Call SendData(ToAll, 0, 0, "||PREMIO. 150 puntos de honor y 1 punto de canjeo" & FONTTYPE_FENIX & ENDC)
Call WriteVar(IniPath & "quest.ini", "INIT", "Castillo", "ROJO")
Call WriteVar(IniPath & "quest.ini", "INIT", "DiaC", Date)
Call WriteVar(IniPath & "quest.ini", "INIT", "HoraC", Time)
Call SendData(ToAll, 0, 0, "TW" & SND_CREACIONCLAN)
yaganoo = True

End If


 If UserList(LoopC).flags.GrupoRojo = True Then

 Quest_cantidad = 0
Questt = 0


UserList(LoopC).flags.Honor = UserList(LoopC).flags.Honor + 150
UserList(LoopC).flags.Canje = UserList(LoopC).flags.Canje + 1
Call WarpUserChar(LoopC, 1, 50, 50, True)
hay_Quest = False
XX1 = 0
XX2 = 0

End If


Next LoopC
 
 Call QuitarNPC(NpcIndex)


If UserIndex > 0 Then Call SubirSkill(UserIndex, Supervivencia, 1) 'cambiar el 1 por la dificultad de que suba supervivencia

Exit Sub

End Sub

