Attribute VB_Name = "mdlRetos"
Public Sub ComensarDuelo(ByVal UserIndex As Integer, ByVal TIndex As Integer)
    UserList(UserIndex).flags.EstaDueleando = True
    UserList(UserIndex).flags.Oponente = TIndex
    UserList(TIndex).flags.EstaDueleando = True
    Call WarpUserChar(TIndex, 199, 37, 44)
    UserList(TIndex).flags.Oponente = UserIndex
    Call WarpUserChar(UserIndex, 199, 58, 56)
    Call SendData(ToAll, 0, 0, "||RETOS> " & UserList(TIndex).Name & " y " & UserList(UserIndex).Name & " van a competir en un reto por 25 puntos de honor." & FONTTYPE_RETOS)
End Sub
Public Sub ResetDuelo(ByVal UserIndex As Integer, ByVal TIndex As Integer)
    UserList(UserIndex).flags.EsperandoDuelo = False
    UserList(UserIndex).flags.Oponente = 0
    UserList(UserIndex).flags.EstaDueleando = False
    UserList(UserIndex).flags.Honor = UserList(UserIndex).flags.Honor + 20
    Call WarpUserChar(UserIndex, 200, 50, 50)
    Call WarpUserChar(TIndex, 200, 51, 51)
    UserList(TIndex).flags.Honor = UserList(TIndex).flags.Honor - 20
    UserList(TIndex).flags.EsperandoDuelo = False
    UserList(TIndex).flags.Oponente = 0
    UserList(TIndex).flags.EstaDueleando = False
End Sub
Public Sub TerminarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(ToAll, Ganador, 0, "||RETOS> " & UserList(Ganador).Name & " venció a " & UserList(Perdedor).Name & " en un reto." & FONTTYPE_TALK)
    Call ResetDuelo(Ganador, Perdedor)
End Sub
Public Sub DesconectarDuelo(ByVal Ganador As Integer, ByVal Perdedor As Integer)
    Call SendData(ToAll, Ganador, 0, "||RETOS> El reto ha sido cancelado por la desconexión de " & UserList(Perdedor).Name & "." & FONTTYPE_TALK)
    Call ResetDuelo(Ganador, Perdedor)
End Sub
