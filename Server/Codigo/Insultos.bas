Attribute VB_Name = "Insultos"
Public Function EsMalaPalabra(ByVal rdata As String)
If ReconocerPalabra("TROLO", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("PUTO", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("GAY", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("PUTOS", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("WWW.", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("AO", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("GM", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("GAME MASTER", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("BOLUDO", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("PELOTUDO", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("PT", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("PTS", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("MANCO", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("MANCOS", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("PUTA", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("IDIOTA", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("MIERDA", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("TURBINAS", UCase$(rdata)) Then EsMalaPalabra = True
If ReconocerPalabra("PETE", UCase$(rdata)) Then EsMalaPalabra = True
 
End Function
 
Public Function HayAdminsOnline() As Boolean
    Dim i As Integer
        For i = 1 To LastUser
            If UserList(i).flags.Privilegios > 0 Then HayAdminsOnline = True
        Next i
End Function
Private Function ReconocerPalabra(ByVal Palabra As String, ByVal Donde As String) As Boolean
Dim i As Integer
For i = 1 To (Len(Donde) - Len(Palabra) + 1)
 If UCase(Mid(Donde, i, 1)) = UCase(Mid(Palabra, 1, 1)) Then
       If UCase(Mid(Donde, i, Len(Palabra))) = UCase(Mid(Palabra, 1, Len(Palabra))) Then
         ReconocerPalabra = True
         Exit Function 'Gracias Rheniek
       Else
         ReconocerPalabra = False
        End If
  Else
        ReconocerPalabra = False
  End If
Next
End Function
