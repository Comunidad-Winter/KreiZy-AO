Attribute VB_Name = "Mod_TileEngine"
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







Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521


Public Const SRCCOPY = &HCC0020









Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type


Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type


Public Type Position
    x As Integer
    Y As Integer
End Type

Public Type Position2
    x As Double
    Y As Double
End Type

Public Type WorldPos
    Map As Integer
    x As Integer
    Y As Integer
End Type



Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames(1 To 25) As Integer
    Speed As Integer
End Type


Public Type Grh
    GrhIndex As Integer
    FrameCounter As Double
    SpeedCounter As Byte
    Started As Byte
End Type


Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type


Public Type HeadData
    Head(1 To 4) As Grh
End Type


Type WeaponAnimData
    WeaponWalk(1 To 4) As Grh
End Type


Type ShieldAnimData
    ShieldWalk(1 To 4) As Grh
End Type



Public Type FxData
    FX As Grh
    OffsetX As Long
    OffsetY As Long
End Type


Public Type Char
    Active As Byte
    Heading As Byte
    POS As Position

    Body As BodyData
    Head As HeadData
    casco As HeadData
    arma As WeaponAnimData
    escudo As ShieldAnimData
    UsandoArma As Boolean
    FX As Integer
    FxLoopTimes As Integer
    Criminal As Byte
    Navegando As Byte
    
    Nombre As String
    GM As Integer
    
    haciendoataque As Byte
    Moving As Byte
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    MoveOffset As Position2
    ServerIndex As Integer
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    
End Type


Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type


Public Type MapBlock
    Graphic(1 To 4) As Grh
    charindex As Integer
    ObjGrh As Grh

    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type


Public Type MapInfo
    Music As String
    name As String
    StartPos As WorldPos
    MapVersion As Integer
    
    
    Changed As Byte
End Type


Public IniPath As String
Public MapPath As String



Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte


Public CurMap As Integer
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position
Public AddtoUserPos As Position
Public UserCharIndex As Integer

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long


Public WindowTileWidth As Integer
Public WindowTileHeight As Integer


Public MainViewTop As Integer
Public MainViewLeft As Integer




Public TileBufferSize As Integer


Public DisplayFormhWnd As Long


Public TilePixelHeight As Integer
Public TilePixelWidth As Integer



Public NumBodies As Integer
Public Numheads As Integer
Public NumFxs As Integer

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer
Public NumShieldAnims As Integer



Public lastTime As Long



Public MainDestRect   As RECT

Public MainViewRect   As RECT
Public BackBufferRect As RECT

Public MainViewWidth As Integer
Public MainViewHeight As Integer





Public GrhData() As GrhData
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Grh() As Grh



Public MapData() As MapBlock
Public MapInfo As MapInfo



Public CharList(1 To 10000) As Char




Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uRetrunLength As Long, ByVal hwndCallback As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long






Public bRain        As Boolean
Public bRainST      As Boolean
Public bTecho       As Boolean
Public brstTick     As Long

Private RLluvia(7)  As RECT
Private iFrameIndex As Byte
Private llTick      As Long
Private LTLluvia(4) As Integer
            

Public Enum TextureStatus
    tsOriginal = 0
    tsNight = 1
    tsFog = 2
End Enum


    Public Enum PlayLoop
        plNone = 0
        plLluviain = 1
        plLluviaout = 2
        plFogata = 3
    End Enum





Sub CargarCabezas()
On Error Resume Next
Dim n As Integer, i As Integer, Numheads As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

n = FreeFile
Open App.Path & "\init\Cabezas.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , Numheads


ReDim HeadData(0 To Numheads + 1) As HeadData
ReDim Miscabezas(0 To Numheads + 1) As tIndiceCabeza

For i = 1 To Numheads
    Get #n, , Miscabezas(i)
    InitGrh HeadData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh HeadData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh HeadData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh HeadData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #n

End Sub

Sub CargarCascos()
On Error Resume Next
Dim n As Integer, i As Integer, NumCascos As Integer, Index As Integer

Dim Miscabezas() As tIndiceCabeza

n = FreeFile
Open App.Path & "\init\Cascos.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , NumCascos


ReDim CascoAnimData(0 To NumCascos + 1) As HeadData
ReDim Miscabezas(0 To NumCascos + 1) As tIndiceCabeza

For i = 1 To NumCascos
    Get #n, , Miscabezas(i)
    InitGrh CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0
    InitGrh CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0
    InitGrh CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0
    InitGrh CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0
Next i

Close #n

End Sub

Sub CargarCuerpos()
On Error Resume Next
Dim n As Integer, i As Integer
Dim NumCuerpos As Integer
Dim MisCuerpos() As tIndiceCuerpo

n = FreeFile
Open App.Path & "\init\Personajes.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , NumCuerpos


ReDim BodyData(0 To NumCuerpos + 1) As BodyData
ReDim MisCuerpos(0 To NumCuerpos + 1) As tIndiceCuerpo

For i = 1 To NumCuerpos
    Get #n, , MisCuerpos(i)
    InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
    InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
    InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
    InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
    BodyData(i).HeadOffset.x = MisCuerpos(i).HeadOffsetX
    BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
Next i

Close #n

End Sub
Sub CargarFxs()
On Error Resume Next
Dim n As Integer, i As Integer
Dim NumFxs As Integer
Dim MisFxs() As tIndiceFx

n = FreeFile
Open App.Path & "\init\Fxs.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , NumFxs


ReDim FxData(0 To NumFxs + 1) As FxData
ReDim MisFxs(0 To NumFxs + 1) As tIndiceFx

For i = 1 To NumFxs
    Get #n, , MisFxs(i)
    Call InitGrh(FxData(i).FX, MisFxs(i).Animacion, 1)
    FxData(i).OffsetX = MisFxs(i).OffsetX
    FxData(i).OffsetY = MisFxs(i).OffsetY
Next i

Close #n

End Sub
Sub CargarArrayLluvia()
On Error Resume Next
Dim n As Integer, i As Integer
Dim Nu As Integer

n = FreeFile
Open App.Path & "\init\fk.ind" For Binary Access Read As #n


Get #n, , MiCabecera


Get #n, , Nu


ReDim bLluvia(1 To Nu) As Byte

For i = 1 To Nu
    Get #n, , bLluvia(i)
Next i

Close #n

End Sub
Sub ConvertCPtoTP(StartPixelLeft As Integer, StartPixelTop As Integer, ByVal cx As Single, ByVal cy As Single, tX As Integer, tY As Integer)




Dim HWindowX As Integer
Dim HWindowY As Integer

cx = cx - StartPixelLeft
cy = cy - StartPixelTop

HWindowX = (WindowTileWidth \ 2)
HWindowY = (WindowTileHeight \ 2)


cx = (cx \ TilePixelWidth)
cy = (cy \ TilePixelHeight)

If cx > HWindowX Then
    cx = (cx - HWindowX)
Else
    If cx < HWindowX Then
        cx = (0 - (HWindowX - cx))
    Else
        cx = 0
    End If
End If

If cy > HWindowY Then
    cy = cy + 0 - HWindowY
Else
    If cy < HWindowY Then
        cy = (cy - HWindowY)
    Else
        cy = 0
    End If
End If

tX = UserPos.x + cx
tY = UserPos.Y + cy

End Sub
Sub MakeChar(ByVal charindex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal Y As Integer, ByVal arma As Integer, ByVal escudo As Integer, ByVal casco As Integer)
On Error Resume Next


If charindex > LastChar Then LastChar = charindex

NumChars = NumChars + 1

If arma = 0 Then arma = 2
If escudo = 0 Then escudo = 2
If casco = 0 Then casco = 2

CharList(charindex).Head = HeadData(Head)

CharList(charindex).Body = BodyData(Body)

If Body > 83 And Body < 88 Then
    CharList(charindex).Navegando = 1
Else: CharList(charindex).Navegando = 0
End If

CharList(charindex).arma = WeaponAnimData(arma)
    
CharList(charindex).escudo = ShieldAnimData(escudo)
CharList(charindex).casco = CascoAnimData(casco)

CharList(charindex).Heading = Heading


CharList(charindex).Moving = 0
CharList(charindex).MoveOffset.x = 0
CharList(charindex).MoveOffset.Y = 0


CharList(charindex).POS.x = x
CharList(charindex).POS.Y = Y


CharList(charindex).Active = 1


MapData(x, Y).charindex = charindex

End Sub

Sub ResetCharInfo(ByVal charindex As Integer)

CharList(charindex).Active = 0
CharList(charindex).Criminal = 0
CharList(charindex).FX = 0
CharList(charindex).FxLoopTimes = 0
CharList(charindex).invisible = False
CharList(charindex).Moving = 0
CharList(charindex).muerto = False
CharList(charindex).Nombre = ""
CharList(charindex).pie = False
CharList(charindex).POS.x = 0
CharList(charindex).POS.Y = 0
CharList(charindex).UsandoArma = False

End Sub

Sub EraseChar(ByVal charindex As Integer)
On Error Resume Next





CharList(charindex).Active = 0


If charindex = LastChar Then
    Do Until CharList(LastChar).Active = 1
        LastChar = LastChar - 1
        If LastChar = 0 Then Exit Do
    Loop
End If


MapData(CharList(charindex).POS.x, CharList(charindex).POS.Y).charindex = 0

Call ResetCharInfo(charindex)


NumChars = NumChars - 1

End Sub

Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional Started As Byte = 2)



If GrhIndex = 0 Then Exit Sub
Grh.GrhIndex = GrhIndex

If Started = 2 Then
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        Grh.Started = 1
    Else
        Grh.Started = 0
    End If
Else
    Grh.Started = Started
End If

Grh.FrameCounter = 1







If Grh.GrhIndex <> 0 Then Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed



End Sub

Sub MoveCharByHead(charindex As Integer, nheading As Byte)



Dim addX As Integer
Dim addY As Integer
Dim x As Integer
Dim Y As Integer
Dim nX As Integer
Dim nY As Integer

x = CharList(charindex).POS.x
Y = CharList(charindex).POS.Y


Select Case nheading

    Case NORTH
        addY = -1

    Case EAST
        addX = 1

    Case SOUTH
        addY = 1
    
    Case WEST
        addX = -1
        
End Select

nX = x + addX
nY = Y + addY

MapData(nX, nY).charindex = charindex
CharList(charindex).POS.x = nX
CharList(charindex).POS.Y = nY
MapData(x, Y).charindex = 0

CharList(charindex).MoveOffset.x = -1 * (TilePixelWidth * addX)
CharList(charindex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(charindex).scrollDirectionX = Sgn(addX)
CharList(charindex).scrollDirectionY = Sgn(addY)
CharList(charindex).Moving = 1
CharList(charindex).Heading = nheading

CharList(charindex).scrollDirectionX = addX
CharList(charindex).scrollDirectionY = addY
If UserEstado <> 1 Then Call DoPasosFx(charindex)


End Sub


Public Sub DoFogataFx()
If FX = 0 Then
    If bFogata Then
        bFogata = HayFogata()
        If Not bFogata Then frmPrincipal.StopSound
    Else
        bFogata = HayFogata()
        If bFogata Then frmPrincipal.Play "fuego.wav", True
    End If
End If
End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean

Dim x As Integer, Y As Integer

For Y = UserPos.Y - MinYBorder + 1 To UserPos.Y + MinYBorder - 1
  For x = UserPos.x - MinXBorder + 1 To UserPos.x + MinXBorder - 1
            
            If MapData(x, Y).charindex = Index2 Then
                EstaPCarea = True
                Exit Function
            End If
        
  Next x
Next Y

EstaPCarea = False

End Function
Public Function TickON(Cual As Integer, Cont As Integer) As Boolean
Static TickCount(200) As Integer
If Cont = 999 Then Exit Function
TickCount(Cual) = TickCount(Cual) + 1
If TickCount(Cual) < Cont Then
    TickON = False
Else
    TickCount(Cual) = 0
    TickON = True
End If
End Function
Sub DoPasosFx(ByVal charindex As Integer)
Static pie As Boolean

If CharList(charindex).Navegando = 0 Then
    If UserMontando And EstaPCarea(charindex) And charindex = UserCharIndex Then
        If TickON(0, 4) Then Call PlayWaveDS(SND_MONTANDO)
    Else
        If CharList(charindex).Criminal = 1 Then Exit Sub
        If Not CharList(charindex).muerto And EstaPCarea(charindex) Then
            CharList(charindex).pie = Not CharList(charindex).pie
            If CharList(charindex).pie Then
                Call PlayWaveDS(SND_PASOS1)
            Else
                Call PlayWaveDS(SND_PASOS2)
            End If
        End If
    End If
Else: Call PlayWaveDS(SND_NAVEGANDO)
End If

End Sub
Sub MoveCharByPosAndHead(charindex As Integer, nX As Integer, nY As Integer, nheading As Byte)

On Error Resume Next

Dim x As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer



x = CharList(charindex).POS.x
Y = CharList(charindex).POS.Y

MapData(x, Y).charindex = 0

addX = nX - x
addY = nY - Y




MapData(nX, nY).charindex = charindex


CharList(charindex).POS.x = nX
CharList(charindex).POS.Y = nY

CharList(charindex).MoveOffset.x = -1 * (TilePixelWidth * addX)
CharList(charindex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(charindex).scrollDirectionX = Sgn(addX)
CharList(charindex).scrollDirectionY = Sgn(addY)
CharList(charindex).Moving = 1
CharList(charindex).Heading = nheading




End Sub
Sub MoveCharByPos(charindex As Integer, nX As Integer, nY As Integer)
On Error Resume Next

Dim x As Integer
Dim Y As Integer
Dim addX As Integer
Dim addY As Integer
Dim nheading As Byte

x = CharList(charindex).POS.x
Y = CharList(charindex).POS.Y

MapData(x, Y).charindex = 0

addX = nX - x
addY = nY - Y


If Sgn(addX) = 1 Then nheading = EAST
If Sgn(addX) = -1 Then nheading = WEST
If Sgn(addY) = -1 Then nheading = NORTH
If Sgn(addY) = 1 Then nheading = SOUTH

MapData(nX, nY).charindex = charindex

CharList(charindex).POS.x = nX
CharList(charindex).POS.Y = nY

CharList(charindex).MoveOffset.x = -1 * (TilePixelWidth * addX)
CharList(charindex).MoveOffset.Y = -1 * (TilePixelHeight * addY)

CharList(charindex).scrollDirectionX = Sgn(addX)
CharList(charindex).scrollDirectionY = Sgn(addY)
CharList(charindex).Moving = 1
CharList(charindex).Heading = nheading

End Sub
Sub MoveCharByPosConHeading(charindex As Integer, nX As Integer, nY As Integer, nheading As Byte)
On Error Resume Next

If InMapBounds(CharList(charindex).POS.x, CharList(charindex).POS.Y) Then MapData(CharList(charindex).POS.x, CharList(charindex).POS.Y).charindex = 0

MapData(nX, nY).charindex = charindex

CharList(charindex).POS.x = nX
CharList(charindex).POS.Y = nY

CharList(charindex).Moving = 0
CharList(charindex).MoveOffset.x = 0
CharList(charindex).MoveOffset.Y = 0

CharList(charindex).Heading = nheading

End Sub

Sub MoveScreen(Heading As Byte)



Dim x As Integer
Dim Y As Integer
Dim tX As Integer
Dim tY As Integer


Select Case Heading

    Case NORTH
        Y = -1

    Case EAST
        x = 1

    Case SOUTH
        Y = 1
    
    Case WEST
        x = -1
        
End Select


tX = UserPos.x + x
tY = UserPos.Y + Y


If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
    Exit Sub
Else
    
    AddtoUserPos.x = x
    UserPos.x = tX
    AddtoUserPos.Y = Y
    UserPos.Y = tY
    UserMoving = 1
   
End If


    

End Sub


Function HayFogata() As Boolean
Dim j As Integer, k As Integer
For j = UserPos.x - 8 To UserPos.x + 8
    For k = UserPos.Y - 6 To UserPos.Y + 6
        If InMapBounds(j, k) Then
            If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    HayFogata = True
                    Exit Function
            End If
        End If
    Next k
Next j
End Function
Private Function AmigoClan(ByVal charindex As Integer) As Boolean
Dim Nombre1 As String
Dim Nombre2 As String
 
Nombre1 = CharList(UserCharIndex).Nombre
Nombre2 = CharList(charindex).Nombre
 
If InStr(Nombre1, "<") > 0 And InStr(Nombre2, "<") > 0 Then
 
AmigoClan = Trim$(Mid$(Nombre2, InStr(Nombre2, "<"))) = _
                Trim$(Mid$(Nombre1, InStr(Nombre1, "<")))
             End If
End Function

Function NextOpenChar() As Integer
Dim loopc As Integer

loopc = 1
Do While CharList(loopc).Active
    loopc = loopc + 1
Loop

NextOpenChar = loopc

End Function
Sub LoadGrhData()
On Error GoTo ErrorHandler

Dim Grh As Integer
Dim Frame As Integer
Dim tempint As Integer


ReDim GrhData(1 To Config_Inicio.NumeroDeBMPs) As GrhData

Open IniPath & "Graficos.ind" For Binary Access Read As #1
Seek #1, 1

Get #1, , MiCabecera
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint
Get #1, , tempint

Get #1, , Grh

Do Until Grh <= 0
    
    
    Get #1, , GrhData(Grh).NumFrames
    If GrhData(Grh).NumFrames <= 0 Then GoTo ErrorHandler
    
    If GrhData(Grh).NumFrames > 1 Then
    
        
        For Frame = 1 To GrhData(Grh).NumFrames
        
            Get #1, , GrhData(Grh).Frames(Frame)
            If GrhData(Grh).Frames(Frame) <= 0 Or GrhData(Grh).Frames(Frame) > Config_Inicio.NumeroDeBMPs Then
                GoTo ErrorHandler
            End If
        
        Next Frame
    
        Get #1, , GrhData(Grh).Speed
        If GrhData(Grh).Speed <= 0 Then GoTo ErrorHandler
        
        
        GrhData(Grh).pixelHeight = GrhData(GrhData(Grh).Frames(1)).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).pixelWidth = GrhData(GrhData(Grh).Frames(1)).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileWidth = GrhData(GrhData(Grh).Frames(1)).TileWidth
        If GrhData(Grh).TileWidth <= 0 Then GoTo ErrorHandler
        
        GrhData(Grh).TileHeight = GrhData(GrhData(Grh).Frames(1)).TileHeight
        If GrhData(Grh).TileHeight <= 0 Then GoTo ErrorHandler
    
    Else
    
        
        Get #1, , GrhData(Grh).FileNum
        If GrhData(Grh).FileNum <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sX
        If GrhData(Grh).sX < 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).sY
        If GrhData(Grh).sY < 0 Then GoTo ErrorHandler
            
        Get #1, , GrhData(Grh).pixelWidth
        If GrhData(Grh).pixelWidth <= 0 Then GoTo ErrorHandler
        
        Get #1, , GrhData(Grh).pixelHeight
        If GrhData(Grh).pixelHeight <= 0 Then GoTo ErrorHandler
        
        
        GrhData(Grh).TileWidth = GrhData(Grh).pixelWidth / TilePixelHeight
        GrhData(Grh).TileHeight = GrhData(Grh).pixelHeight / TilePixelWidth
        
        GrhData(Grh).Frames(1) = Grh
            
    End If
    
    
    Get #1, , Grh

Loop


Close #1

Exit Sub

ErrorHandler:
Close #1
MsgBox "Error while loading the Grh.dat! Stopped at GRH number: " & Grh

End Sub

Function LegalPos(x As Integer, Y As Integer) As Boolean





If x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPos = False
    Exit Function
End If

    
    If MapData(x, Y).Blocked = 1 Then
        LegalPos = False
        Exit Function
    End If
    
    
    If MapData(x, Y).charindex > 0 Then
        LegalPos = False
        Exit Function
    End If
   
    If Not UserNavegando Then
        If HayAgua(x, Y) Then
            LegalPos = False
            Exit Function
        End If
    Else
        If Not HayAgua(x, Y) Then
            LegalPos = False
            Exit Function
        End If
    End If
    
LegalPos = True

End Function

Function LegalPosMuerto(x As Integer, Y As Integer) As Boolean





If x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    LegalPosMuerto = False
    Exit Function
End If

    
    If MapData(x, Y).Blocked = 1 Then
        LegalPosMuerto = False
        Exit Function
    End If
    
    
    If MapData(x, Y).charindex > 0 Then
    If CharList(MapData(x, Y).charindex).muerto = True Then
        LegalPosMuerto = False
        Exit Function
    End If
    End If
   
    If Not UserNavegando Then
        If HayAgua(x, Y) Then
            LegalPosMuerto = False
            Exit Function
        End If
    Else
        If Not HayAgua(x, Y) Then
            LegalPosMuerto = False
            Exit Function
        End If
    End If
    
LegalPosMuerto = True

End Function




Function InMapLegalBounds(x As Integer, Y As Integer) As Boolean





If x < MinXBorder Or x > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
    InMapLegalBounds = False
    Exit Function
End If

InMapLegalBounds = True

End Function

Function InMapBounds(x As Integer, Y As Integer) As Boolean




If x < XMinMapSize Or x > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
    InMapBounds = False
    Exit Function
End If

InMapBounds = True

End Function

Sub DDrawGrhtoSurface(surface As DirectDrawSurface7, Grh As Grh, ByVal x As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)

Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + (1 / (8 / Velocidad))
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                End If
            End If
        End If
    End If
End If

CurrentGrh.GrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

If center Then
    If GrhData(CurrentGrh.GrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(CurrentGrh.GrhIndex).TileWidth * 16) + 16
    End If
    If GrhData(CurrentGrh.GrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(CurrentGrh.GrhIndex).TileHeight * 32) + 32
    End If
End If
With SourceRect
        .Left = GrhData(CurrentGrh.GrhIndex).sX
        .Top = GrhData(CurrentGrh.GrhIndex).sY
        .Right = .Left + GrhData(CurrentGrh.GrhIndex).pixelWidth
        .Bottom = .Top + GrhData(CurrentGrh.GrhIndex).pixelHeight
End With
surface.BltFast x, Y, SurfaceDB(GrhData(CurrentGrh.GrhIndex).FileNum), SourceRect, DDBLTFAST_WAIT
End Sub

Sub DDrawTransGrhIndextoSurface(surface As DirectDrawSurface7, Grh As Integer, ByVal x As Integer, ByVal Y As Integer, center As Byte, Animate As Byte)
Dim CurrentGrh As Grh
Dim destRect As RECT
Dim SourceRect As RECT
Dim SurfaceDesc As DDSURFACEDESC2

With destRect
    .Left = x
    .Top = Y
    .Right = .Left + GrhData(Grh).pixelWidth
    .Bottom = .Top + GrhData(Grh).pixelHeight
End With

surface.GetSurfaceDesc SurfaceDesc


If destRect.Left >= 0 And destRect.Top >= 0 And destRect.Right <= SurfaceDesc.lWidth And destRect.Bottom <= SurfaceDesc.lHeight Then
    With SourceRect
        .Left = GrhData(Grh).sX
        .Top = GrhData(Grh).sY
        .Right = .Left + GrhData(Grh).pixelWidth
        .Bottom = .Top + GrhData(Grh).pixelHeight
    End With
    
    surface.BltFast destRect.Left, destRect.Top, SurfaceDB.GetBMP(GrhData(Grh).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End If

End Sub



    Sub DDrawTransGrhtoSurface(surface As DirectDrawSurface7, Grh As Grh, ByVal x As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0)











Dim iGrhIndex As Integer

Dim SourceRect As RECT

Dim QuitarAnimacion As Boolean


If Animate Then
    If Grh.Started = 1 Then
        If Grh.SpeedCounter > 0 Then
            Grh.SpeedCounter = Grh.SpeedCounter - 1
            If Grh.SpeedCounter = 0 Then
                Grh.SpeedCounter = GrhData(Grh.GrhIndex).Speed
                Grh.FrameCounter = Grh.FrameCounter + (1 / (8 / Velocidad))
                If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                    Grh.FrameCounter = 1
                    If KillAnim Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                            
                            If CharList(KillAnim).FxLoopTimes > 0 Then CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes < 1 Then
                                CharList(KillAnim).FX = 0
                                Exit Sub
                            End If
                            
                        End If
                    End If
               End If
            End If
        End If
    End If
End If

If Grh.GrhIndex = 0 Then Exit Sub


iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)


If center Then
    If GrhData(iGrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
    End If
    If GrhData(iGrhIndex).TileHeight <> 1 Then
        Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
    End If
End If

With SourceRect
    .Left = GrhData(iGrhIndex).sX
    .Top = GrhData(iGrhIndex).sY
    .Right = .Left + GrhData(iGrhIndex).pixelWidth
    .Bottom = .Top + GrhData(iGrhIndex).pixelHeight
End With

surface.BltFast x, Y, SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum), SourceRect, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT

End Sub
Sub DrawBackBufferSurface()

Call PrimarySurface.Blt(MainViewRect, BackBufferSurface, MainDestRect, DDBLT_WAIT)

End Sub
Function GetBitmapDimensions(BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
Dim BMHeader As BITMAPFILEHEADER
Dim BINFOHeader As BITMAPINFOHEADER

Open BmpFile For Binary Access Read As #1
Get #1, , BMHeader
Get #1, , BINFOHeader
Close #1
bmWidth = BINFOHeader.biWidth
bmHeight = BINFOHeader.biHeight
End Function
Sub DrawGrhtoHdc(hwnd As Long, Hdc As Long, Grh As Integer, SourceRect As RECT, destRect As RECT)

If Grh <= 0 Then Exit Sub

SecundaryClipper.SetHWnd hwnd
SurfaceDB.GetBMP(GrhData(Grh).FileNum).BltToDC Hdc, SourceRect, destRect

End Sub
Sub PlayWaveAPI(File As String)
Dim rc As Integer

rc = sndPlaySound(File, SND_ASYNC)

End Sub
Sub RenderScreen(tilex As Integer, tiley As Integer, PixelOffsetX As Double, PixelOffsetY As Double)
On Error Resume Next
If UserCiego Then Exit Sub
Dim Y        As Integer
Dim x        As Integer
Dim minY     As Integer
Dim maxY     As Integer
Dim minX     As Integer
Dim maxX     As Integer
Dim ScreenX  As Integer
Dim ScreenY  As Integer
Dim Moved    As Byte
Dim Grh      As Grh
Dim tempchar As Char
Dim TextX    As Integer
Dim TextY    As Integer
Dim iPPx     As Integer
Dim iPPy     As Integer
Dim rSourceRect      As RECT
Dim iGrhIndex        As Integer
Dim PixelOffsetXTemp As Integer
Dim PixelOffsetYTemp As Integer
Dim Tiempo As Double
 
minY = (tiley - 15)
maxY = (tiley + 15)
minX = (tilex - 17)
maxX = (tilex + 17)
 
Tiempo = GetTickCount
ScreenY = 8 + RenderMod.iImageSize
For Y = (minY + 8) + RenderMod.iImageSize To (maxY - 8) - RenderMod.iImageSize
    ScreenX = 8 + RenderMod.iImageSize
    For x = (minX + 8) + RenderMod.iImageSize To (maxX - 8) - RenderMod.iImageSize
        If x > 100 Or Y < 1 Then Exit For
        
        With MapData(x, Y).Graphic(1)
            If (.Started = 1) Then
                If (.SpeedCounter > 0) Then
                    .SpeedCounter = .SpeedCounter - 1
                    If (.SpeedCounter = 0) Then
                        .SpeedCounter = GrhData(.GrhIndex).Speed
                        .FrameCounter = .FrameCounter + 1
                        If (.FrameCounter > GrhData(.GrhIndex).NumFrames) Then .FrameCounter = 1
                    End If
                End If
            End If
 
            iGrhIndex = GrhData(.GrhIndex).Frames(.FrameCounter)
        End With
 
        rSourceRect.Left = GrhData(iGrhIndex).sX
        rSourceRect.Top = GrhData(iGrhIndex).sY
        rSourceRect.Right = rSourceRect.Left + GrhData(iGrhIndex).pixelWidth
        rSourceRect.Bottom = rSourceRect.Top + GrhData(iGrhIndex).pixelHeight



        Call BackBufferSurface.BltFast(((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, SurfaceDB.GetBMP(GrhData(iGrhIndex).FileNum), rSourceRect, DDBLTFAST_WAIT)
        
        If Not RenderMod.bNoCostas And MapData(x, Y).Graphic(2).GrhIndex <> 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(x, Y).Graphic(2), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 1)
        ScreenX = ScreenX + 1
    Next x
    ScreenY = ScreenY + 1
    If Y > 100 Then Exit For
Next Y
  
ScreenY = 8 + RenderMod.iImageSize
For Y = (minY + 8) + RenderMod.iImageSize To (maxY - 1) - RenderMod.iImageSize
    ScreenX = 5 + RenderMod.iImageSize
     For x = (minX + 5) + RenderMod.iImageSize To (maxX - 5) - RenderMod.iImageSize
        If Not (x > 100 Or x < 1) Then
            iPPx = ((32 * ScreenX) - 32) + PixelOffsetX
            iPPy = ((32 * ScreenY) - 32) + PixelOffsetY
    
            If MapData(x, Y).ObjGrh.GrhIndex <> 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(x, Y).ObjGrh, iPPx, iPPy, 1, 1)
            
            If MapData(x, Y).charindex > 0 Then
                tempchar = CharList(MapData(x, Y).charindex)
                PixelOffsetXTemp = PixelOffsetX
                PixelOffsetYTemp = PixelOffsetY
                Moved = 0
    
            'If needed, move left and right
            If tempchar.scrollDirectionX <> 0 Then
                tempchar.Body.Walk(tempchar.Heading).Started = 1
                tempchar.arma.WeaponWalk(tempchar.Heading).Started = 1
                tempchar.escudo.ShieldWalk(tempchar.Heading).Started = 1
                PixelOffsetXTemp = PixelOffsetXTemp + tempchar.MoveOffset.x
                tempchar.MoveOffset.x = tempchar.MoveOffset.x - (Velocidad * Sgn(tempchar.MoveOffset.x))
                'tempChar.MoveOffset.X = tempChar.MoveOffset.X - (2 * Sgn(tempChar.MoveOffset.X))
                Moved = 1
               
                'Check if we already got there
                If (Sgn(tempchar.scrollDirectionX) = 1 And tempchar.MoveOffset.x >= 0) Or _
                        (Sgn(tempchar.scrollDirectionX) = -1 And tempchar.MoveOffset.x <= 0) Then
                    tempchar.MoveOffset.x = 0
                    tempchar.scrollDirectionX = 0
                End If
            End If
 
            'If needed, move up and down
            If tempchar.scrollDirectionY <> 0 Then
                tempchar.Body.Walk(tempchar.Heading).Started = 1
                tempchar.arma.WeaponWalk(tempchar.Heading).Started = 1
                tempchar.escudo.ShieldWalk(tempchar.Heading).Started = 1
                PixelOffsetYTemp = PixelOffsetYTemp + tempchar.MoveOffset.Y
                tempchar.MoveOffset.Y = tempchar.MoveOffset.Y - (Velocidad * Sgn(tempchar.MoveOffset.Y))
                Moved = 1
                If (Sgn(tempchar.scrollDirectionY) = 1 And tempchar.MoveOffset.Y >= 0) Or _
                        (Sgn(tempchar.scrollDirectionY) = -1 And tempchar.MoveOffset.Y <= 0) Then
                    tempchar.MoveOffset.Y = 0
                    tempchar.scrollDirectionY = 0
                End If
            End If
 
            If Moved = 0 And tempchar.Moving = 1 Then
                tempchar.Moving = 0
                tempchar.Body.Walk(tempchar.Heading).FrameCounter = 1
                tempchar.Body.Walk(tempchar.Heading).Started = 0
                tempchar.arma.WeaponWalk(tempchar.Heading).FrameCounter = 1
                tempchar.arma.WeaponWalk(tempchar.Heading).Started = 0
                tempchar.escudo.ShieldWalk(tempchar.Heading).FrameCounter = 1
                tempchar.escudo.ShieldWalk(tempchar.Heading).Started = 0
                tempchar.haciendoataque = 0
            End If
            
            If tempchar.haciendoataque = 0 And tempchar.MoveOffset.x = 0 And tempchar.MoveOffset.Y = 0 Then
                tempchar.arma.WeaponWalk(tempchar.Heading).Started = 0
                tempchar.arma.WeaponWalk(tempchar.Heading).FrameCounter = 1
                End If
            If tempchar.haciendoataque = 1 Then
                tempchar.arma.WeaponWalk(tempchar.Heading).Started = 1
                tempchar.haciendoataque = 0
            End If
            
                iPPx = ((32 * ScreenX) - 32) + PixelOffsetXTemp
                iPPy = ((32 * ScreenY) - 32) + PixelOffsetYTemp
                
                If Len(tempchar.Nombre) = 0 Then
                    Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.Body.Walk(tempchar.Heading), iPPx, iPPy, 1, 1)
                    If tempchar.Head.Head(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.Head.Head(tempchar.Heading), iPPx + tempchar.Body.HeadOffset.x, iPPy + tempchar.Body.HeadOffset.Y, 1, 0)
                Else
                     If tempchar.Navegando = 1 Then
                        Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.Body.Walk(tempchar.Heading), iPPx, iPPy, 1, 1)
                    ElseIf Not CharList(MapData(x, Y).charindex).invisible And tempchar.Head.Head(tempchar.Heading).GrhIndex > 0 Then
                        Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.Body.Walk(tempchar.Heading), iPPx, iPPy, 1, 1)
                        If tempchar.Head.Head(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.Head.Head(tempchar.Heading), iPPx + tempchar.Body.HeadOffset.x, iPPy + tempchar.Body.HeadOffset.Y, 1, 0)
                        If tempchar.casco.Head(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.casco.Head(tempchar.Heading), iPPx + tempchar.Body.HeadOffset.x, iPPy + tempchar.Body.HeadOffset.Y, 1, 0)
                        If tempchar.arma.WeaponWalk(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.arma.WeaponWalk(tempchar.Heading), iPPx, iPPy, 1, 1)
                        If tempchar.escudo.ShieldWalk(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.escudo.ShieldWalk(tempchar.Heading), iPPx, iPPy, 1, 1)
                       ElseIf CharList(MapData(x, Y).charindex).invisible And (CharList(MapData(x, Y).charindex).Nombre = CharList(UserCharIndex).Nombre Or AmigoClan(MapData(x, Y).charindex)) Then
                        Call DDrawBlendGrhtoSurface(BackBufferSurface, tempchar.Body.Walk(tempchar.Heading), iPPx, iPPy, 1, 1)
                        If tempchar.Head.Head(tempchar.Heading).GrhIndex > 0 Then Call DDrawBlendGrhtoSurface(BackBufferSurface, tempchar.Head.Head(tempchar.Heading), iPPx + tempchar.Body.HeadOffset.x, iPPy + tempchar.Body.HeadOffset.Y, 1, 0)
                        If tempchar.casco.Head(tempchar.Heading).GrhIndex > 0 Then Call DDrawBlendGrhtoSurface(BackBufferSurface, tempchar.casco.Head(tempchar.Heading), iPPx + tempchar.Body.HeadOffset.x, iPPy + tempchar.Body.HeadOffset.Y, 1, 0)
                        If tempchar.arma.WeaponWalk(tempchar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, tempchar.arma.WeaponWalk(tempchar.Heading), iPPx, iPPy, 1, 1)
                        If tempchar.escudo.ShieldWalk(tempchar.Heading).GrhIndex > 0 Then Call DDrawBlendGrhtoSurface(BackBufferSurface, tempchar.escudo.ShieldWalk(tempchar.Heading), iPPx, iPPy, 1, 1)
                        End If
                 
                                      Dim sClan As String
                       Dim lCenter As Long
                       
            
                     
      
                      If Nombres Then
                        
                        If Not (tempchar.invisible Or tempchar.Navegando = 1) Then
                       
                          
                            If InStr(tempchar.Nombre, "<") > 0 And InStr(tempchar.Nombre, ">") > 0 Then
                           
                                lCenter = (frmPrincipal.TextWidth(Left$(tempchar.Nombre, InStr(tempchar.Nombre, "<") - 1)) / 2) - 16
                                sClan = Mid$(tempchar.Nombre, InStr(tempchar.Nombre, "<"))
                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, Left$(tempchar.Nombre, InStr(tempchar.Nombre, "<") - 1), RGB(RG(tempchar.Criminal, 1), RG(tempchar.Criminal, 2), RG(tempchar.Criminal, 3)))
                                lCenter = (frmPrincipal.TextWidth(sClan) / 2) - 16
                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 45, sClan, RGB(RG(tempchar.Criminal, 1), RG(tempchar.Criminal, 2), RG(tempchar.Criminal, 3)))
                            Else
                                lCenter = (frmPrincipal.TextWidth(tempchar.Nombre) / 2) - 16
                                Call Dialogos.DrawText(iPPx - lCenter, iPPy + 30, tempchar.Nombre, RGB(RG(tempchar.Criminal, 1), RG(tempchar.Criminal, 2), RG(tempchar.Criminal, 3)))
                            End If
                      
                        End If
                       
                    End If
                End If
    
                If Dialogos.CantidadDialogos > 0 Then Call Dialogos.Update_Dialog_Pos((iPPx + tempchar.Body.HeadOffset.x), (iPPy + tempchar.Body.HeadOffset.Y), MapData(x, Y).charindex)
                
                CharList(MapData(x, Y).charindex) = tempchar
    
                If CharList(MapData(x, Y).charindex).FX <> 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, FxData(tempchar.FX).FX, iPPx + FxData(tempchar.FX).OffsetX, iPPy + FxData(tempchar.FX).OffsetY, 1, 1, MapData(x, Y).charindex)
                
            End If

        End If
        If MapData(x, Y).Graphic(3).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(x, Y).Graphic(3), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 1)
        ScreenX = ScreenX + 1
    Next x
    ScreenY = ScreenY + 1
    If Y >= 100 Or Y < 1 Then Exit For
Next Y

If Not bTecho Then
    ScreenY = 5 + RenderMod.iImageSize
    For Y = (minY + 5) + RenderMod.iImageSize To (maxY - 1) - RenderMod.iImageSize
        ScreenX = 5 + RenderMod.iImageSize
        For x = (minX + 5) + RenderMod.iImageSize To (maxX - 0) - RenderMod.iImageSize
            
            If x < 101 And x > 0 And Y < 101 And Y > 0 Then
                If MapData(x, Y).Graphic(4).GrhIndex <> 0 Then Call DDrawTransGrhtoSurface(BackBufferSurface, MapData(x, Y).Graphic(4), ((32 * ScreenX) - 32) + PixelOffsetX, ((32 * ScreenY) - 32) + PixelOffsetY, 1, 1)
            End If
            
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next Y
End If
 If frmPrincipal.Timer6.Enabled = True Then
  Call Dialogos.Conteosc(400, 305, lardatac, vbRed)
End If
  
 If frmPrincipal.Timer7.Enabled = True Then
 Call Dialogos.Conteosb(400, 290, lardatab, vbRed)
 End If
  If frmPrincipal.Timer8.Enabled = True Then
 Call Dialogos.Conteosa(400, 275, lardataa, vbRed)
End If

If bLluvia(UserMap) = 1 Then
    If bRain Or bRainST Then
        If llTick < DirectX.TickCount - 50 Then
            iFrameIndex = iFrameIndex + 1
            If iFrameIndex > 7 Then iFrameIndex = 0
            llTick = DirectX.TickCount
        End If
 
        For Y = 0 To 4
            For x = 0 To 4
                Call BackBufferSurface.BltFast(LTLluvia(Y), LTLluvia(x), SurfaceDB.GetBMP(5556), RLluvia(iFrameIndex), DDBLTFAST_SRCCOLORKEY + DDBLTFAST_WAIT)
            Next x
        Next Y
    End If
End If

 
End Sub
Public Function RenderSounds()

    If bLluvia(UserMap) = 1 Then
        If bRain Then
            If bTecho Then
                If frmPrincipal.IsPlaying <> plLluviain Then
                    Call frmPrincipal.StopSound
                    Call frmPrincipal.Play("lluviain.wav", True)
                    frmPrincipal.IsPlaying = plLluviain
                End If
                
                
            Else
                If frmPrincipal.IsPlaying <> plLluviaout Then
                    Call frmPrincipal.StopSound
                    Call frmPrincipal.Play("lluviaout.wav", True)
                    frmPrincipal.IsPlaying = plLluviaout
                End If
                
                
            End If
        End If
    End If

End Function


Function HayUserAbajo(ByVal x As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean

If GrhIndex > 0 Then
        
        HayUserAbajo = _
            CharList(UserCharIndex).POS.x >= x - (GrhData(GrhIndex).TileWidth \ 2) _
        And CharList(UserCharIndex).POS.x <= x + (GrhData(GrhIndex).TileWidth \ 2) _
        And CharList(UserCharIndex).POS.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
        And CharList(UserCharIndex).POS.Y <= Y
        
End If

End Function



Function PixelPos(x As Integer) As Integer




PixelPos = (TilePixelWidth * x) - TilePixelWidth

End Function


Sub LoadGraphics()
        Dim loopc As Integer
        Dim SurfaceDesc As DDSURFACEDESC2
        Dim ddck As DDCOLORKEY
        Dim ddsd As DDSURFACEDESC2
        Dim iLoopUpdate As Integer

        SurfaceDB.MaxEntries = 150
        SurfaceDB.lpDirectDraw7 = DirectDraw
        SurfaceDB.Path = DirGraficos
        Call SurfaceDB.Init(mododinamico)

        
        Call GetBitmapDimensions(DirGraficos & "5556.bmp", ddsd.lWidth, ddsd.lHeight)
              
        RLluvia(0).Top = 0:      RLluvia(1).Top = 0:      RLluvia(2).Top = 0:      RLluvia(3).Top = 0
        RLluvia(0).Left = 0:     RLluvia(1).Left = 128:   RLluvia(2).Left = 256:   RLluvia(3).Left = 384
        RLluvia(0).Right = 128:  RLluvia(1).Right = 256:  RLluvia(2).Right = 384:  RLluvia(3).Right = 512
        RLluvia(0).Bottom = 128: RLluvia(1).Bottom = 128: RLluvia(2).Bottom = 128: RLluvia(3).Bottom = 128
    
    
        RLluvia(4).Top = 128:    RLluvia(5).Top = 128:    RLluvia(6).Top = 128:    RLluvia(7).Top = 128
        RLluvia(4).Left = 0:     RLluvia(5).Left = 128:   RLluvia(6).Left = 256:   RLluvia(7).Left = 384
        RLluvia(4).Right = 128:  RLluvia(5).Right = 256:  RLluvia(6).Right = 384:  RLluvia(7).Right = 512
        RLluvia(4).Bottom = 256: RLluvia(5).Bottom = 256: RLluvia(6).Bottom = 256: RLluvia(7).Bottom = 256
        AddtoRichTextBox frmCargando.Status, "Hecho.", 255, 150, 50, 1, , False
End Sub



Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setMainViewTop As Integer, setMainViewLeft As Integer, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, setTileBufferSize As Integer) As Boolean





Dim SurfaceDesc As DDSURFACEDESC2
Dim ddck As DDCOLORKEY

IniPath = App.Path & "\Init\"


UserPos.x = MinXBorder
UserPos.Y = MinYBorder



DisplayFormhWnd = setDisplayFormhWnd
MainViewTop = setMainViewTop
MainViewLeft = setMainViewLeft
TilePixelWidth = setTilePixelWidth
TilePixelHeight = setTilePixelHeight
WindowTileHeight = setWindowTileHeight
WindowTileWidth = setWindowTileWidth
TileBufferSize = setTileBufferSize

MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

MainViewWidth = (TilePixelWidth * WindowTileWidth)
MainViewHeight = (TilePixelHeight * WindowTileHeight)


ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock





DirectDraw.SetCooperativeLevel DisplayFormhWnd, DDSCL_NORMAL

If Musica = 0 Or FX = 0 Then
    DirectSound.SetCooperativeLevel DisplayFormhWnd, DSSCL_PRIORITY
End If



With SurfaceDesc
    .lFlags = DDSD_CAPS
    .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
End With



Set PrimarySurface = DirectDraw.CreateSurface(SurfaceDesc)

Set PrimaryClipper = DirectDraw.CreateClipper(0)
PrimaryClipper.SetHWnd frmPrincipal.hwnd
PrimarySurface.SetClipper PrimaryClipper

Set SecundaryClipper = DirectDraw.CreateClipper(0)

With BackBufferRect
    .Left = 0 + 32 * RenderMod.iImageSize
    .Top = 0 + 32 * RenderMod.iImageSize
    .Right = (TilePixelWidth * (WindowTileWidth + (2 * TileBufferSize))) - 32 * RenderMod.iImageSize
    .Bottom = (TilePixelHeight * (WindowTileHeight + (2 * TileBufferSize))) - 32 * RenderMod.iImageSize
End With
With SurfaceDesc
    .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    If RenderMod.bUseVideo Then
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Else
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    End If
    .lHeight = BackBufferRect.Bottom
    .lWidth = BackBufferRect.Right
End With

Set BackBufferSurface = DirectDraw.CreateSurface(SurfaceDesc)

ddck.low = 0
ddck.high = 0
BackBufferSurface.SetColorKey DDCKEY_SRCBLT, ddck
Call InitBlend(BackBufferSurface)

Call LoadGrhData
Call CargarCuerpos
Call CargarCabezas
Call CargarCascos
Call CargarFxs


LTLluvia(0) = 224
LTLluvia(1) = 352
LTLluvia(2) = 480
LTLluvia(3) = 608
LTLluvia(4) = 736

AddtoRichTextBox frmCargando.Status, "Cargando Gráficos....", 255, 150, 50, , , True
Call LoadGraphics

InitTileEngine = True
End Function





Sub ShowNextFrame()












    Static OffsetCounterX As Double
    Static OffsetCounterY As Double

    If EngineRun Then
        
            If AddtoUserPos.x <> 0 Then
                OffsetCounterX = (OffsetCounterX - (Velocidad * Sgn(AddtoUserPos.x)))
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.x) Then
                    OffsetCounterX = 0
                    AddtoUserPos.x = 0
                    UserMoving = 0
                End If
            

            
            
            ElseIf AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - (Velocidad * Sgn(AddtoUserPos.Y))
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = 0
                End If
            End If

            
            Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY)
            If ModoTrabajo Then Call Dialogos.DrawText(260, 260, "MODO TRABAJO", vbRed)
            
            Call DibujarCartel
            
            Call DrawBackBufferSurface
            
            
            FramesPerSecCounter = FramesPerSecCounter + 1
    End If
End Sub

Sub CrearGrh(GrhIndex As Integer, Index As Integer)
ReDim Preserve Grh(1 To Index) As Grh
Grh(Index).FrameCounter = 1
Grh(Index).GrhIndex = GrhIndex
Grh(Index).SpeedCounter = GrhData(GrhIndex).Speed
Grh(Index).Started = 1
End Sub

Sub CargarAnimsExtra()
Call CrearGrh(6580, 1)
Call CrearGrh(534, 2)
End Sub

Function ControlVelocidad(ByVal lastTime As Long) As Boolean
ControlVelocidad = (GetTickCount - lastTime > 20)
End Function
