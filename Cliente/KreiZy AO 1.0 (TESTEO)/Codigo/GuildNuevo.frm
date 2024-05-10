VERSION 5.00
Begin VB.Form frmGuildsNuevo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   Picture         =   "GuildNuevo.frx":0000
   ScaleHeight     =   7650
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000008&
      Caption         =   "Información"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   6840
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   600
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3240
      Width           =   6495
   End
   Begin VB.ListBox MembersList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   6495
   End
   Begin VB.ListBox GuildList 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   600
      TabIndex        =   0
      Top             =   5760
      Width           =   6495
   End
   Begin VB.Image command5 
      Height          =   375
      Left            =   3000
      MouseIcon       =   "GuildNuevo.frx":DA24
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Image command4 
      Height          =   255
      Left            =   3000
      MouseIcon       =   "GuildNuevo.frx":DD2E
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Image command8 
      Height          =   255
      Left            =   120
      MouseIcon       =   "GuildNuevo.frx":E038
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   735
   End
End
Attribute VB_Name = "frmGuildsNuevo"
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

Public Function ListaDeClanes(ByVal Data As String) As Integer
Dim a As Integer
Dim i As Integer

a = Val(ReadField(1, Data, Asc("¬")))
ReDim oClan(1 To a) As Clan

For i = 1 To a
    oClan(i).name = Left$(ReadField(i + 1, Data, Asc("¬")), Len(ReadField(i + 1, Data, Asc("¬"))) - 2)
    oClan(i).Relation = Right$(ReadField(1 + i, Data, Asc("¬")), 1)
Next

For i = 1 To a
    If oClan(i).Relation = 4 Then
        Call GuildList.AddItem(oClan(i).name)
    End If
Next

For i = 1 To a
    If oClan(i).Relation = 1 Then
        Call GuildList.AddItem(oClan(i).name & " (A)")
    End If
Next

For i = 1 To a
    If oClan(i).Relation = 2 Then
        Call GuildList.AddItem(oClan(i).name & " (E)")
    End If
Next

For i = 1 To a
    If oClan(i).Relation = 0 Then
        Call GuildList.AddItem(oClan(i).name)
    End If
Next

ListaDeClanes = a + 2

End Function
Public Sub ParseMemberInfo(ByVal Data As String)

GuildList.Clear
MembersList.Clear
Text1 = ""

If Me.Visible Then Exit Sub

Dim a As Integer
Dim b As Integer
Dim i As Integer

b = ListaDeClanes(Data)

a = Val(ReadField(b, Data, Asc("¬")))

For i = 1 To a
    Call MembersList.AddItem(ReadField(b + i, Data, Asc("¬")))
Next

b = b + a + 1

Text1 = Replace(ReadField(b, Data, Asc("¬")), "º", vbCrLf)

Call Me.Show(vbModeless, frmPrincipal)
Call Me.SetFocus

End Sub

Private Sub Command1_Click()
Dim GuildName As String


GuildName = GuildList.List(GuildList.ListIndex)
If Right$(GuildName, 1) = ")" Then GuildName = Left$(GuildName, Len(GuildName) - 4)

Call SendData("CLANDETAILS" & GuildName)

End Sub

Private Sub Command4_Click()

frmCharInfo.frmmiembros = 2
Call SendData("1HRINFO<" & MembersList.List(MembersList.ListIndex))

End Sub
Private Sub Command8_Click()

Me.Visible = False
frmPrincipal.SetFocus

End Sub
Private Sub Form_Load()

Me.Picture = LoadPicture(DirGraficos & "GuildMember.gif")

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

If bmoving = False And Button = vbLeftButton Then
    DX = x
    dy = Y
    bmoving = True
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

If bmoving And ((x <> DX) Or (Y <> dy)) Then Move Left + (x - DX), Top + (Y - dy)

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

If Button = vbLeftButton Then bmoving = False

End Sub

