VERSION 5.00
Begin VB.Form frmParty2 
   BorderStyle     =   0  'None
   Caption         =   "Party"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Party2.frx":0000
   ScaleHeight     =   1920
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Rechazar 
      Height          =   375
      Left            =   1920
      MouseIcon       =   "Party2.frx":516F
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Image Aceptar 
      Height          =   375
      Left            =   720
      MouseIcon       =   "Party2.frx":5479
      MousePointer    =   99  'Custom
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Juancito"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmParty2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'F�nixAO 1.0
'
'Based on Argentum Online 0.99z
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'You can contact me at:
'elpresi@fenixao.com.ar
'www.fenixao.com.ar

Private Sub Acepta_Click(Index As Integer)

End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "invitacionparty.gif")
End Sub
Private Sub Rechazar_Click()
Call SendData("PARREC")
frmParty2.Visible = False
End Sub
Private Sub Aceptar_Click()
Call SendData("PARACE")
frmParty2.Visible = False
End Sub
