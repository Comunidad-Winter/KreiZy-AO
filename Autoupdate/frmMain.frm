VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C00000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Autoupdate - KreiZy AO"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   7470
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":08CA
   ScaleHeight     =   2970
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   2040
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6480
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   6000
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Image Command1 
      Height          =   495
      Left            =   240
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lEstado 
      BackStyle       =   0  'Transparent
      Caption         =   "Esperando sincronización."
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   240
      Width           =   6975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño del archivo:"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "A:"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "De:"
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lURL 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2040
      Y1              =   720
      Y2              =   2325
   End
   Begin VB.Label lDirectorio 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label lSize 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   1680
      Width           =   5655
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   7320
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   7095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* UpdateInteligente v4.0 *
'**************************
' Contacto: (dudas o cualquier cosa)
'   MSN/MAIL: shedark@live.com.ar
'   GSZone: www.gs-zone.com.ar, mensaje privado a Shed
' Configuracion:
'   Leer manual adjunto al código
' Nuevo:
'   Código reescrito y simplificado, adaptandolo a las únicas necesidades del programa
'   Posibilidad de elegir que se creen los links automaticamente (EJ: http://host/Parche1.zip) o _
    redirigir hacia un link elegido por ustedes, puede ser cualquiera (pero debe ubicarse en EJ: http://host/Link1.txt) _
    Esto se cambia llendo a Proyecto > Propiedades del proyecto > Generar > BuscarLinks = (0 o 1). Por defecto automático (0).
'   Nueva forma de descarga de archivos más efectiva y que nos permite informar, a medida que se realiza la descarga, _
    el tamaño del archivo descargado, su ubicacion, host y nombre.
'   Nueva forma de escritura y lectura de archivos (destinado unicamente a la búsqueda del Integer del número de actualización)
'   La progressbar nos indica un porcentaje preciso del tamaño del archivo
'   Eliminación de elementos que quedaron en deshuso
' Bugs:
'   En caso de encontrar un error enviar un e-mail o MP (ver Contacto) con:
'       - Una imágen del error (en modo depuración si es posible)
'       - Modificaciones del código (incluyendo links modificados)
'   e intentaré responder cuanto antes
' Los créditos del código del programa corresponden a SHEDARK (Shed)
' AVISO: MANTENTE AL TANTO, NUEVAS VERSIONES MÁS AUTOMÁTICAS

Option Explicit

Rem Programado por Shedark

Dim Directory As String, bDone As Boolean, dError As Boolean, F As Integer
        
Private Sub Command1_Click()
    Command1.Enabled = False
    Call Analizar
End Sub

Private Sub Analizar()
    Dim i As Integer, iX As Integer, tX As Integer, DifX As Integer, dNum As String
    
    lEstado.Caption = "Obteniendo datos..."
    
    iX = Inet1.OpenURL("http://kreizy-ao.ucoz.com/VEREXE.txt") 'Host
    tX = LeerInt(App.Path & "\INIT\Update.ini")
    DifX = iX - tX
    
    If Not (DifX = 0) Then
        For i = 1 To DifX
            Inet1.AccessType = icUseDefault
            dNum = i + tX
            
            #If BuscarLinks Then 'Buscamos el link en el host (1)
                Inet1.URL = Inet1.OpenURL("http://kreizy-ao.ucoz.com/Parche" & dNum & ".txt") 'Host
            #Else                'Generamos Link por defecto (0)
                Inet1.URL = "http://kreizy-ao.ucoz.com/Parche" & dNum & ".zip" 'Host
            #End If
            
            Directory = App.Path & "\INIT\Parche" & dNum & ".zip"
            bDone = False
            dError = False
            
            lURL.Caption = Inet1.URL
            lName.Caption = "Parche" & dNum & ".zip"
            lDirectorio.Caption = App.Path & "\"
                
            frmMain.Inet1.Execute , "GET"
            
            Do While bDone = False
            DoEvents
            Loop
            
            If dError Then Exit Sub
            
            Unzip Directory, App.Path & "\"
            Kill Directory
        Next i
    End If
     
    Call GuardarInt(App.Path & "\INIT\Update.ini", iX)
    
    Command1.Enabled = True
       lEstado.Caption = "Cliente actualizado correctamente."
 

End Sub

Private Sub Command2_Click()
Form1.Show
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    Select Case State
        Case icError
            lEstado.Caption = "Error en la coneccion, descarga abortada."
            bDone = True
            dError = True
        Case icResponseCompleted
            Dim vtData As Variant
            Dim tempArray() As Byte
            Dim FileSize As Long
            
            FileSize = Inet1.GetHeader("Content-length")
            ProgressBar1.Max = FileSize
            
            lEstado.Caption = "Descarga iniciada."
            
            Open Directory For Binary Access Write As #1
                vtData = Inet1.GetChunk(1024, icByteArray)
                DoEvents
                
                
                Do While Not Len(vtData) = 0
                    tempArray = vtData
                    Put #1, , tempArray
                    
                vtData = Inet1.GetChunk(1024, icByteArray)
                    
                    ProgressBar1.Value = ProgressBar1.Value + Len(vtData) * 2
                    lSize.Caption = ProgressBar1.Value & "bytes de " & FileSize & "bytes"

                    DoEvents
                Loop
            Close #1
            
            lEstado.Caption = "Descarga finalizada."
            lSize.Caption = FileSize & "bytes"
            ProgressBar1.Value = 0
            
            bDone = True
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Function LeerInt(ByVal Ruta As String) As Integer
    F = FreeFile
    Open Ruta For Input As F
    LeerInt = Input$(LOF(F), #F)
    Close #F
End Function

Private Sub GuardarInt(ByVal Ruta As String, ByVal data As Integer)
    F = FreeFile
    Open Ruta For Output As F
    Print #F, data
    Close #F
End Sub

