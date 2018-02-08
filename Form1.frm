VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   11280
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2850
      Top             =   900
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1950
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   170
      ImageHeight     =   150
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":05BC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdVisualizar 
      Caption         =   "Comenzar"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   8250
      Width           =   1365
   End
   Begin VB.Image Arana_Img 
      Height          =   615
      Left            =   3900
      Top             =   300
      Width           =   765
   End
   Begin VB.Image AranaDos_img 
      Appearance      =   0  'Flat
      Height          =   765
      Left            =   600
      Stretch         =   -1  'True
      Top             =   600
      Width           =   915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function GetTickCount& Lib "kernel32" () '-> Api GetTickCount
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) ' -> funcion sleep

Private Sub Form_Load()
 Form1.DrawWidth = 2
      img = 1
    img1 = 1
End Sub
'Boton visualizar
Private Sub CmdVisualizar_Click()

Timer1.Enabled = True

Form1.Cls
Form1.ScaleMode = 2
Form1.Scale (-4, 4)-(4, -4)

 Dibujar_linea_horizontal_negativo X, Y, Arana_Img, AranaDos_img ' eje X negativo '''1)
'
 Dibujar_linea_horizontal_positivo X, Y, Arana_Img, AranaDos_img ' eje X Positivo '''2)
'
 Dibujar_linea_vertical_positivo X, Y, Arana_Img, AranaDos_img ' eje Y positivo  '3)))))))
'
 Dibujar_linea_vertical_negativo X, Y, Arana_Img, AranaDos_img 'eje Y negativo   '4)))))))
'
 Dibujar_linea2 X, Y, Arana_Img, AranaDos_img '/ X positivo, Y Positivo
'
 Dibujar_linea6 X, Y, Arana_Img, AranaDos_img '/ X negativo, Y negativo
'
 Dibujar_linea4 X, Y, Arana_Img, AranaDos_img '\ X negativo, Y positivo
'
 Dibujar_linea8 X, Y, Arana_Img, AranaDos_img 'X positivo, Y negativo


Graficar_circunferencia X, Y, Arana_Img, AranaDos_img

Graficar_circunferencia2 X, Y, Arana_Img, AranaDos_img

MsgBox "Finalizo la graficación", vbInformation
Timer1.Enabled = False

End Sub
'
Private Sub Timer1_Timer()
    Mover_Patas_Aranas ImageList1, AranaDos_img
End Sub

