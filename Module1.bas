Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function GetTickCount& Lib "kernel32" () '-> Api GetTickCount
Public X As Single, Y As Single
Public I As Single, r As Single, j As Single
Public Dospi, conT As Single
Public Retraso As Integer

Public img As Integer
Public img1 As Integer
'Realiza un retardo
Public Sub Retardo(ByRef P_Retraso As Long)
      
    'GetTickCount devuelve un valor _
    inicial, y se lo sumamos al de retraso
    
    P_Retraso = P_Retraso + GetTickCount&
  
   While P_Retraso >= GetTickCount&
        DoEvents
    Wend
  
    'Fin
    'MsgBox "Terminó el retardo", vbInformation
  
End Sub

Public Function Dibujar_linea_horizontal_positivo(X As Single, Y As Single, Arana_Img As Object, AranaDos_img As Object) 'X Positivo   '2)))))

Dim I As Single

For I = 0 To 4 Step 0.005 '

Dim conT As Single
conT = conT + 0.005

If conT >= 0.2 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

X = I
Y = 0
'
'Arana_Img.Left = X
'Arana_Img.Top = Y
AranaDos_img.Move X, Y

Form1.PSet (X, Y), vbWhite

Next I

'Arana vuelve al centro
For I = -4 To 0 Step 0.005 '

conT = conT + 0.005

If conT >= 0.2 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

X = -(I)
Y = 0
Y = -(Y)

AranaDos_img.Move X, Y
'Arana_Img.Left = X
'Arana_Img.Top = Y


Next I

End Function
'
Public Function Dibujar_linea2(X As Single, Y As Single, Arana_Img As Object, AranaDos_img As Object) '/   '5))))))))))))

Dim I As Single

For I = 0 To 4 Step 0.005

Dim conT As Single
conT = conT + 0.005

If conT >= 0.2 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

X = I
Y = X

AranaDos_img.Move X, Y
'Arana_Img.Left = X
'Arana_Img.Top = Y

Form1.PSet (X, Y)

Next I

'Vuelve la arana
For I = -4 To 0 Step 0.005

conT = conT + 0.005

If conT >= 0.2 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

X = -(I)
Y = X

AranaDos_img.Move X, Y
'Arana_Img.Left = X
'Arana_Img.Top = Y

Next I



End Function
'
Public Function Dibujar_linea_vertical_positivo(X As Single, Y As Single, Arana_Img As Object, AranaDos_img As Object) ' Positivo

Dim I As Single


For I = 0 To 4 Step 0.005

Dim conT As Single
conT = conT + 0.005

If conT >= 0.2 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

X = 0
Y = I

AranaDos_img.Move X, Y
'Arana_Img.Left = X
'Arana_Img.Top = Y

Form1.PSet (X, Y)

Next I

'Arana vuelve
For I = -4 To 0 Step 0.005

conT = conT + 0.005

If conT >= 0.2 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

X = 0
Y = I
Y = -(I)

AranaDos_img.Move X, Y
'Arana_Img.Left = X
'Arana_Img.Top = Y


Next I

End Function

Public Function Dibujar_linea4(X As Single, Y As Single, Arana_Img As Object, AranaDos_img As Object) '\ X negativo, Y positivo

Dim I As Single

For I = 0 To 4 Step 0.005 '\

Dim conT As Single
conT = conT + 0.005

If conT >= 0.2 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

X = -(I)
Y = -(X)
'
'Arana_Img.Left = X
'Arana_Img.Top = Y
AranaDos_img.Move X, Y

Form1.PSet (X, Y)

Next I

'Arana vuelve
For I = -4 To 0 Step 0.005 '\

conT = conT + 0.005

If conT >= 0.2 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

X = I
Y = X
Y = -(Y)
'
AranaDos_img.Move X, Y
'Arana_Img.Left = X
'Arana_Img.Top = Y

Next I

End Function

Public Function Dibujar_linea_horizontal_negativo(X As Single, Y As Single, Arana_Img As Object, AranaDos_img As Object)   ''''''''1)))))))

Dim I As Single

For I = -4 To 0 Step 0.005 '

Dim conT As Single
conT = conT + 0.005

If conT >= 0.2 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

X = I
Y = 0

AranaDos_img.Move X, Y
'Arana_Img.Left = X
'Arana_Img.Top = Y

Form1.PSet (X, Y)


Next I

End Function

Public Function Dibujar_linea6(X As Single, Y As Single, Arana_Img As Object, AranaDos_img As Object) '/ negativo

Dim I As Single


For I = 0 To 4 Step 0.001

Dim conT As Single
conT = conT + 0.005

If conT >= 0.2 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

X = -(I)
Y = X

AranaDos_img.Move X, Y
'Arana_Img.Left = X
'Arana_Img.Top = Y

Form1.PSet (X, Y)

Next I


'Vuelve la arana
For I = -4 To 0 Step 0.001

conT = conT + 0.005

If conT >= 0.2 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

X = I
Y = X

AranaDos_img.Move X, Y
'Arana_Img.Left = X
'Arana_Img.Top = Y

Next I


End Function

Public Function Dibujar_linea_vertical_negativo(X As Single, Y As Single, Arana_Img As Object, AranaDos_img As Object) ' negativo    '4)))))))

Dim I As Single

For I = 0 To 4 Step 0.001

Dim conT As Single
conT = conT + 0.005

If conT >= 0.2 Then
 'Sleep 1000
 Retraso = 50
 Retardo (Retraso)
 conT = 0
End If

X = 0
Y = -(I)

AranaDos_img.Move X, Y
'Arana_Img.Left = X
'Arana_Img.Top = Y

Form1.PSet (X, Y)

Next I

'arana vuelve
For I = -4 To 0 Step 0.001

conT = conT + 0.005

If conT >= 0.2 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

X = 0
Y = I

AranaDos_img.Move X, Y
'Arana_Img.Left = X
'Arana_Img.Top = Y

Next I

End Function
'
Public Function Dibujar_linea8(X As Single, Y As Single, Arana_Img As Object, AranaDos_img As Object)  '\ X positivo, Y negativo

Dim I As Single

For I = 0 To 4 Step 0.001 '\

Dim conT As Single
conT = conT + 0.005

If conT >= 0.2 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

X = I
Y = X
Y = -(Y)

AranaDos_img.Move X, Y
'Arana_Img.Left = X
'Arana_Img.Top = Y

Form1.PSet (X, Y)

Next I

'Arana vuelve
For I = -4 To 0 Step 0.001 '\

conT = conT + 0.005

If conT >= 0.2 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

X = -(I)
Y = -(X)

AranaDos_img.Move X, Y
'Arana_Img.Left = X
'Arana_Img.Top = Y

Next I

End Function
'
Public Function F1(ByVal theta As Single) As Single
 
 F1 = theta / 80 'Modifica el radio de la circunferencia
 'F2 = Exp(theta / 8)

End Function
'
Public Function Graficar_circunferencia(X As Single, Y As Single, Arana_Img As Object, AranaDos_img As Object)

Dospi = 45 ' aumenta el tamano del espiral

For I = 0 To Dospi Step 0.001

Dim conT As Single
conT = conT + 0.005

If conT >= 1 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

 r = F1(I)
 X = r * Cos(I)
 Y = r * Sin(I)

'se mueve la araña
AranaDos_img.Move X, Y
'Arana_Img.Left = X
'Arana_Img.Top = Y

'Grafíca el espiral
Form1.PSet (X, Y), vbBlack
'Form1.PSet (Arana_Img.Left, Arana_Img.Top), vbBlack
 
Next I

End Function

Public Function F2(ByVal theta As Single) As Single
 
 F2 = Exp(theta / 40) 'Modifica el radio de la circunferencia
' F2 = (0.1) + (0.1) * theta 'donde a= 0 se mueve a travez del eje y  y b= -2,2 es el radio

End Function
'Grafica la seguanda circunferencia
Public Function Graficar_circunferencia2(X As Single, Y As Single, Arana_Img As Object, AranaDos_img As Object)

Dospi = 60 ' aumenta el tamano del espiral

For I = 0 To Dospi Step 0.005

Dim conT As Single
conT = conT + 0.005

If conT >= 1 Then
 'Sleep 1000
 Retraso = 100
 Retardo (Retraso)
 conT = 0
End If

 r = F2(I)
 X = r * Cos(I)
 Y = r * Sin(I)

'se mueve la araña
AranaDos_img.Move X, Y
'Arana_Img.Left = X
'Arana_Img.Top = Y

'Grafíca el espiral
Form1.PSet (X, Y), vbBlack
'Form1.PSet (Arana_Img.Left, Arana_Img.Top), vbBlack
 
Next I

End Function

'Esta función simula movimietno de las patas de la araña
Public Function Mover_Patas_Aranas(ImageList1, imagen1 As Object) As Boolean

If img1 = 1 Then
    img1 = 2
Else

If img1 = 2 Then
    img1 = 1
End If
End If

imagen1.Picture = ImageList1.ListImages(img1).Picture
End Function

