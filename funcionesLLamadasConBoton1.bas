Attribute VB_Name = "funcionesLLamadasConBoton1"
Sub copiaryPegar()

Sheets("Hoja1").Select 've y activa la hoja 1
Range("B6").Activate 'posicionate en la celda b6 y activala
Range(ActiveCell, ActiveCell.Offset(0, 9)).Copy 'copia la celda b6 y 8 mas hacia la derecha

'copiado = Worksheets("Hoja1").Range(ActiveCell, ActiveCell.Offset(0, 5)).Copy

ActiveCell.Offset(1, 0).Select 'recorrete una fila hacia abajo y cero columnas

Sheets("Enviar").Select 'posicionate en la hoja enviar
Range("A2").Select 'posicionate en la celda a2 y seleccionala
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False 'el _ es para continuar el codigo en otra linea

Call enviarDatos2 'envia los datos a google

Sheets("Hoja1").Select 'selecciona la hoja 1

Do While Not IsEmpty(ActiveCell) 'mientras la celda no este limpia
Range(ActiveCell, ActiveCell.Offset(0, 9)).Copy 'selecciona donde se quedo la celda y dos más a la derecha
Sheets("Enviar").Select 'selecciona hoja enviar
Range("A2").Select 'selecciona celda a2
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False 'pegado especial solo valores

Call enviarDatos2 'manda a google

Sheets("Hoja1").Select ' ve a la hoja uno
ActiveCell.Offset(1, 0).Select 'recorrete una fila hacia abajo y cero columnas
Loop 'termina el ciclo

Call mensajeConfirmacion 'llama la funcion

End Sub
' Toda esta funcion sirve para enviar los datos a google drive
Sub enviarDatos2()
Dim Resultado As String
Dim Url As String, DatoMetodoPost As String
Dim winHttpSolicitud As Object
Set winHttpSolicitud = CreateObject("winhttp.WinhttpRequest.5.1")

'Aqui se cambia la URL Response desde el formulario
Url = "https://docs.google.com/forms/u/0/d/e/1FAIpQLSfcU7sBEcFCSmKjOuaAopVHWoFMX76zLtqKUobVkITNuZM9NQ/formResponse"

'Aqui empieza el metodo post, se cambian todos los entry
DatoMetodoPost = _
  "entry.1373528307=" & Cells(2, 1).Value _
& "&entry.1363908145=" & Cells(2, 2).Value _
& "&entry.97144210=" & Cells(2, 3).Value _
& "&entry.1191197248=" & Cells(2, 4).Value _
& "&entry.313875435=" & Cells(2, 5).Value _
& "&entry.1082578021=" & Cells(2, 6).Value _
& "&entry.598515716=" & Cells(2, 7).Value _
& "&entry.1794081162=" & Cells(2, 8).Value _
& "&entry.614410034=" & Cells(2, 9).Value _
& "&entry.310876649=" & Cells(2, 10).Value
winHttpSolicitud.Open "post", Url, False
winHttpSolicitud.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
winHttpSolicitud.send (DatoMetodoPost)

Resultado = winHttpSolicitud.responseText

End Sub

Sub mensajeConfirmacion()
Application.Speech.Speak ("Tus datos han sido enviados con éxito") 'mensaje de voz
MsgBox ("Tus datos han sido enviados con éxito") ' mensaje de confirmación
End Sub






