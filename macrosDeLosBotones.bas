Attribute VB_Name = "macrosDeLosBotones"
Private Sub CommandButton1_Click()
Application.Speech.Speak ("¿Estás listo para enviar tus datos?") 'mensaje de voz

respuesta = MsgBox("Recuerda que solo puedes enviar este registro una vez en al día. Asegurate de tener buena conexión de internet, ¿Deseas enviar tus datos ahora?", vbYesNo)

If respuesta = vbYes Then
Application.Speech.Speak ("Iniciando el proceso de envio de datos") 'mensaje de voz

CommandButton1.Enabled = False

Call copiaryPegar

'agregar un día
fecha = Worksheets("Hoja2").Range("A1").Value
Worksheets("Hoja2").Range("A1").Value = fecha + 1

Else

End If
End Sub
Private Sub CommandButton2_Click()

    
    ActiveSheet.Unprotect ("armando223")
    
    Application.Speech.Speak ("Espera unos segundos para la actualización de la lista")
    

    
    ActiveWorkbook.RefreshAll
    
    Application.Speech.Speak ("Los datos han sido actualizados")

    
    ActiveSheet.Protect ("armando223")

    
    Call habilitarBoton
    

End Sub

Sub habilitarBoton()

fecha = Worksheets("Hoja2").Range("A1").Value

If fecha = Date Then

CommandButton1.Enabled = True
Else
End If


End Sub

