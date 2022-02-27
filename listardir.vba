xSub BotC3nl Haga clic en ()
Dim ruta, archivos As String
With Application.FileDialog (msoFileDialogFolderPicker)
.Title = "Seleccionar carpeta"
.Show
ruta = .SelectedItems (1)
End With
Dimi As Integer
archivos = dir (ruta & "\*.CIF*")
Sheets ("CARGA").Select
Hoja2. Range ("A1:A" & Hoja2.Cells (Rows.Count, 1) .End (x1Up). Row)
i = 1
Do While Len (archivos) > 0
Hoja2.Cells(i, 1) = archivos
archivos = dir()
1 = 1 + 1
Loop
End Sub



