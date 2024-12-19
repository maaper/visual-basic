Sub GestionarHorasAverias()
    Dim Ranges As Variant
    Dim RangeNames As Variant
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim TotalHoras() As Double
    Dim EmptyCells As Boolean
    Dim InvalidCells As Boolean
    Dim InvalidCellsList As Collection
    Dim i As Integer
    Dim FileName As String
    Dim FilePath As String
    Dim ReportFile As String
    Dim FSO As Object
    Dim FileOut As Object
    Dim EmailBody As String
    Dim Resumen As String
    Dim EquiposConAverias As Boolean

    ' Definir rangos de las máquinas
    Ranges = Array("F397:F403", "F408:F411", "F416:F418", "F423:F424", "F431:F433")

    RangeNames = Array("Carretillas Gasoil", "Carretillas Eléctricas", "Trackers", "Plataformas", "Reach Stacker")

    ' Dimensionar dinámicamente el array de horas totales
    ReDim TotalHoras(1 To UBound(Ranges) + 1)

    ' Asegurarse de trabajar con la hoja correcta
    If ThisWorkbook.Sheets.Count = 0 Then
        MsgBox "El libro no contiene hojas, verifica el archivo.", vbCritical
        Exit Sub
    End If

    On Error GoTo HandleError
    Set ws = ThisWorkbook.Sheets("Libro")

    If ws Is Nothing Then
        MsgBox "La hoja llamada 'Libro' no se encuentra en este libro. Verifica el nombre de la hoja.", vbCritical
        Exit Sub
    End If

    EmptyCells = True
    InvalidCells = False
    EquiposConAverias = False
    Set InvalidCellsList = New Collection

    ' Validar celdas y calcular sumas
    For i = LBound(Ranges) To UBound(Ranges)
        Set rng = ws.Range(Ranges(i))
        For Each cell In rng
            ' Procesar celdas
            If IsEmpty(cell.Value) Then
                EmptyCells = EmptyCells And True
            Else
                EmptyCells = False
                If Not IsNumeric(cell.Value) Then
                    InvalidCells = True
                    InvalidCellsList.Add cell.Address
                Else
                    TotalHoras(i + 1) = TotalHoras(i + 1) + cell.Value
                    EquiposConAverias = True
                End If
            End If
        Next cell
    Next i

    ' Mostrar errores si hay celdas inválidas
    If InvalidCells Then
        Dim ErrorMsg As String
        ErrorMsg = "Por favor, asegúrate de que las siguientes celdas contengan valores numéricos (enteros o decimales):" & vbCrLf
        Dim itm
        For Each itm In InvalidCellsList
            ErrorMsg = ErrorMsg & itm & vbCrLf
        Next itm
        MsgBox ErrorMsg, vbCritical
        ws.Range(InvalidCellsList(1)).Select
        Exit Sub
    End If

    ' Si todas las celdas están vacías
    If EmptyCells Then
        Resumen = "No se han reportado averías de equipos, todos los equipos estaban disponibles."
        If MsgBox(Resumen & vbCrLf & "¿Es correcto?", vbQuestion + vbYesNo) = vbNo Then
            ws.Range("F397").Select
            Exit Sub
        End If
    Else
        ' Generar resumen si hay averías
        Resumen = "Se han reportado indisponibilidades en los equipos:" & vbCrLf & vbCrLf
        For i = LBound(TotalHoras) To UBound(TotalHoras)
            If TotalHoras(i) > 0 Then
                Resumen = Resumen & RangeNames(i - 1) & ": " & TotalHoras(i) & " horas" & vbCrLf
            End If
        Next i
        If MsgBox(Resumen & vbCrLf & "¿Es correcto?", vbQuestion + vbYesNo) = vbNo Then
            ws.Range("F397").Select
            Exit Sub
        End If
    End If

    ' Crear archivo CSV
    FileName = Format(Now, "hhmmYYYYMMDD") & ".csv"
    FilePath = Environ("TEMP") & "\" & FileName

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set FileOut = FSO.CreateTextFile(FilePath, True)

    FileOut.WriteLine "Máquina,Equipo,Horas_Esperadas,Horas_Averias"

    Dim EquipoCounter As Integer
    EquipoCounter = 1

    For i = LBound(Ranges) To UBound(Ranges)
        Set rng = ws.Range(Ranges(i))
        For Each cell In rng
            FileOut.WriteLine RangeNames(i) & "," & EquipoCounter & ",8," & IIf(IsEmpty(cell.Value), 0, cell.Value)
            EquipoCounter = EquipoCounter + 1
        Next cell
        EquipoCounter = 1
    Next i

    FileOut.Close

    ' Enviar correo
    Dim OutlookApp As Object
    Dim MailItem As Object

    Set OutlookApp = CreateObject("Outlook.Application")
    Set MailItem = OutlookApp.CreateItem(0)

    With MailItem
        .To = "miguel.antunez@es.indorama.net"
        .Subject = "Reporte de horas de indisponibilidad - " & Format(Now, "hh:mm dd/MM/yyyy")
        .Body = "Estimado Miguel, adjunto el reporte de horas de indisponibilidad." & vbCrLf & vbCrLf & Resumen
        .Attachments.Add FilePath
        .Send
    End With

    ' Borrar archivo temporal
    Kill FilePath

    Exit Sub

HandleError:
    MsgBox "Se produjo un error inesperado: " & Err.Description, vbCritical
    Exit Sub

End Sub
