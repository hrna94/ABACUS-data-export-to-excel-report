Option Explicit

' Main setup function - call this to create the UI
Sub SetupSummaryUI()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Zusammenfassung")

    ' Clear any existing controls
    ClearExistingControls ws

    ' Create UI elements
    CreateUILabels ws
    CreateDropdowns ws
    CreateInputFields ws
    CreateButtons ws

    ' Format the UI area
    FormatUIArea ws

    ' UI setup completed - no message box during automated export
End Sub

' Clear existing controls
Private Sub ClearExistingControls(ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If InStr(shp.Name, "UI_") > 0 Then
            shp.Delete
        End If
    Next shp
End Sub

' Create UI labels
Private Sub CreateUILabels(ws As Worksheet)
    ' Main header
    ws.Range("E1").Value = "PARKING ANALYSIS TOOL"
    ws.Range("E1").Font.Bold = True
    ws.Range("E1").Font.Size = 14

    ' Time period filter
    ws.Range("E3").Value = "Zeitraum Filter:"
    ws.Range("E3").Font.Bold = True

    ws.Range("E4").Value = "Von Datum/Zeit:"
    ws.Range("E5").Value = "Bis Datum/Zeit:"

    ' Interval settings
    ws.Range("E7").Value = "Analyse Intervall:"
    ws.Range("E7").Font.Bold = True

    ws.Range("E8").Value = "Intervall Typ:"

    ' Device filters
    ws.Range("E10").Value = "Gerät Filter:"
    ws.Range("E10").Font.Bold = True

    ws.Range("E11").Value = "Einfahrt Gerät:"
    ws.Range("E12").Value = "Ausfahrt Gerät:"

    ' Results area
    ws.Range("E14").Value = "Analyse Ergebnisse:"
    ws.Range("E14").Font.Bold = True
End Sub

' Create dropdown lists
Private Sub CreateDropdowns(ws As Worksheet)
    ' Interval Type dropdown (F8)
    With ws.Range("F8").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="Stündlich,Täglich,Wöchentlich,Monatlich"
    End With
    ws.Range("F8").Value = "Täglich"

    ' Get unique devices from data
    Dim entryDevices As String
    Dim exitDevices As String
    entryDevices = GetUniqueDevicesString("entry")
    exitDevices = GetUniqueDevicesString("exit")

    ' Entry Device dropdown (F11)
    With ws.Range("F11").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:=entryDevices
    End With
    ws.Range("F11").Value = "[Alle Einfahrt Geräte]"

    ' Exit Device dropdown (F12)
    With ws.Range("F12").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:=exitDevices
    End With
    ws.Range("F12").Value = "[Alle Ausfahrt Geräte]"
End Sub

' Get unique devices as comma-separated string
Private Function GetUniqueDevicesString(deviceType As String) As String
    Dim dataWs As Worksheet
    Dim devices As Object
    Dim lastRow As Long
    Dim i As Long
    Dim device As String
    Dim columnIndex As Long
    Dim result As String

    Set dataWs = ThisWorkbook.Worksheets("Parking Report")
    Set devices = CreateObject("Scripting.Dictionary")

    ' Find last row with data
    lastRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row

    ' Determine which column to read based on device type
    ' Column E (5) = Device Name (Entry)
    ' Column I (9) = Device Name (Exit)
    If deviceType = "entry" Then
        columnIndex = 5 ' Column E - Device Name (Entry)
        result = "[Alle Einfahrt Geräte],"
    Else
        columnIndex = 9 ' Column I - Device Name (Exit)
        result = "[Alle Ausfahrt Geräte],"
    End If

    ' Collect unique devices
    For i = 2 To lastRow ' Skip header row
        device = Trim(CStr(dataWs.Cells(i, columnIndex).Value))
        If device <> "" And device <> "N/A" And Not IsEmpty(dataWs.Cells(i, columnIndex).Value) Then
            devices(device) = True
        End If
    Next i

    ' Build comma-separated string
    Dim key As Variant
    For Each key In devices.Keys
        result = result & key & ","
    Next key

    ' Remove trailing comma
    If Right(result, 1) = "," Then
        result = Left(result, Len(result) - 1)
    End If

    GetUniqueDevicesString = result
End Function

' Create input fields with default values
Private Sub CreateInputFields(ws As Worksheet)
    ' From Date/Time (F4) - clear any validation and set default
    With ws.Range("F4").Validation
        .Delete
    End With
    ws.Range("F4").Value = DateSerial(Year(Date), Month(Date), 1)
    ws.Range("F4").NumberFormat = "dd.mm.yyyy hh:mm"

    ' To Date/Time (F5) - clear any validation and set default
    With ws.Range("F5").Validation
        .Delete
    End With
    ws.Range("F5").Value = Now
    ws.Range("F5").NumberFormat = "dd.mm.yyyy hh:mm"
End Sub

' Create action buttons - positioned at X: 700
Private Sub CreateButtons(ws As Worksheet)
    ' Analyze button - positioned at X: 700
    Dim btnAnalyze As Shape
    Set btnAnalyze = ws.Shapes.AddFormControl(xlButtonControl, 700, 65, 100, 25)
    btnAnalyze.Name = "UI_AnalyzeButton"
    btnAnalyze.TextFrame.Characters.Text = "Analyse starten"
    btnAnalyze.OnAction = "RunParkingAnalysis"

    ' Help button - positioned at X: 700
    Dim btnHelp As Shape
    Set btnHelp = ws.Shapes.AddFormControl(xlButtonControl, 700, 95, 80, 25)
    btnHelp.Name = "UI_HelpButton"
    btnHelp.TextFrame.Characters.Text = "Hilfe"
    btnHelp.OnAction = "ShowHelp"

    ' Clear button - positioned at X: 700
    Dim btnClear As Shape
    Set btnClear = ws.Shapes.AddFormControl(xlButtonControl, 700, 125, 80, 25)
    btnClear.Name = "UI_ClearButton"
    btnClear.TextFrame.Characters.Text = "Löschen"
    btnClear.OnAction = "ClearAnalysis"
End Sub

' Format the UI area
Private Sub FormatUIArea(ws As Worksheet)
    ' Format labels
    ws.Range("E3:E14").Font.Bold = True
    ws.Range("E4:E5,E8,E11:E12").Font.Bold = False

    ' Add borders around UI area
    ws.Range("E1:F14").Borders.LineStyle = xlContinuous
    ws.Range("E1:F14").Borders.Weight = xlThin

    ' Auto-fit columns
    ws.Columns("E:F").AutoFit

    ' Set background color for UI area
    ws.Range("E1:F14").Interior.Color = RGB(240, 248, 255)
End Sub


' Main analysis function
Sub RunParkingAnalysis()
    Dim ws As Worksheet
    Dim dataWs As Worksheet

    Set ws = ThisWorkbook.Worksheets("Zusammenfassung")
    Set dataWs = ThisWorkbook.Worksheets("Parking Report")

    ' Clear previous results
    ClearResults

    ' Get parameters from UI
    Dim startDateTime As Date
    Dim endDateTime As Date
    Dim intervalType As String
    Dim entryDevice As String
    Dim exitDevice As String

    ' Get date range from UI
    startDateTime = ws.Range("F4").Value
    endDateTime = ws.Range("F5").Value
    intervalType = ws.Range("F8").Value
    entryDevice = Trim(CStr(ws.Range("F11").Value))
    exitDevice = Trim(CStr(ws.Range("F12").Value))

    ' Handle "All Devices" selection
    If InStr(entryDevice, "[Alle") > 0 Then entryDevice = ""
    If InStr(exitDevice, "[Alle") > 0 Then exitDevice = ""

    ' Validate inputs
    If Not IsDate(startDateTime) Or Not IsDate(endDateTime) Then
        MsgBox "Bitte geben Sie gültige Start- und Enddaten/-zeiten ein", vbCritical
        Exit Sub
    End If

    If startDateTime >= endDateTime Then
        MsgBox "Startdatum muss vor dem Enddatum liegen", vbCritical
        Exit Sub
    End If

    ' Show analysis parameters
    Dim paramMsg As String
    paramMsg = "ANALYSE PARAMETER:" & vbCrLf
    paramMsg = paramMsg & "Zeitraum: " & Format(startDateTime, "dd.mm.yyyy hh:mm") & " - " & Format(endDateTime, "dd.mm.yyyy hh:mm") & vbCrLf
    paramMsg = paramMsg & "Intervall: " & intervalType & vbCrLf
    paramMsg = paramMsg & "Einfahrt Filter: " & IIf(entryDevice = "", "Alle Geräte", entryDevice) & vbCrLf
    paramMsg = paramMsg & "Ausfahrt Filter: " & IIf(exitDevice = "", "Alle Geräte", exitDevice) & vbCrLf & vbCrLf
    paramMsg = paramMsg & "Mit der Analyse fortfahren?"

    If MsgBox(paramMsg, vbYesNo + vbQuestion) = vbNo Then Exit Sub

    ' Analyze data
    Dim results As Collection
    Set results = AnalyzeData(dataWs, startDateTime, endDateTime, intervalType, entryDevice, exitDevice)

    ' Display results
    DisplayResults ws, results, startDateTime, endDateTime, intervalType

    ' Analysis completed - results displayed in worksheet
End Sub

' Main data analysis function
Private Function AnalyzeData(dataWs As Worksheet, startDate As Date, endDate As Date, _
                           intervalType As String, entryFilter As String, exitFilter As String) As Collection

    Dim results As New Collection
    Dim lastRow As Long
    Dim i As Long
    Dim entryTime As Date
    Dim intervalKey As String
    Dim entryDevice As String
    Dim exitDevice As String
    Dim deviceKey As String
    Dim processedCount As Long

    ' Find last row with data
    lastRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row

    ' Create dictionary for aggregation
    Dim analysisDict As Object
    Set analysisDict = CreateObject("Scripting.Dictionary")

    ' Process each row
    processedCount = 0
    For i = 2 To lastRow ' Skip header
        ' Check if we have entry time data in column B (Entry Time)
        Dim entryTimeValue As Variant
        entryTimeValue = dataWs.Cells(i, 2).Value

        If Not IsEmpty(entryTimeValue) And entryTimeValue <> "" Then
            ' Convert Excel serial number to Date
            If IsNumeric(entryTimeValue) Then
                entryTime = CDate(entryTimeValue)
            ElseIf IsDate(entryTimeValue) Then
                entryTime = CDate(entryTimeValue)
            Else
                ' Skip this row if we can't parse the date
                GoTo NextRow
            End If

            ' Check if within date range
            If entryTime >= startDate And entryTime <= endDate Then
                ' Get device information using correct column mappings
                entryDevice = Trim(CStr(dataWs.Cells(i, 5).Value)) ' Column E - Device Name (Entry)
                exitDevice = Trim(CStr(dataWs.Cells(i, 9).Value))  ' Column I - Device Name (Exit)

                ' Handle empty exit device (no exit recorded)
                If exitDevice = "" Or exitDevice = "N/A" Or IsEmpty(dataWs.Cells(i, 9).Value) Then
                    exitDevice = "No Exit Recorded"
                End If

                ' Apply device filters
                Dim entryMatch As Boolean
                Dim exitMatch As Boolean

                entryMatch = (entryFilter = "" Or InStr(LCase(entryDevice), LCase(entryFilter)) > 0)
                exitMatch = (exitFilter = "" Or InStr(LCase(exitDevice), LCase(exitFilter)) > 0)

                If entryMatch And exitMatch Then
                    ' Create interval key
                    intervalKey = GetIntervalKey(entryTime, intervalType)

                    ' Create device combination key
                    deviceKey = entryDevice & " → " & exitDevice

                    ' Create combined key
                    Dim combinedKey As String
                    combinedKey = intervalKey & "|" & deviceKey

                    ' Increment counter
                    If analysisDict.Exists(combinedKey) Then
                        analysisDict(combinedKey) = analysisDict(combinedKey) + 1
                    Else
                        analysisDict(combinedKey) = 1
                    End If

                    processedCount = processedCount + 1
                End If
            End If
        End If
NextRow:
    Next i

    ' Show processing statistics in immediate window
    Debug.Print "Processed " & processedCount & " records out of " & (lastRow - 1) & " total records"

    ' Convert dictionary to collection for easier handling
    Dim key As Variant
    For Each key In analysisDict.Keys
        Dim parts() As String
        parts = Split(key, "|")

        Dim resultItem As Object
        Set resultItem = CreateObject("Scripting.Dictionary")
        resultItem("Interval") = parts(0)
        resultItem("DevicePath") = parts(1)
        resultItem("Count") = analysisDict(key)

        results.Add resultItem
    Next key

    Set AnalyzeData = results
End Function

' Get interval key based on date and interval type
Private Function GetIntervalKey(dateTime As Date, intervalType As String) As String
    Select Case intervalType
        Case "Stündlich"
            GetIntervalKey = Format(dateTime, "dd.mm.yyyy hh") & ":00"
        Case "Täglich"
            GetIntervalKey = Format(dateTime, "dd.mm.yyyy")
        Case "Wöchentlich"
            GetIntervalKey = "Woche " & Format(dateTime, "ww/yyyy")
        Case "Monatlich"
            GetIntervalKey = Format(dateTime, "mm/yyyy")
        Case Else
            GetIntervalKey = Format(dateTime, "dd.mm.yyyy") ' Default to daily
    End Select
End Function

' Display analysis results
Private Sub DisplayResults(ws As Worksheet, results As Collection, startDate As Date, _
                         endDate As Date, intervalType As String)

    Dim startRow As Long
    startRow = 16 ' Moved down to row 16 to avoid overlap with UI

    ' Clear previous results
    ws.Range("E" & startRow & ":J" & ws.Rows.Count).Clear

    ' Add headers
    ws.Range("E" & startRow).Value = "ANALYSE ERGEBNISSE"
    ws.Range("E" & startRow).Font.Bold = True
    ws.Range("E" & startRow).Font.Size = 12

    startRow = startRow + 1

    ws.Range("E" & startRow).Value = "Period: " & Format(startDate, "dd.mm.yyyy hh:mm") & _
                                   " - " & Format(endDate, "dd.mm.yyyy hh:mm")
    ws.Range("E" & startRow).Font.Italic = True

    startRow = startRow + 2

    ' Column headers
    ws.Range("E" & startRow).Value = "Zeit Intervall"
    ws.Range("F" & startRow).Value = "Einfahrt Gerät"
    ws.Range("G" & startRow).Value = "Ausfahrt Gerät"
    ws.Range("H" & startRow).Value = "Anzahl"
    ws.Range("E" & startRow & ":H" & startRow).Font.Bold = True

    startRow = startRow + 1

    ' Data rows
    Dim i As Long
    Dim totalCount As Long
    totalCount = 0

    For i = 1 To results.Count
        Dim item As Object
        Set item = results(i)

        ws.Range("E" & (startRow + i - 1)).Value = item("Interval")

        ' Split device path
        Dim deviceParts() As String
        deviceParts = Split(item("DevicePath"), " → ")
        ws.Range("F" & (startRow + i - 1)).Value = deviceParts(0)
        If UBound(deviceParts) > 0 Then
            ws.Range("G" & (startRow + i - 1)).Value = deviceParts(1)
        End If

        ws.Range("H" & (startRow + i - 1)).Value = item("Count")

        ' Make "No Exit Recorded" bold
        If ws.Range("G" & (startRow + i - 1)).Value = "No Exit Recorded" Then
            ws.Range("G" & (startRow + i - 1)).Font.Bold = True
        End If

        totalCount = totalCount + item("Count")
    Next i

    ' Add total row
    If results.Count > 0 Then
        Dim totalRow As Long
        totalRow = startRow + results.Count
        ws.Range("G" & totalRow).Value = "GESAMT:"
        ws.Range("H" & totalRow).Value = totalCount
        ws.Range("G" & totalRow & ":H" & totalRow).Font.Bold = True
    End If

    ' Format results area
    Dim resultsRange As String
    resultsRange = "E" & (startRow - 1) & ":H" & (startRow + results.Count)
    ws.Range(resultsRange).Borders.LineStyle = xlContinuous
    ws.Columns("E:H").AutoFit

    If results.Count = 0 Then
        ws.Range("E" & startRow).Value = "Keine Daten für die angegebenen Kriterien gefunden."
        ws.Range("E" & startRow).Font.Italic = True
        ws.Range("E" & (startRow + 1)).Value = "Versuchen Sie, den Datumsbereich oder Gerätfilter anzupassen."
        ws.Range("E" & (startRow + 1)).Font.Italic = True
    End If
End Sub

' Clear analysis results (internal function)
Private Sub ClearResults()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Zusammenfassung")

    ' Clear results area (from row 16 down)
    ws.Range("E16:J" & ws.Rows.Count).Clear
End Sub

' Show help information
Sub ShowHelp()
    Dim helpText As String
    helpText = "PARKANALYSE-TOOL HILFE" & vbCrLf & vbCrLf
    helpText = helpText & "DATUMSBEREICH-FILTER:" & vbCrLf
    helpText = helpText & "- Von Datum/Zeit: Beginn des Analysezeitraums" & vbCrLf
    helpText = helpText & "- Bis Datum/Zeit: Ende des Analysezeitraums" & vbCrLf
    helpText = helpText & "- Format: dd.mm.yyyy hh:mm (z.B. 01.08.2025 00:00)" & vbCrLf & vbCrLf
    helpText = helpText & "ANALYSEINTERVALL:" & vbCrLf
    helpText = helpText & "- Wie die Ergebnisse gruppiert werden (stündlich, täglich, wöchentlich, monatlich)" & vbCrLf & vbCrLf
    helpText = helpText & "GERÄTEFILTER:" & vbCrLf
    helpText = helpText & "- Einfahrtsgerät: Spezifisches Einfahrtsgerät oder [Alle Einfahrtsgeräte] auswählen" & vbCrLf
    helpText = helpText & "- Ausfahrtsgerät: Spezifisches Ausfahrtsgerät oder [Alle Ausfahrtsgeräte] auswählen" & vbCrLf & vbCrLf
    helpText = helpText & "ERGEBNISSE:" & vbCrLf
    helpText = helpText & "- Zeigt Verkehrsfluss zwischen Ein- und Ausfahrtsgeräten" & vbCrLf
    helpText = helpText & "- Gruppiert nach ausgewählten Zeitintervallen" & vbCrLf
    helpText = helpText & "- Enthält Gesamtanzahl für den Zeitraum" & vbCrLf & vbCrLf
    helpText = helpText & "BEISPIEL:" & vbCrLf
    helpText = helpText & "- Von: 01.08.2025 00:00, Bis: 31.08.2025 23:59, Intervall: Täglich"

    MsgBox helpText, vbInformation, "Parkanalyse-Tool Hilfe"
End Sub

' Clear analysis results
Sub ClearAnalysis()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Zusammenfassung")

    ' Use the same clearing logic as the analysis function
    ' Clear results area (from row 16 down, columns E to J)
    ws.Range("E16:J" & ws.Rows.Count).Clear

    MsgBox "Analyseergebnisse wurden gelöscht.", vbInformation, "Löschen erfolgreich"
End Sub