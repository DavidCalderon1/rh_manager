' ================================================================
'  RenderRaceTables  -  Drone Race Dynamic Table Generator
'  Compatible with: Formato_Carrera_de_drones.xlsm
'
'  WHAT THIS MODULE DOES:
'    - Calls RotorHazard API (/api/rhm/race_config_data) via frmCreateHeats
'    - Parses JSON response: groups -> heats -> pilots
'    - Renders heat tables on "Carrera" sheet, starting at COLUMN B --
'      column A is reserved for the general macro buttons (see
'      SetupColumnAButtons below).
'
'  THIS VERSION ADDS:
'    - A small refresh button floating over the top-right corner of every
'      HEAT title, and two buttons (refresh + remix) over every GROUP
'      header. Excel form controls can't take parameters, so each button is
'      a Shape whose *name* encodes the heat/class id ("btnRH_39",
'      "btnRG_5", "btnMX_5") and all buttons of a kind share one dispatcher
'      macro that reads Application.Caller to recover the id.
'    - A hidden sheet "RaceMeta" recording, for every pilot row drawn, which
'      heat/node it belongs to and which cells hold its Fr/round values --
'      this is what lets RefreshHeatCells/RefreshGroupCells update just the
'      right cells in place instead of redrawing the whole sheet.
'    - Columns R1..Rn now show TIME, POINTS, or POSITION depending on the
'      class's win_condition ("Laps_Time__Best_X_Rounds",
'      "Cumulative_Points", "Last_Heat_Position") instead of always laps.
'    - The Fr column now uses the real frequency label the server computed
'      from the active profile (`freq_label` in the JSON), not the old
'      Config!K5/L5 cross-reference.
'
'  HOW TO INSTALL / UPDATE:
'    1. Alt+F11 to open VBA Editor
'    2. Remove the existing "RenderRaceTables" module (don't export first)
'    3. File > Import File... > select this .bas file
'
'  HOW TO CALL:
'    - RenderRaceTablesFromJSON (via frmCreateHeats, unchanged) generates
'      and draws a class's tables for the first time.
'    - RefreshResults (full re-render, kept as a fallback / used after a
'      remix, since pilot identities per row change) still recalls whatever
'      classes were last drawn.
'    - RefreshHeatCells / RefreshGroupCells update just the affected cells.
' ================================================================

Option Explicit

Private Const GAP_COLS      As Integer = 1
Private Const START_ROW     As Integer = 2
Private Const START_COL     As Integer = 2      ' column B -- column A is reserved for buttons
Private Const CLR_GRP_HDR   As Long = 3041331     ' #2E4053 dark slate
Private Const CLR_COL_HDR   As Long = 4890841     ' #4A90D9 blue
Private Const CLR_HEAT_TTL  As Long = 12376046    ' #BCCFEE powder blue
Private Const CLR_ALT_ROW   As Long = 15791611    ' #F0F7FB ice white
Private Const CLR_TOTAL     As Long = 16636630    ' #FDE9D6 peach

Private Const BTN_W As Double = 20
Private Const BTN_H As Double = 14
Private Const BTN_GAP As Double = 2

' NOTE: shape type/text-frame formatting below uses plain numeric literals
' (5 = rounded rectangle, -1/0 = tri-state true/false, 3 = vertical-center
' anchor, 2 = center paragraph alignment) instead of the mso* named
' constants or local Const aliases for them -- kept as bare literals since
' this project has repeatedly hit "variable not defined" compile errors
' that never fully explained themselves, and literals can't be undefined.

' Module-level trackers for the RaceMeta sheets, valid only during a single
' RenderRaceTablesFromData pass (reset at the top of that sub).
Private mMetaWs As Worksheet
Private mMetaGroupWs As Worksheet
Private mMetaRow As Long
Private mMetaGroupRow As Long


' -- PUBLIC: render from a JSON string --------------------------
Public Sub RenderRaceTablesFromJSON(jsonText As String)
    Dim data As Object

    On Error GoTo ErrHandler
    Set data = JsonConverter.ParseJson(jsonText)
    If Not data.Exists("groups") Then
        MsgBox "API response missing 'groups' key.", vbExclamation: Exit Sub
    End If

    RenderRaceTablesFromData data("groups")
    Exit Sub
ErrHandler:
    MsgBox "Render error: " & Err.Description, vbCritical
End Sub

' -- PUBLIC: render from an already-parsed groups Dictionary ----
' Used both by RenderRaceTablesFromJSON (fresh generation, no results yet)
' and by RefreshResults (merged results from one or more classes).
Public Sub RenderRaceTablesFromData(groups As Object)
    Dim ws       As Worksheet
    Dim colOff   As Integer
    Dim i        As Integer
    Dim keys()   As String
    Dim keyCount As Integer
    Dim k        As Variant
    Dim renderedIds As String
    Dim shp As Shape

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Sheets("Carrera")

    ' wipe everything, including buttons from a previous render
    ws.Cells.Clear
    Do While ws.Shapes.count > 0
        ws.Shapes(1).Delete
    Loop
    ws.Cells.Font.name = "Calibri"
    ws.Cells.Font.Size = 10

    Set mMetaWs = GetOrCreateHiddenSheet("RaceMeta")
    mMetaWs.Range("A1:H1").Value = Array("ClassId", "HeatId", "NodeIndex", "Row", "FrCol", "FirstRoundCol", "Rounds", "WinCondition")
    mMetaRow = 2

    Set mMetaGroupWs = GetOrCreateHiddenSheet("RaceMetaGroups")
    mMetaGroupWs.Range("A1:F1").Value = Array("ClassId", "HeaderRow", "StartCol", "EndCol", "StandingsStartRow", "StandingsEndRow")
    mMetaGroupRow = 2

    With ws.Cells(1, 1)
        .Value = "Race Schedule"
        .Font.Bold = True: .Font.Size = 14
        .Font.Color = RGB(44, 62, 80)
    End With

    keyCount = 0
    For Each k In groups.keys
        keyCount = keyCount + 1
        ReDim Preserve keys(1 To keyCount)
        keys(keyCount) = CStr(k)
        renderedIds = renderedIds & CStr(k) & ","
    Next k
    SortAsLong keys, keyCount

    colOff = START_COL
    For i = 1 To keyCount
        Dim grp    As Object
        Dim rounds As Integer
        Set grp = groups(keys(i))
        rounds = CInt(grp("rounds"))
        RenderGroup ws, grp, colOff, START_ROW, rounds
        colOff = colOff + (2 + rounds + 1) + GAP_COLS
    Next i

    SetupColumnAButtons ws

    ' remember which class ids were rendered, so RefreshResults can re-fetch them
    If Len(renderedIds) > 0 Then renderedIds = Left(renderedIds, Len(renderedIds) - 1)
    ThisWorkbook.Sheets("Config").Range("M1").Value = renderedIds

    ws.Activate
    ws.Cells(1, 1).Select
    MsgBox keyCount & " group(s) rendered on the Carrera sheet.", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Render error: " & Err.Description, vbCritical
End Sub

' -- PUBLIC: load every class that exists in RotorHazard right now, not just
' the ones this workbook remembers generating (covers classes created
' directly in the RotorHazard UI, via MultiGP import, etc). --
Public Sub LoadAllFromServer()
    Dim http As Object, url As String
    Dim resp As Object
    Dim g As Object

    If ThisWorkbook.WorkOffline() Then
        MsgBox "Esta opcion trae todas las clases que existen en RotorHazard -- no tiene equivalente sin conexion. Usa 'Actualizar Todo' para recalcular lo que ya esta en esta hoja.", vbInformation
        Exit Sub
    End If

    On Error GoTo ErrConexion
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = ThisWorkbook.Sheets("Config").Range("H2").Value & "/api/rhm/raceclasses/results?t=" & Format(Now, "hhmmss")
    http.Open "GET", url, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.Send

    If http.Status <> 200 And http.Status <> 201 Then
        MsgBox "Error consultando las clases (" & http.Status & ")", vbExclamation
        Exit Sub
    End If

    Set resp = JsonConverter.ParseJson(http.responseText)
    Set g = resp("groups")
    If g.count = 0 Then
        MsgBox "No hay ninguna clase en RotorHazard todavia.", vbInformation
        Exit Sub
    End If

    RenderRaceTablesFromData g
    Exit Sub

ErrConexion:
    MsgBox "No se pudo conectar con RotorHazard.", vbCritical
End Sub

' -- PUBLIC: re-fetch real results for the classes last rendered and redraw --
' Full re-render. Used as a general fallback, and after a remix (pilot
' identities change per row, so a surgical update isn't enough there).
Public Sub RefreshResults()
    Dim idsCsv As String
    Dim mergedGroups As Object

    idsCsv = Trim(CStr(ThisWorkbook.Sheets("Config").Range("M1").Value))
    If idsCsv = "" Then
        MsgBox "No hay ninguna tabla generada todavia para refrescar.", vbExclamation
        Exit Sub
    End If

    If ThisWorkbook.WorkOffline() Then
        ' no server to fetch from -- recompute standings for every class
        ' already on the sheet from whatever is currently typed into R1..Rn,
        ' and redraw each one in place.
        Dim ids() As String, i As Long, classId As Long, classDict As Object
        ids = Split(idsCsv, ",")
        For i = LBound(ids) To UBound(ids)
            If Trim(ids(i)) <> "" Then
                classId = CLng(Trim(ids(i)))
                Set classDict = HarvestClassFromSheet(classId)
                If Not classDict Is Nothing Then
                    Set classDict("standings") = ComputeStandingsLocal(classDict)
                    RenderSingleClassInPlace classDict
                End If
            End If
        Next i
        Exit Sub
    End If

    Set mergedGroups = FetchClassesById(idsCsv, False)

    If mergedGroups.count = 0 Then
        MsgBox "No se pudieron obtener resultados.", vbExclamation
        Exit Sub
    End If

    RenderRaceTablesFromData mergedGroups
End Sub

' Fetch (via API) and merge results for a csv list of class ids into a
' Scripting.Dictionary keyed by class id (string). Used by RefreshResults to
' rebuild "everything that was on the sheet" without duplicating the
' fetch/merge loop. If silent is True,
' a class that fails to load (e.g. deleted since last rendered) is just
' skipped instead of raising a MsgBox -- used when merging in the background
' after "Generar Heats", where a stale id shouldn't interrupt the flow.
Private Function FetchClassesById(idsCsv As String, silent As Boolean) As Object
    Dim http As Object, url As String, baseUrl As String
    Dim ids() As String, i As Long, classId As String
    Dim resp As Object, g As Object, k As Variant
    Dim merged As Object

    Set merged = CreateObject("Scripting.Dictionary")
    idsCsv = Trim(idsCsv)
    If idsCsv = "" Then
        Set FetchClassesById = merged
        Exit Function
    End If

    baseUrl = ThisWorkbook.Sheets("Config").Range("H2").Value
    ids = Split(idsCsv, ",")

    For i = LBound(ids) To UBound(ids)
        classId = Trim(ids(i))
        If classId <> "" Then
            On Error Resume Next
            Set http = Nothing
            Set http = CreateObject("MSXML2.XMLHTTP")
            http.Open "GET", baseUrl & "/api/rhm/raceclass/" & classId & "/results?t=" & Format(Now, "hhmmss"), False
            http.setRequestHeader "Cache-Control", "no-cache"
            http.Send
            On Error GoTo 0

            If Not http Is Nothing Then
                If http.Status = 200 Or http.Status = 201 Then
                    Set resp = JsonConverter.ParseJson(http.responseText)
                    If resp("status") = "ok" Then
                        Set g = resp("groups")
                        For Each k In g.keys
                            Set merged(k) = g(k)
                        Next k
                    End If
                ElseIf Not silent Then
                    MsgBox "Error consultando resultados de la clase " & classId & " (" & http.Status & ")", vbExclamation
                End If
            End If
        End If
    Next i

    Set FetchClassesById = merged
End Function

' -- PUBLIC: render ONE class's block surgically, without touching any other
' class's tables or doing a full sheet clear. Used after "Generar Heats"
' (new class or heats added to an existing one) from frmCreateHeats -- since
' classes sit in independent, non-overlapping column strips on "Carrera",
' redrawing just one doesn't disturb the rest of the sheet. --
Public Sub RenderSingleClassInPlace(grp As Object)
    Dim ws As Worksheet, metaWs As Worksheet, metaGroupWs As Worksheet
    Dim classId As Long
    Dim groupMetaRow As Long
    Dim sCol As Integer, sRow As Integer, totalCol As Integer, rounds As Integer
    Dim lastGroupRow As Long
    Dim firstEverRender As Boolean

    classId = CLng(grp("id"))
    rounds = CInt(grp("rounds"))
    Set ws = ThisWorkbook.Sheets("Carrera")
    Set metaWs = EnsureMetaWs()
    Set metaGroupWs = EnsureMetaGroupWs()
    Set mMetaWs = metaWs
    Set mMetaGroupWs = metaGroupWs

    lastGroupRow = metaGroupWs.Cells(metaGroupWs.Rows.count, 1).End(xlUp).row
    firstEverRender = (lastGroupRow < 2)

    If firstEverRender Then
        ' totally blank sheet (first class ever rendered in this workbook) --
        ' safe to clear everything since there is nothing to lose.
        ws.Cells.Clear
        Do While ws.Shapes.count > 0
            ws.Shapes(1).Delete
        Loop
        ws.Cells.Font.name = "Calibri": ws.Cells.Font.Size = 10
        With ws.Cells(1, 1)
            .Value = "Race Schedule"
            .Font.Bold = True: .Font.Size = 14
            .Font.Color = RGB(44, 62, 80)
        End With
        sCol = START_COL
        sRow = START_ROW
        mMetaRow = 2
        mMetaGroupRow = 2
        RenderGroup ws, grp, sCol, sRow, rounds
        SetupColumnAButtons ws
    Else
        groupMetaRow = FindGroupMetaRow(metaGroupWs, classId)
        If groupMetaRow > 0 Then
            ' this class is already drawn somewhere on the sheet -- wipe just
            ' its column strip (from its header row down to the sheet
            ' bottom) and redraw it fresh in the exact same spot.
            sCol = CInt(metaGroupWs.Cells(groupMetaRow, 3).Value)
            sRow = CInt(metaGroupWs.Cells(groupMetaRow, 2).Value)
            totalCol = CInt(metaGroupWs.Cells(groupMetaRow, 4).Value)

            ws.Range(ws.Cells(sRow, sCol), ws.Cells(ws.Rows.count, totalCol + 1)).Clear
            DeleteShapesInColumnRange ws, sCol, totalCol + 1

            RemoveMetaRowsForClass metaWs, classId
            metaGroupWs.Rows(groupMetaRow).Delete
        Else
            ' brand new class -- append after the rightmost class currently
            ' on the sheet, nothing existing needs to move.
            sCol = NextFreeColOffset(metaGroupWs)
            sRow = START_ROW
        End If

        mMetaRow = metaWs.Cells(metaWs.Rows.count, 1).End(xlUp).row + 1
        mMetaGroupRow = metaGroupWs.Cells(metaGroupWs.Rows.count, 1).End(xlUp).row + 1

        RenderGroup ws, grp, sCol, sRow, rounds
    End If

    ' keep Config!M1 (the class ids "Actualizar Todo" reloads) in sync
    AddClassIdToRenderedList classId
End Sub

' Ensure the RaceMeta hidden sheet exists (with headers), without touching
' its existing rows -- unlike GetOrCreateHiddenSheet, which always wipes it.
Private Function EnsureMetaWs() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("RaceMeta")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = "RaceMeta"
        ws.Range("A1:H1").Value = Array("ClassId", "HeatId", "NodeIndex", "Row", "FrCol", "FirstRoundCol", "Rounds", "WinCondition")
        ws.Visible = xlSheetVeryHidden
    End If
    Set EnsureMetaWs = ws
End Function

' Same as EnsureMetaWs, for RaceMetaGroups.
Private Function EnsureMetaGroupWs() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("RaceMetaGroups")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = "RaceMetaGroups"
        ws.Range("A1:F1").Value = Array("ClassId", "HeaderRow", "StartCol", "EndCol", "StandingsStartRow", "StandingsEndRow")
        ws.Visible = xlSheetVeryHidden
    End If
    Set EnsureMetaGroupWs = ws
End Function

' Remove every RaceMeta row belonging to a class, so RenderGroup can rewrite
' them fresh (bottom-up so deleting doesn't skip rows).
Private Sub RemoveMetaRowsForClass(metaWs As Worksheet, classId As Long)
    Dim lastRow As Long, i As Long
    lastRow = metaWs.Cells(metaWs.Rows.count, 1).End(xlUp).row
    For i = lastRow To 2 Step -1
        If CLng(metaWs.Cells(i, 1).Value) = classId Then
            metaWs.Rows(i).Delete
        End If
    Next i
End Sub

' Delete every Shape whose left edge falls within [colFrom, colTo) -- used to
' drop one class's mini-buttons before redrawing it, without touching
' buttons belonging to any other class (each class occupies its own,
' non-overlapping column strip, so this is a safe, precise selector).
Private Sub DeleteShapesInColumnRange(ws As Worksheet, colFrom As Integer, colTo As Integer)
    Dim shp As Shape
    Dim names() As String
    Dim n As Long, i As Long
    Dim leftBound As Double, rightBound As Double

    leftBound = ws.Cells(1, colFrom).Left
    rightBound = ws.Cells(1, colTo).Left + ws.Cells(1, colTo).Width

    n = 0
    ReDim names(0 To ws.Shapes.count)
    For Each shp In ws.Shapes
        If shp.Left >= leftBound And shp.Left < rightBound Then
            names(n) = shp.Name
            n = n + 1
        End If
    Next shp

    For i = 0 To n - 1
        On Error Resume Next
        ws.Shapes(names(i)).Delete
        On Error GoTo 0
    Next i
End Sub

' Next free starting column for a brand new class, one gap past whatever
' class currently reaches furthest right -- mirrors the increment
' RenderRaceTablesFromData uses across its main loop (colOff + totalCols +
' GAP_COLS), so a class appended here lines up exactly the same way.
Private Function NextFreeColOffset(metaGroupWs As Worksheet) As Integer
    Dim lastRow As Long, i As Long, maxTotalCol As Integer
    maxTotalCol = 0
    lastRow = metaGroupWs.Cells(metaGroupWs.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If CInt(metaGroupWs.Cells(i, 4).Value) > maxTotalCol Then maxTotalCol = CInt(metaGroupWs.Cells(i, 4).Value)
    Next i
    If maxTotalCol = 0 Then
        NextFreeColOffset = START_COL
    Else
        NextFreeColOffset = maxTotalCol + 1 + GAP_COLS
    End If
End Function

' Add a class id to Config!M1 (the csv list "Actualizar Todo" reloads) if
' it isn't already tracked there.
Private Sub AddClassIdToRenderedList(classId As Long)
    Dim idsCsv As String, ids() As String, i As Long
    Dim alreadyThere As Boolean

    idsCsv = Trim(CStr(ThisWorkbook.Sheets("Config").Range("M1").Value))
    alreadyThere = False
    If idsCsv <> "" Then
        ids = Split(idsCsv, ",")
        For i = LBound(ids) To UBound(ids)
            If Trim(ids(i)) = CStr(classId) Then
                alreadyThere = True
                Exit For
            End If
        Next i
    End If

    If Not alreadyThere Then
        If idsCsv = "" Then
            idsCsv = CStr(classId)
        Else
            idsCsv = idsCsv & "," & CStr(classId)
        End If
        ThisWorkbook.Sheets("Config").Range("M1").Value = idsCsv
    End If
End Sub

' -- PUBLIC: surgical refresh of a single heat's cells ----------
Public Sub RefreshHeatCells(heatId As Long)
    Dim http As Object, url As String
    Dim json As Object, heatData As Object
    Dim ws As Worksheet, metaWs As Worksheet
    Dim winCondition As String
    Dim nd As Object
    Dim lastRow As Long, i As Long
    Dim nodeIndex As Long
    Dim found As Boolean

    If ThisWorkbook.WorkOffline() Then
        ' no server truth to reconcile against -- whatever's typed into
        ' R1..Rn already IS the data. Just find this heat's class and
        ' recompute+redraw it (same as the group-level R button offline).
        Dim offlineClassId As Long
        offlineClassId = FindClassIdForHeat(heatId)
        If offlineClassId = 0 Then
            MsgBox "No se encontro la clase de este heat.", vbExclamation
            Exit Sub
        End If
        Dim offlineClassDict As Object
        Set offlineClassDict = HarvestClassFromSheet(offlineClassId)
        If Not offlineClassDict Is Nothing Then
            Set offlineClassDict("standings") = ComputeStandingsLocal(offlineClassDict)
            RenderSingleClassInPlace offlineClassDict
        End If
        Exit Sub
    End If

    Set http = CreateObject("MSXML2.XMLHTTP")
    url = ThisWorkbook.Sheets("Config").Range("H2").Value & "/api/rhm/heat/" & heatId & "/results?t=" & Format(Now, "hhmmss")

    On Error GoTo ErrConexion
    http.Open "GET", url, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.Send

    If http.Status <> 200 And http.Status <> 201 Then
        MsgBox "Error consultando el heat " & heatId & " (" & http.Status & ")", vbExclamation
        Exit Sub
    End If

    Set json = JsonConverter.ParseJson(http.responseText)
    Set heatData = json("heat")
    winCondition = CStr(heatData("win_condition"))

    Set ws = ThisWorkbook.Sheets("Carrera")
    Set metaWs = ThisWorkbook.Sheets("RaceMeta")
    lastRow = metaWs.Cells(metaWs.Rows.count, 1).End(xlUp).row

    ' only occupied slots have a row/RaceMeta entry (empty ones are skipped
    ' at render time so they don't clutter the table). If a slot's occupied
    ' state (per the fresh API data) no longer matches whether it has a row
    ' here, a row needs to appear or disappear -- something a surgical cell
    ' update can't do, so fall back to a full re-render for that case.
    Dim structureChanged As Boolean, nodeHasPilot As Boolean
    structureChanged = False
    For Each nd In heatData("nodes")
        ' a slot pending "Seed Now" resolution in RotorHazard (not enough
        ' configured frequencies for every pilot) has node_index=Null --
        ' skip it here instead of crashing on CLng(Null); once resolved it
        ' gets a real node_index and is picked up as a newly-occupied slot
        ' (structureChanged below) like any other add.
        If HasNodeIndex(nd("node_index")) Then
            nodeIndex = CLng(nd("node_index"))
            nodeHasPilot = HasPilotValue(nd("pilot_id"))
            found = False
            For i = 2 To lastRow
                If CLng(metaWs.Cells(i, 2).Value) = heatId And CLng(metaWs.Cells(i, 3).Value) = nodeIndex Then
                    found = True
                    Exit For
                End If
            Next i

            If nodeHasPilot And found Then
                ApplyNodeResultRow ws, metaWs, i, nd, winCondition
            ElseIf nodeHasPilot <> found Then
                structureChanged = True
            End If
        End If
    Next nd

    If structureChanged Then
        ' a row needs to appear or disappear -- redraw just THIS heat's
        ' class in place (not the whole sheet: every other class stays
        ' exactly as it was). heatData already carries class_id, so no
        ' extra lookup is needed to know which class that is.
        MsgBox "Un piloto se agrego o se quito de un slot de este heat -- eso cambia cuantas filas se ven, asi que se redibuja esta clase.", vbInformation
        RefreshOneClassById CLng(heatData("class_id"))
    End If

    Exit Sub

ErrConexion:
    MsgBox "No se pudo conectar con RotorHazard.", vbCritical
End Sub

' Fetch one class's fresh results and redraw just that class in place --
' shared by the structural-change fallbacks in RefreshHeatCells and
' RefreshGroupCells, and by DeleteHeatAndRedraw.
Private Sub RefreshOneClassById(classId As Long)
    Dim http As Object, url As String
    Dim json As Object, k As Variant, grpObj As Object

    On Error GoTo ErrConexion
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = ThisWorkbook.Sheets("Config").Range("H2").Value & "/api/rhm/raceclass/" & classId & "/results?t=" & Format(Now, "hhmmss")
    http.Open "GET", url, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.Send

    If http.Status <> 200 And http.Status <> 201 Then
        MsgBox "Error consultando la clase " & classId & " (" & http.Status & ")", vbExclamation
        Exit Sub
    End If

    Set json = JsonConverter.ParseJson(http.responseText)
    For Each k In json("groups").keys
        Set grpObj = json("groups")(k)
        RenderSingleClassInPlace grpObj
    Next k
    Exit Sub

ErrConexion:
    MsgBox "No se pudo conectar con RotorHazard.", vbCritical
End Sub

' -- PUBLIC: surgical refresh of every heat's cells in one class, plus its
' "Posiciones" standings block (both come from the same API response, so
' this costs no extra request). --
Public Sub RefreshGroupCells(classId As Long)
    Dim http As Object, url As String
    Dim json As Object, grp As Object, heats As Object, heatObj As Object
    Dim ws As Worksheet, metaWs As Worksheet, metaGroupWs As Worksheet
    Dim winCondition As String
    Dim hk As Variant, nd As Object
    Dim lastRow As Long, i As Long
    Dim heatIdVal As Long, nodeIndex As Long
    Dim groupMetaRow As Long
    Dim sCol As Integer, totalCol As Integer, standingsStartRow As Integer
    Dim standingsEndRow As Long, newEndRow As Long

    If ThisWorkbook.WorkOffline() Then
        ' no server truth to reconcile against -- whatever's typed into
        ' R1..Rn already IS the data. Recompute standings from it and
        ' redraw this class fresh (simpler than a cell-by-cell surgical
        ' update, and offline actions are infrequent enough that a full
        ' redraw of just this one class costs nothing noticeable).
        Dim offlineClassDict As Object
        Set offlineClassDict = HarvestClassFromSheet(classId)
        If offlineClassDict Is Nothing Then
            MsgBox "No se encontro esta clase en la hoja.", vbExclamation
            Exit Sub
        End If
        Set offlineClassDict("standings") = ComputeStandingsLocal(offlineClassDict)
        RenderSingleClassInPlace offlineClassDict
        Exit Sub
    End If

    Set http = CreateObject("MSXML2.XMLHTTP")
    url = ThisWorkbook.Sheets("Config").Range("H2").Value & "/api/rhm/raceclass/" & classId & "/results?t=" & Format(Now, "hhmmss")

    On Error GoTo ErrConexion
    http.Open "GET", url, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.Send

    If http.Status <> 200 And http.Status <> 201 Then
        MsgBox "Error consultando la clase " & classId & " (" & http.Status & ")", vbExclamation
        Exit Sub
    End If

    Set json = JsonConverter.ParseJson(http.responseText)
    If Not json("groups").Exists(CStr(classId)) Then
        MsgBox "Clase no encontrada en la respuesta.", vbExclamation: Exit Sub
    End If
    Set grp = json("groups")(CStr(classId))
    winCondition = CStr(grp("win_condition"))
    Set heats = grp("heats")

    Set ws = ThisWorkbook.Sheets("Carrera")
    Set metaWs = ThisWorkbook.Sheets("RaceMeta")
    lastRow = metaWs.Cells(metaWs.Rows.count, 1).End(xlUp).row

    ' only occupied slots have a row/RaceMeta entry (empty ones are skipped
    ' at render time). If a slot's occupied state (per the fresh API data)
    ' no longer matches whether it has a row here, a row needs to appear or
    ' disappear -- not something a surgical cell update can do, so that
    ' triggers a full re-render below instead.
    Dim structureChanged As Boolean, nodeFound As Boolean, nodeHasPilot As Boolean
    structureChanged = False

    For Each hk In heats.keys
        Set heatObj = heats(hk)
        heatIdVal = CLng(hk)
        For Each nd In heatObj("nodes")
            ' a slot pending "Seed Now" resolution in RotorHazard (not
            ' enough configured frequencies for every pilot) has
            ' node_index=Null -- skip it instead of crashing on CLng(Null);
            ' once resolved it gets a real node_index and is picked up as a
            ' newly-occupied slot (structureChanged below) like any other.
            If HasNodeIndex(nd("node_index")) Then
                nodeIndex = CLng(nd("node_index"))
                nodeHasPilot = HasPilotValue(nd("pilot_id"))
                nodeFound = False
                For i = 2 To lastRow
                    If CLng(metaWs.Cells(i, 2).Value) = heatIdVal And CLng(metaWs.Cells(i, 3).Value) = nodeIndex Then
                        nodeFound = True
                        Exit For
                    End If
                Next i

                If nodeHasPilot And nodeFound Then
                    ApplyNodeResultRow ws, metaWs, i, nd, winCondition
                ElseIf nodeHasPilot <> nodeFound Then
                    structureChanged = True
                End If
            End If
        Next nd
    Next hk

    ' also refresh the "Posiciones" standings block for this class -- reuses
    ' the grp("standings") already fetched above, no extra request needed.
    Set metaGroupWs = ThisWorkbook.Sheets("RaceMetaGroups")
    groupMetaRow = FindGroupMetaRow(metaGroupWs, classId)
    If groupMetaRow > 0 Then
        sCol = CInt(metaGroupWs.Cells(groupMetaRow, 3).Value)
        totalCol = CInt(metaGroupWs.Cells(groupMetaRow, 4).Value)
        standingsStartRow = CInt(metaGroupWs.Cells(groupMetaRow, 5).Value)
        standingsEndRow = CLng(metaGroupWs.Cells(groupMetaRow, 6).Value)

        If standingsStartRow > 0 And grp.Exists("standings") Then
            If Not IsNull(grp("standings")) Then
                If standingsEndRow >= standingsStartRow Then
                    ws.Range(ws.Cells(standingsStartRow, sCol), ws.Cells(standingsEndRow, totalCol)).Clear
                End If
                newEndRow = RenderStandings(ws, grp("standings"), sCol, standingsStartRow, totalCol) - 1
                metaGroupWs.Cells(groupMetaRow, 6).Value = newEndRow
            End If
        End If
    End If

    ' a slot's occupied/empty state no longer matching whether it has a row
    ' here means either a brand-new heat (never drawn), or a pilot added to
    ' /removed from a slot that changes how many rows this heat should show
    ' -- a surgical cell update can't insert/remove a row in the middle of
    ' an existing layout. Redraw just THIS class in place instead (grp
    ' already holds this class's full fresh data, new heat included, so no
    ' extra fetch is needed) -- every other class on the sheet is untouched.
    If structureChanged Then
        MsgBox "Un heat nuevo o un cambio de pilotos afecta cuantas filas se ven en esta clase -- se redibuja esta clase.", vbInformation
        RenderSingleClassInPlace grp
    End If

    Exit Sub

ErrConexion:
    MsgBox "No se pudo conectar con RotorHazard.", vbCritical
End Sub

' Find which class a heat belongs to by scanning RaceMeta, 0 if not found.
Private Function FindClassIdForHeat(heatId As Long) As Long
    Dim metaWs As Worksheet
    Dim lastRow As Long, i As Long
    On Error Resume Next
    Set metaWs = ThisWorkbook.Sheets("RaceMeta")
    On Error GoTo 0
    If metaWs Is Nothing Then
        FindClassIdForHeat = 0
        Exit Function
    End If
    lastRow = metaWs.Cells(metaWs.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If CLng(metaWs.Cells(i, 2).Value) = heatId Then
            FindClassIdForHeat = CLng(metaWs.Cells(i, 1).Value)
            Exit Function
        End If
    Next i
    FindClassIdForHeat = 0
End Function

' Find the RaceMetaGroups row for a given class id, 0 if not found.
Private Function FindGroupMetaRow(metaGroupWs As Worksheet, classId As Long) As Long
    Dim lastRow As Long, i As Long
    lastRow = metaGroupWs.Cells(metaGroupWs.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If CLng(metaGroupWs.Cells(i, 1).Value) = classId Then
            FindGroupMetaRow = i
            Exit Function
        End If
    Next i
    FindGroupMetaRow = 0
End Function

' -- PUBLIC: remix a group's pilots (native "Balanced Random Fill" generator)
' then fully redraw, since pilot identities per row change -- a surgical
' cell update doesn't apply here the way it does for a plain results refresh.
Public Sub RemixGroupAndRedraw(classId As Long)
    Dim http As Object, url As String
    Dim json As Object, groupsArr As Object
    Dim groupCount As Long

    ' fetch this class's heat-groups (G1, G2, ...) so the user can pick which
    ' ones to remix instead of always remixing the whole class -- online via
    ' the API, offline by parsing "G<n>" out of the heat names already on
    ' the sheet (same Collection-of-Dictionary shape either way).
    If ThisWorkbook.WorkOffline() Then
        Set groupsArr = GetLocalClassGroups(classId)
    Else
        On Error GoTo ErrConexion
        Set http = CreateObject("MSXML2.XMLHTTP")
        url = ThisWorkbook.Sheets("Config").Range("H2").Value & "/api/rhm/raceclass/" & classId & "/groups?t=" & Format(Now, "hhmmss")
        http.Open "GET", url, False
        http.setRequestHeader "Cache-Control", "no-cache"
        http.Send

        If http.Status <> 200 And http.Status <> 201 Then
            MsgBox "Error consultando los grupos de la clase " & classId & " (" & http.Status & ")", vbExclamation
            Exit Sub
        End If

        Set json = JsonConverter.ParseJson(http.responseText)
        Set groupsArr = json("groups")
    End If
    groupCount = groupsArr.count

    If groupCount = 0 Then
        MsgBox "Esta clase no tiene grupos para remixar.", vbExclamation
        Exit Sub
    End If

    If groupCount = 1 Then
        ' only one group in this class -- nothing to pick, remix it directly
        PerformRemix classId, CStr(groupsArr(1)("group_id"))
    Else
        ' more than one group -- let the user pick which ones via a real
        ' form (frmRemixGroups) instead of typing numbers into an InputBox
        Dim frm As frmRemixGroups
        Set frm = New frmRemixGroups
        frm.LoadClass classId
        frm.Show
    End If
    Exit Sub

ErrConexion:
    MsgBox "No se pudo establecer conexion con RotorHazard.", vbCritical
End Sub

' -- PUBLIC: confirm, POST the remix for a specific csv of (0-based) group
' ids, then redraw just this class in place. Shared by the single-group
' fast path above and frmRemixGroups' "Remixar" button. Returns True only
' if the remix actually went through (not cancelled, no server error) --
' callers use this to decide whether it's safe to close their picker UI. --
Public Function PerformRemix(classId As Long, groupIdsCsv As String) As Boolean
    Dim http As Object, url As String
    Dim respJson As Object, k As Variant, grpObj As Object

    PerformRemix = False

    If MsgBox("Esto reordena los pilotos del/de los grupo(s) seleccionado(s) usando 'Balanced Random Fill (Minimize Repeats)'. Los tiempos ya corridos en esos heats se pierden (los demas grupos de la clase no se tocan). Continuar?", _
              vbYesNo + vbExclamation, "Confirmar Remix") <> vbYes Then Exit Function

    If ThisWorkbook.WorkOffline() Then
        PerformRemix = PerformRemixLocal(classId, groupIdsCsv)
        Exit Function
    End If

    On Error GoTo ErrConexion
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = ThisWorkbook.Sheets("Config").Range("H2").Value & "/api/rhm/raceclass/remix/" & classId
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send "{""group_ids"":[" & groupIdsCsv & "]}"

    If http.Status <> 200 And http.Status <> 201 Then
        MsgBox "Error del servidor (" & http.Status & "): " & http.responseText, vbCritical
        Exit Function
    End If

    ' surgical: redraw only this class in place, leave every other class's
    ' table exactly as it was (remix never changes heat/row COUNT, so this
    ' is safe -- unlike add/remove of a heat, which needs structureChanged
    ' detection instead)
    Set respJson = JsonConverter.ParseJson(http.responseText)
    For Each k In respJson("groups").keys
        Set grpObj = respJson("groups")(k)
        RenderSingleClassInPlace grpObj
    Next k
    PerformRemix = True
    Exit Function

ErrConexion:
    MsgBox "No se pudo establecer conexion con RotorHazard.", vbCritical
End Function

' -- PUBLIC: delete a group (race class) and its heats/times entirely, then
' redraw the sheet with whatever groups remain. --
Public Sub DeleteGroupAndRedraw(classId As Long)
    Dim http As Object, url As String
    Dim remainingCsv As String

    If MsgBox("Esto elimina PERMANENTEMENTE el grupo " & classId & " y todos sus heats/tiempos. Continuar?", _
              vbYesNo + vbCritical, "Confirmar eliminar grupo") <> vbYes Then Exit Sub

    If ThisWorkbook.WorkOffline() Then
        ' clear just this class's own column strip -- reuses the same
        ' clear/shape-drop/meta-cleanup RenderSingleClassInPlace already
        ' does when redrawing a class "already on the sheet" in place,
        ' just without redrawing anything after (there's nothing left to
        ' draw for a deleted class). Every other class is untouched.
        Dim ws As Worksheet, metaWs As Worksheet, metaGroupWs As Worksheet
        Dim groupMetaRow As Long, sColL As Integer, sRowL As Integer, totalColL As Integer

        Set ws = ThisWorkbook.Sheets("Carrera")
        Set metaWs = ThisWorkbook.Sheets("RaceMeta")
        Set metaGroupWs = ThisWorkbook.Sheets("RaceMetaGroups")

        groupMetaRow = FindGroupMetaRow(metaGroupWs, classId)
        If groupMetaRow = 0 Then
            MsgBox "No se encontro este grupo en la hoja.", vbExclamation
            Exit Sub
        End If

        sColL = CInt(metaGroupWs.Cells(groupMetaRow, 3).Value)
        sRowL = CInt(metaGroupWs.Cells(groupMetaRow, 2).Value)
        totalColL = CInt(metaGroupWs.Cells(groupMetaRow, 4).Value)

        ws.Range(ws.Cells(sRowL, sColL), ws.Cells(ws.Rows.count, totalColL + 1)).Clear
        DeleteShapesInColumnRange ws, sColL, totalColL + 1

        RemoveMetaRowsForClass metaWs, classId
        metaGroupWs.Rows(groupMetaRow).Delete

        RemoveClassIdFromRenderedList classId
        Exit Sub
    End If

    On Error GoTo ErrConexion
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = ThisWorkbook.Sheets("Config").Range("H2").Value & "/api/rhm/raceclass/" & classId
    http.Open "DELETE", url, False
    http.Send

    If http.Status <> 200 And http.Status <> 201 Then
        MsgBox "Error del servidor (" & http.Status & "): " & http.responseText, vbCritical
        Exit Sub
    End If

    remainingCsv = RemoveClassIdFromRenderedList(classId)

    If remainingCsv = "" Then
        ' nothing left rendered -- clear the sheet (still redraws column A buttons)
        RenderRaceTablesFromData CreateObject("Scripting.Dictionary")
    Else
        RefreshResults
    End If
    Exit Sub

ErrConexion:
    MsgBox "No se pudo establecer conexion con RotorHazard.", vbCritical
End Sub

' -- PUBLIC: delete a single heat, then redraw just its class (heat count
' changes, so a surgical cell update can't do this -- reuses the same
' single-class-in-place render used everywhere else structure changes). --
Public Sub DeleteHeatAndRedraw(heatId As Long)
    Dim http As Object, url As String
    Dim classId As Long
    Dim json As Object
    Dim heatName As String

    classId = FindClassIdForHeat(heatId)
    heatName = "#" & heatId

    If ThisWorkbook.WorkOffline() Then
        ' no API to ask for the display_name -- harvest the class once
        ' (pure sheet reads) and pull it from there instead.
        If classId <> 0 Then
            Dim previewDict As Object
            Set previewDict = HarvestClassFromSheet(classId)
            If Not previewDict Is Nothing Then
                If previewDict("heats").Exists(CStr(heatId)) Then
                    heatName = CStr(previewDict("heats")(CStr(heatId))("display_name"))
                End If
            End If
        End If
    Else
        ' fetch the heat's display_name so the confirmation shows something
        ' meaningful instead of just the raw id
        On Error Resume Next
        Set http = CreateObject("MSXML2.XMLHTTP")
        url = ThisWorkbook.Sheets("Config").Range("H2").Value & "/api/rhm/heat/" & heatId & "/results?t=" & Format(Now, "hhmmss")
        http.Open "GET", url, False
        http.setRequestHeader "Cache-Control", "no-cache"
        http.Send
        If http.Status = 200 Or http.Status = 201 Then
            Set json = JsonConverter.ParseJson(http.responseText)
            If Not json Is Nothing Then
                If json.Exists("heat") Then
                    If json("heat").Exists("display_name") Then heatName = CStr(json("heat")("display_name"))
                End If
            End If
        End If
        On Error GoTo 0
    End If

    If MsgBox("Esto elimina PERMANENTEMENTE el heat '" & heatName & "' y sus tiempos. Continuar?", _
              vbYesNo + vbCritical, "Confirmar eliminar heat") <> vbYes Then Exit Sub

    If ThisWorkbook.WorkOffline() Then
        If classId = 0 Then
            MsgBox "No se encontro la clase de este heat.", vbExclamation
            Exit Sub
        End If
        DeleteHeatLocal classId, heatId
        Exit Sub
    End If

    On Error GoTo ErrConexion
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = ThisWorkbook.Sheets("Config").Range("H2").Value & "/api/rhm/heat/" & heatId
    http.Open "DELETE", url, False
    http.Send

    If http.Status <> 200 And http.Status <> 201 Then
        MsgBox "Error del servidor (" & http.Status & "): " & http.responseText, vbCritical
        Exit Sub
    End If

    If classId = 0 Then
        ' didn't find its class in RaceMeta (shouldn't normally happen --
        ' this button only exists on an already-rendered heat) -- fall back
        ' to a full re-render so the sheet still ends up correct.
        RefreshResults
        Exit Sub
    End If

    RefreshOneClassById classId
    Exit Sub

ErrConexion:
    MsgBox "No se pudo establecer conexion con RotorHazard.", vbCritical
End Sub

' Removes one id from the Config!M1 comma list (the classes last rendered)
' and returns the updated list.
Private Function RemoveClassIdFromRenderedList(classId As Long) As String
    Dim idsCsv As String, ids() As String, i As Long, newCsv As String
    idsCsv = Trim(CStr(ThisWorkbook.Sheets("Config").Range("M1").Value))
    If idsCsv <> "" Then
        ids = Split(idsCsv, ",")
        For i = LBound(ids) To UBound(ids)
            If Trim(ids(i)) <> CStr(classId) And Trim(ids(i)) <> "" Then
                newCsv = newCsv & Trim(ids(i)) & ","
            End If
        Next i
        If Len(newCsv) > 0 Then newCsv = Left(newCsv, Len(newCsv) - 1)
    End If
    ThisWorkbook.Sheets("Config").Range("M1").Value = newCsv
    RemoveClassIdFromRenderedList = newCsv
End Function

' -- Button dispatchers (Application.Caller carries the shape's name) --
Public Sub RefreshHeatButton_Click()
    Dim parts() As String
    parts = Split(CStr(Application.Caller), "_")
    If UBound(parts) >= 1 Then RefreshHeatCells CLng(parts(1))
End Sub

Public Sub EditHeatButton_Click()
    Dim parts() As String
    parts = Split(CStr(Application.Caller), "_")
    If UBound(parts) >= 1 Then EditHeatPilotsById CLng(parts(1))
End Sub

Public Sub DeleteHeatButton_Click()
    Dim parts() As String
    parts = Split(CStr(Application.Caller), "_")
    If UBound(parts) >= 1 Then DeleteHeatAndRedraw CLng(parts(1))
End Sub

Public Sub RefreshGroupButton_Click()
    Dim parts() As String
    parts = Split(CStr(Application.Caller), "_")
    If UBound(parts) >= 1 Then RefreshGroupCells CLng(parts(1))
End Sub

Public Sub RemixGroupButton_Click()
    Dim parts() As String
    parts = Split(CStr(Application.Caller), "_")
    If UBound(parts) >= 1 Then RemixGroupAndRedraw CLng(parts(1))
End Sub

Public Sub DeleteGroupButton_Click()
    Dim parts() As String
    parts = Split(CStr(Application.Caller), "_")
    If UBound(parts) >= 1 Then DeleteGroupAndRedraw CLng(parts(1))
End Sub

' -- Column A: general macro launcher buttons (static, no id needed) --
' Column A width (character units) is 24; button width in points is 118 --
' kept as plain literals (see the note above) with a comfortable margin
' between them, since ColumnWidth units don't map 1:1 to the points Shapes
' use for .Left/.Width -- this avoids the button spilling into column B.
Private Sub SetupColumnAButtons(ws As Worksheet)
    ws.Columns(1).ColumnWidth = 24

    AddLabeledButton ws, 2, "Generar Heats", "ThisWorkbook.OpenCreateHeats", "Crear heats nuevos, o agregarlos a una clase que ya existe"
    AddLabeledButton ws, 4, "Actualizar Todo", "RefreshResults", "Recargar todas las clases que ya estan dibujadas en esta hoja"
    AddLabeledButton ws, 6, "Cargar Todo RH", "LoadAllFromServer", "Traer TODAS las clases que existen en RotorHazard, incluso las que esta hoja todavia no dibujo"
End Sub

Private Sub AddLabeledButton(ws As Worksheet, atRow As Integer, caption As String, macroName As String, Optional tooltip As String = "")
    Dim shp As Shape
    Dim c As Range
    Set c = ws.Cells(atRow, 1)

    Set shp = ws.Shapes.AddShape(5, c.Left + 2, c.Top + 2, 118, 26)
    shp.Name = "btnCol_" & atRow
    With shp.TextFrame2
        .WordWrap = -1
        .MarginLeft = 4: .MarginRight = 4: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = 3
        .TextRange.Text = caption
        .TextRange.Font.Size = 8
        .TextRange.Font.Bold = True
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = 2
    End With
    shp.Fill.ForeColor.RGB = RGB(74, 144, 217)
    shp.Line.ForeColor.RGB = RGB(46, 64, 83)
    shp.OnAction = macroName
    If tooltip <> "" Then
        shp.AlternativeText = tooltip
        On Error Resume Next
        shp.ActionSettings(1).ScreenTip = tooltip
        On Error GoTo 0
    End If
End Sub

' -- Small floating button over the top-right of a merged range --
' offsetIndex 0 = rightmost, 1 = next one to its left, etc. tooltip shows on
' mouse hover (via ActionSettings.ScreenTip, set independently of the click
' macro in .OnAction) since these buttons are too small to fit a clear label
' on their own (R/E/MX/X).
Private Sub AddMiniButton(ws As Worksheet, anchorCell As Range, offsetIndex As Integer, _
                           caption As String, macroName As String, shapeNamePrefix As String, uniqueId As String, _
                           Optional fillColor As Long = -1, Optional tooltip As String = "")
    Dim shp As Shape
    Dim shapeName As String

    shapeName = shapeNamePrefix & "_" & uniqueId

    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(5, _
        anchorCell.Left + anchorCell.Width - ((BTN_W + BTN_GAP) * (offsetIndex + 1)), _
        anchorCell.Top + 2, BTN_W, BTN_H)
    shp.Name = shapeName
    With shp.TextFrame2
        .WordWrap = 0
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = 3
        .TextRange.Text = caption
        .TextRange.Font.Size = 7
        .TextRange.Font.Bold = True
        .TextRange.Font.Fill.ForeColor.RGB = RGB(40, 55, 71)
        .TextRange.ParagraphFormat.Alignment = 2
    End With
    If fillColor = -1 Then
        shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Else
        shp.Fill.ForeColor.RGB = fillColor
    End If
    shp.Line.ForeColor.RGB = RGB(120, 120, 120)
    shp.OnAction = macroName
    If tooltip <> "" Then
        shp.AlternativeText = tooltip
        On Error Resume Next
        shp.ActionSettings(1).ScreenTip = tooltip
        On Error GoTo 0
    End If
End Sub

' Render one group block
Private Sub RenderGroup(ws As Worksheet, grp As Object, _
                        sCol As Integer, sRow As Integer, rounds As Integer)
    Dim totalCols As Integer
    Dim totalCol  As Integer
    Dim heats     As Object
    Dim hkeys()   As String
    Dim hcount    As Integer
    Dim curRow    As Integer
    Dim hk        As Variant
    Dim i         As Integer
    Dim winCondition As String
    Dim classId   As String
    Dim groupMetaRow As Long
    Dim standingsStartRow As Long, standingsEndRow As Long

    totalCols = 2 + rounds + 1
    totalCol = sCol + totalCols - 1
    Set heats = grp("heats")
    winCondition = ""
    If grp.Exists("win_condition") Then winCondition = CStr(grp("win_condition"))
    classId = CStr(grp("id"))

    ' Column widths -- MUST be set before drawing/placing anything below: the
    ' mini-button shapes are positioned using absolute point coordinates read
    ' from cell .Left/.Width, so resizing columns *after* placing them would
    ' leave the buttons stranded at their old (pre-resize) position.
    ws.Columns(sCol).ColumnWidth = 15
    ws.Columns(sCol + 1).ColumnWidth = 5
    For i = 1 To rounds
        ws.Columns(sCol + 1 + i).ColumnWidth = IIf(winCondition = "Laps_Time__Best_X_Rounds", 10, 6)
    Next i
    ws.Columns(totalCol).ColumnWidth = 7
    If totalCol + 1 <= 702 Then ws.Columns(totalCol + 1).ColumnWidth = 2

    ' Group header
    With ws.Range(ws.Cells(sRow, sCol), ws.Cells(sRow, totalCol))
        .Merge
        .Value = grp("display_name")
        .Font.Bold = True: .Font.Size = 11
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = CLR_GRP_HDR
        .HorizontalAlignment = xlCenter
        .RowHeight = 24
    End With
    AddMiniButton ws, ws.Cells(sRow, totalCol), 0, "R", "RefreshGroupButton_Click", "btnRG", classId, tooltip:="Recargar este grupo (heats y posiciones)"
    AddMiniButton ws, ws.Cells(sRow, totalCol), 1, "MX", "RemixGroupButton_Click", "btnMX", classId, tooltip:="Remix: elegir que grupo(s) de esta clase reordenar (Balanced Random Fill)"
    AddMiniButton ws, ws.Cells(sRow, totalCol), 2, "X", "DeleteGroupButton_Click", "btnDG", classId, RGB(255, 214, 214), "Eliminar este grupo/clase completo (heats y tiempos)"

    groupMetaRow = mMetaGroupRow
    mMetaGroupWs.Cells(groupMetaRow, 1).Value = classId
    mMetaGroupWs.Cells(groupMetaRow, 2).Value = sRow
    mMetaGroupWs.Cells(groupMetaRow, 3).Value = sCol
    mMetaGroupWs.Cells(groupMetaRow, 4).Value = totalCol
    mMetaGroupRow = mMetaGroupRow + 1

    ' Column headers
    curRow = sRow + 1
    WriteHeaders ws, sCol, curRow, rounds, totalCols, winCondition
    curRow = curRow + 1

    ' Sort heats
    hcount = 0
    For Each hk In heats.keys
        hcount = hcount + 1
        ReDim Preserve hkeys(1 To hcount)
        hkeys(hcount) = CStr(hk)
    Next hk
    SortAsLong hkeys, hcount

    For i = 1 To hcount
        curRow = RenderHeat(ws, heats(hkeys(i)), sCol, curRow, rounds, totalCols, winCondition, classId)
        curRow = curRow + 1
    Next i

    standingsStartRow = 0
    standingsEndRow = 0
    If grp.Exists("standings") Then
        If Not IsNull(grp("standings")) Then
            standingsStartRow = curRow
            curRow = RenderStandings(ws, grp("standings"), sCol, curRow, totalCol)
            standingsEndRow = curRow - 1
        End If
    End If
    mMetaGroupWs.Cells(groupMetaRow, 5).Value = standingsStartRow
    mMetaGroupWs.Cells(groupMetaRow, 6).Value = standingsEndRow
End Sub

' Renders the class's overall standings (position across all rounds so far,
' per the class's own ranking method) below its heats. Returns the next
' free row (unused by the caller today, but kept consistent with RenderHeat).
Private Function RenderStandings(ws As Worksheet, standingsObj As Object, _
                                  sCol As Integer, sRow As Integer, totalCol As Integer) As Integer
    Dim curRow As Integer
    Dim st As Object
    Dim methodLabel As String
    Dim rowClr As Long
    Dim i As Integer
    Dim detail As String
    Dim k As Variant

    curRow = sRow
    methodLabel = ""
    If standingsObj.Exists("method_label") Then
        If Not IsNull(standingsObj("method_label")) Then methodLabel = CStr(standingsObj("method_label"))
    End If

    With ws.Range(ws.Cells(curRow, sCol), ws.Cells(curRow, totalCol))
        .Merge
        .Value = "Posiciones" & IIf(methodLabel <> "", " (" & methodLabel & ")", "")
        .Font.Bold = True: .Font.Italic = True
        .Interior.Color = RGB(230, 230, 230)
        .HorizontalAlignment = xlLeft: .IndentLevel = 1
    End With
    curRow = curRow + 1

    i = 0
    If standingsObj.Exists("standings") Then
        For Each st In standingsObj("standings")
            If i Mod 2 = 0 Then rowClr = RGB(255, 255, 255) Else rowClr = CLR_ALT_ROW

            ' "key = value" pairs joined by " , " (spaces around "=" and ",")
            ' e.g. "heat = Practica G3-H1 , heat_rank = 4"
            detail = ""
            If st.Exists("extra") Then
                For Each k In st("extra").keys
                    If detail <> "" Then detail = detail & " , "
                    detail = detail & k & " = " & CStr(st("extra")(k))
                Next k
            End If

            ' callsign goes in the wide "Pilot"-width column (sCol), position
            ' in the narrow "Fr"-width column (sCol+1) -- a 1-2 digit number
            ' fits fine there, but a pilot name doesn't (this was swapped
            ' before and truncated names).
            With ws.Cells(curRow, sCol)
                .Value = st("callsign")
                .Interior.Color = rowClr
            End With
            With ws.Cells(curRow, sCol + 1)
                .Value = IIf(IsNull(st("position")), "", st("position"))
                .Interior.Color = rowClr: .Font.Bold = True
                .HorizontalAlignment = xlCenter
            End With
            With ws.Range(ws.Cells(curRow, sCol + 2), ws.Cells(curRow, totalCol))
                .Merge
                .Value = detail
                .Interior.Color = rowClr
                .Font.Size = 8
                .HorizontalAlignment = xlLeft
            End With
            i = i + 1
            curRow = curRow + 1
        Next st
    End If

    RenderStandings = curRow
End Function

' Write header row
Private Sub WriteHeaders(ws As Worksheet, sCol As Integer, hRow As Integer, _
                          rounds As Integer, totalCols As Integer, winCondition As String)
    Dim c As Integer
    Dim labels() As String
    Dim resultLabel As String
    ReDim labels(1 To totalCols)
    labels(1) = "Pilot": labels(2) = "Fr"

    Select Case winCondition
        Case "Cumulative_Points": resultLabel = "Pts"
        Case "Last_Heat_Position": resultLabel = "Pos"
        Case "Laps_Time__Best_X_Rounds": resultLabel = "Time"
        Case Else: resultLabel = "R"
    End Select

    Dim r As Integer
    For r = 1 To rounds: labels(2 + r) = resultLabel & r: Next r
    labels(totalCols) = "Total"
    For c = 1 To totalCols
        With ws.Cells(hRow, sCol + c - 1)
            .Value = labels(c)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = CLR_COL_HDR
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With
    Next c
End Sub

' Decide what to show in a round cell based on the class's win_condition.
Private Sub GetRoundCellValue(nd As Object, roundNum As Integer, winCondition As String, _
                               ByRef cellVal As Variant, ByRef numFmt As String, ByRef commentText As String)
    Dim rd As Object
    Dim found As Boolean
    Dim lapVal As Variant, timeVal As String, posVal As Variant, ptsVal As Variant

    found = False
    lapVal = 0: timeVal = "": posVal = Empty: ptsVal = Empty

    If nd.Exists("rounds") Then
        For Each rd In nd("rounds")
            If CInt(rd("round")) = roundNum Then
                lapVal = rd("laps")
                If Not IsNull(rd("time")) Then timeVal = CStr(rd("time"))
                If Not IsNull(rd("position")) Then posVal = rd("position")
                If Not IsNull(rd("points")) Then ptsVal = rd("points")
                found = True
                Exit For
            End If
        Next rd
    End If

    Select Case winCondition
        Case "Cumulative_Points"
            numFmt = "0"
            commentText = timeVal
            If found And Not IsEmpty(ptsVal) Then cellVal = ptsVal Else cellVal = 0

        Case "Last_Heat_Position"
            numFmt = "0"
            commentText = timeVal
            If found And Not IsEmpty(posVal) Then cellVal = posVal Else cellVal = ""

        Case "Laps_Time__Best_X_Rounds"
            numFmt = "@"
            If found Then commentText = lapVal & " vueltas" Else commentText = ""
            If found And timeVal <> "" Then cellVal = timeVal Else cellVal = "-"

        Case Else
            numFmt = "0"
            commentText = timeVal
            cellVal = lapVal
    End Select
End Sub

' Apply a fetched node's data to its already-known row (used by both the
' initial full render and the surgical Refresh*Cells routines).
Private Sub ApplyNodeResultRow(ws As Worksheet, metaWs As Worksheet, metaRowIdx As Long, nd As Object, winCondition As String)
    Dim targetRow As Long, frCol As Long, firstRoundCol As Long, pilotCol As Long, rounds As Integer
    Dim r As Integer
    Dim cellVal As Variant, numFmt As String, commentText As String
    Dim pilotID As Variant, hasPilot As Boolean

    targetRow = CLng(metaWs.Cells(metaRowIdx, 4).Value)
    frCol = CLng(metaWs.Cells(metaRowIdx, 5).Value)
    firstRoundCol = CLng(metaWs.Cells(metaRowIdx, 6).Value)
    rounds = CInt(metaWs.Cells(metaRowIdx, 7).Value)
    pilotCol = frCol - 1

    ' pilot name -- picks up a pilot being added, swapped, or removed on
    ' this exact slot (edited from frmManageHeat or directly in RotorHazard)
    pilotID = Empty
    If nd.Exists("pilot_id") Then pilotID = nd("pilot_id")
    hasPilot = HasPilotValue(pilotID)
    With ws.Cells(targetRow, pilotCol)
        If hasPilot Then
            .Value = GetCallsign(CLng(pilotID))
            .Font.Italic = False
        Else
            .Value = "(vacio)"
            .Font.Italic = True
        End If
    End With

    With ws.Cells(targetRow, frCol)
        .Value = ""
        If nd.Exists("freq_label") Then
            If Not IsNull(nd("freq_label")) Then
                If CStr(nd("freq_label")) <> "" Then .Value = nd("freq_label")
            End If
        End If
    End With

    For r = 1 To rounds
        GetRoundCellValue nd, r, winCondition, cellVal, numFmt, commentText
        With ws.Cells(targetRow, firstRoundCol + r - 1)
            .NumberFormat = numFmt
            .Value = cellVal
            .ClearComments
            If commentText <> "" Then .AddComment(commentText).Visible = False
        End With
    Next r
End Sub

' Render one heat block; returns next free row
Private Function RenderHeat(ws As Worksheet, heatObj As Object, _
                             sCol As Integer, sRow As Integer, _
                             rounds As Integer, totalCols As Integer, _
                             winCondition As String, classId As String) As Integer
    Dim totalCol   As Integer
    Dim curRow     As Integer
    Dim pilotCount As Integer
    Dim nd         As Object
    Dim pilotID    As Variant
    Dim fc         As String
    Dim lc         As String
    Dim heatId     As String
    Dim cellVal As Variant, numFmt As String, commentText As String
    Dim r As Integer

    totalCol = sCol + totalCols - 1
    curRow = sRow
    pilotCount = 0
    heatId = CStr(heatObj("id"))

    ' Heat title
    With ws.Range(ws.Cells(curRow, sCol), ws.Cells(curRow, totalCol))
        .Merge
        .Value = heatObj("display_name")
        .Font.Bold = True
        .Font.Color = RGB(40, 55, 71)
        .Interior.Color = CLR_HEAT_TTL
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .RowHeight = 18
    End With
    AddMiniButton ws, ws.Cells(curRow, totalCol), 0, "R", "RefreshHeatButton_Click", "btnRH", heatId, tooltip:="Recargar este heat"
    AddMiniButton ws, ws.Cells(curRow, totalCol), 1, "E", "EditHeatButton_Click", "btnEH", heatId, tooltip:="Editar los pilotos de este heat"
    AddMiniButton ws, ws.Cells(curRow, totalCol), 2, "X", "DeleteHeatButton_Click", "btnDH", heatId, RGB(255, 214, 214), "Eliminar este heat completo"
    curRow = curRow + 1

    ' Only occupied slots get a row -- an empty slot has no pilot worth
    ' showing, so it's skipped entirely (not drawn, not recorded in
    ' RaceMeta). A slot going from empty to occupied (or vice versa) later
    ' is picked up by RefreshHeatCells/RefreshGroupCells detecting that its
    ' RaceMeta presence no longer matches the API's current pilot_id, and
    ' falling back to a full re-render of that class (a row needs to be
    ' added or removed, which a surgical cell update can't do).
    For Each nd In heatObj("nodes")
        pilotID = nd("pilot_id")
        If Not HasPilotValue(pilotID) Then GoTo SkipEmptyNode

        ' a pilot can be seeded into the heat but still have node_index=Null
        ' -- pending "Seed Now" resolution in RotorHazard (typically when
        ' there are fewer configured frequencies than pilots needing one).
        ' Still show the pilot (so they're not silently missing), but flag
        ' the Fr cell instead of a blank/real frequency, and don't record
        ' this row in RaceMeta below -- Null can't serve as a stable match
        ' key, so a surgical refresh can't safely target it; a full
        ' "Actualizar Todo" redraws it fresh each time regardless.
        Dim hasNodeIdx As Boolean
        hasNodeIdx = HasNodeIndex(nd("node_index"))

        Dim rowClr As Long
        If pilotCount Mod 2 = 0 Then rowClr = RGB(255, 255, 255) Else rowClr = CLR_ALT_ROW

        With ws.Cells(curRow, sCol)
            .Value = GetCallsign(CLng(pilotID))
            .Interior.Color = rowClr
            .HorizontalAlignment = xlLeft: .IndentLevel = 1
        End With
        With ws.Cells(curRow, sCol + 1)
            .ClearComments
            If hasNodeIdx Then
                .Value = ""
                If nd.Exists("freq_label") Then
                    If Not IsNull(nd("freq_label")) Then .Value = nd("freq_label")
                End If
            Else
                .Value = "?"
                .AddComment("Sin nodo/frecuencia asignada -- resuelve 'Seed Now' en RotorHazard, luego 'Actualizar Todo'.").Visible = False
            End If
            .Interior.Color = rowClr
            .HorizontalAlignment = xlCenter: .Font.Bold = True
        End With
        For r = 1 To rounds
            GetRoundCellValue nd, r, winCondition, cellVal, numFmt, commentText
            With ws.Cells(curRow, sCol + 1 + r)
                ' NumberFormat MUST be set before Value -- a time-like string
                ' ("0:45.243") assigned to a General-formatted cell gets
                ' auto-converted by Excel into a time serial number before
                ' the "@" text format is applied, so it displays the raw
                ' serial number instead of the original time text.
                .NumberFormat = numFmt
                .Value = cellVal
                .Interior.Color = rowClr
                .HorizontalAlignment = xlRight
                .ClearComments
                If commentText <> "" Then .AddComment(commentText).Visible = False
            End With
        Next r
        fc = ColLtr(sCol + 2) & curRow
        lc = ColLtr(sCol + 1 + rounds) & curRow
        With ws.Cells(curRow, totalCol)
            If winCondition = "Laps_Time__Best_X_Rounds" Then
                .Formula = "" ' text times aren't summable -- leave Total blank for this win_condition
                .Value = ""
            Else
                .Formula = "=SUM(" & fc & ":" & lc & ")"
            End If
            .Interior.Color = CLR_TOTAL: .Font.Bold = True
            .HorizontalAlignment = xlRight: .NumberFormat = "0"
        End With
        ws.Range(ws.Cells(curRow, sCol), ws.Cells(curRow, totalCol)) _
          .Borders(xlEdgeBottom).LineStyle = xlDot

        ' record this row in RaceMeta for later surgical refresh -- only
        ' when it has a real node_index, since Null can't serve as a match
        ' key (a row without one still shows on screen, just isn't tracked
        ' for surgical refresh until "Seed Now" gives it a real node).
        If hasNodeIdx Then
            mMetaWs.Cells(mMetaRow, 1).Value = classId
            mMetaWs.Cells(mMetaRow, 2).Value = heatId
            mMetaWs.Cells(mMetaRow, 3).Value = CStr(nd("node_index"))
            mMetaWs.Cells(mMetaRow, 4).Value = curRow
            mMetaWs.Cells(mMetaRow, 5).Value = sCol + 1
            mMetaWs.Cells(mMetaRow, 6).Value = sCol + 2
            mMetaWs.Cells(mMetaRow, 7).Value = rounds
            mMetaWs.Cells(mMetaRow, 8).Value = winCondition
            mMetaRow = mMetaRow + 1
        End If

        pilotCount = pilotCount + 1
        curRow = curRow + 1
SkipEmptyNode:
    Next nd

    If pilotCount = 0 Then
        ws.Cells(curRow, sCol).Value = "(empty)"
        ws.Cells(curRow, sCol).Font.Italic = True
        curRow = curRow + 1
    End If

    With ws.Range(ws.Cells(sRow, sCol), ws.Cells(curRow - 1, totalCol))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlMedium
    End With

    RenderHeat = curRow
End Function

' True if a node's pilot_id variant is a real, non-blank pilot id. Written as
' explicit If/ElseIf branches (real short-circuiting) rather than
' "IsNull(v) Or CStr(v) = """ -- VBA's Or evaluates BOTH sides even when the
' first is already True, so CStr(Null) still runs and raises "Invalid use of
' Null" (this crashed RenderHeat/ApplyNodeResultRow on any empty slot).
Private Function HasPilotValue(v As Variant) As Boolean
    If IsNull(v) Or IsEmpty(v) Then
        HasPilotValue = False
    ElseIf CStr(v) = "" Then
        HasPilotValue = False
    Else
        HasPilotValue = True
    End If
End Function

' True if a node's node_index variant is a real value (not Null). A slot
' pending "Seed Now" resolution in RotorHazard -- typically when there are
' fewer configured frequencies than pilots needing one -- has
' node_index=Null; CLng(Null)/CStr(Null) both raise "Invalid use of Null",
' so every place that reads node_index must check this first.
Private Function HasNodeIndex(v As Variant) As Boolean
    HasNodeIndex = Not (IsNull(v) Or IsEmpty(v))
End Function

' Get pilot callsign from Config (col A=ID, C=Callsign)
Private Function GetCallsign(pid As Long) As String
    Dim ws  As Worksheet
    Dim lr  As Long
    Dim i   As Long
    Dim cs  As String
    Set ws = ThisWorkbook.Sheets("Config")
    lr = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    For i = 2 To lr
        If CLng(Val(CStr(ws.Cells(i, 1).Value))) = pid Then
            cs = Trim(CStr(ws.Cells(i, 3).Value))
            If cs <> "" Then GetCallsign = cs Else GetCallsign = Trim(CStr(ws.Cells(i, 2).Value))
            Exit Function
        End If
    Next i
    GetCallsign = "Pilot#" & pid
End Function

' Get or create a very-hidden helper sheet, clearing its previous contents.
Private Function GetOrCreateHiddenSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
    ws.Visible = xlSheetVeryHidden
    Set GetOrCreateHiddenSheet = ws
End Function

' Column number to Excel letter (e.g. 28 -> "AB")
Private Function ColLtr(n As Integer) As String
    Dim s As String: s = ""
    Do While n > 0
        s = Chr(65 + (n - 1) Mod 26) & s
        n = (n - 1) \ 26
    Loop
    ColLtr = s
End Function

' Bubble sort string array by numeric value
Private Sub SortAsLong(arr() As String, count As Integer)
    Dim i As Integer, j As Integer, tmp As String
    For i = 1 To count - 1
        For j = i + 1 To count
            If CLng(arr(i)) > CLng(arr(j)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub
