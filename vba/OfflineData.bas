' ================================================================
'  OfflineData  -  Local (no-RotorHazard-connection) data operations
'  Compatible with: Formato_Carrera_de_drones.xlsm
'
'  WHAT THIS MODULE DOES:
'    Backs every "Carrera" sheet action with a LOCAL equivalent when
'    ThisWorkbook.WorkOffline() is True (Config!E14), so nothing needs a
'    network call to RotorHazard: generating heats, recomputing standings,
'    remixing pilots, and deleting a heat all work purely off data already
'    on the sheet plus the pilot list cached in Config!A:C.
'
'    The key idea: JsonConverter.ParseJson (used everywhere online) just
'    returns plain Scripting.Dictionary (objects) and Collection (arrays) --
'    standard VBA/COM types, nothing library-specific. So this module builds
'    the SAME SHAPE by hand (no JSON text involved at all) and hands it to
'    the exact same RenderGroup/RenderSingleClassInPlace/RenderStandings
'    pipeline in RenderRaceTables.bas -- nothing there needed to change to
'    support offline rendering, only to call these functions instead of an
'    HTTP request.
'
'  HOW TO INSTALL:
'    1. Alt+F11 to open VBA Editor
'    2. File > Import File... > select this .bas file
'
'  CONFIG CELLS USED:
'    Config!E14 -- offline toggle (read by the existing ThisWorkbook.WorkOffline())
'    Config!L7 -- "rounds to count" for the Laps_Time__Best_X_Rounds standings
'                 method offline (defaults to 3 if blank/invalid)
'    Config!E16 -- persistent counter for minting local (negative) ids
'    Config!A:C -- pilot id/name/callsign (already used online by GetCallsign)
'
'  IDS: RotorHazard always uses positive ids. Classes/heats created offline
'  get NEGATIVE ids (NextLocalId), so they can never collide with real ids
'  and every existing CLng-based id comparison in RenderRaceTables.bas keeps
'  working unchanged.
' ================================================================

Option Explicit

' -- Mint the next local (negative) id for an offline-created class/heat. --
Public Function NextLocalId() As Long
    Dim cfg As Worksheet
    Dim cur As Variant
    Set cfg = ThisWorkbook.Sheets("Config")
    cur = cfg.Range("E16").Value
    If Not IsNumeric(cur) Then cur = 0
    NextLocalId = CLng(cur) - 1
    cfg.Range("E16").Value = NextLocalId
End Function

' How many of a pilot's best rounds to sum for the offline time-based
' standings method -- Config!L7, default 3 if blank/invalid.
Public Function GetRoundsToCount() As Long
    Dim v As Variant
    v = ThisWorkbook.Sheets("Config").Range("L7").Value
    If IsNumeric(v) Then
        If CLng(v) > 0 Then
            GetRoundsToCount = CLng(v)
            Exit Function
        End If
    End If
    GetRoundsToCount = 3
End Function

' -- Read Config!A:C into a Collection of Dictionary{id, name, callsign} --
' same shape the online pilot pickers already build from the API's
' /api/rhm/pilots response.
Public Function GetOfflinePilots() As Collection
    Dim cfg As Worksheet
    Dim lr As Long, i As Long
    Dim result As New Collection
    Dim p As Object

    Set cfg = ThisWorkbook.Sheets("Config")
    lr = cfg.Cells(cfg.Rows.count, 1).End(xlUp).row
    For i = 2 To lr
        If Trim(CStr(cfg.Cells(i, 1).Value)) <> "" Then
            Set p = CreateObject("Scripting.Dictionary")
            p("id") = CLng(Val(CStr(cfg.Cells(i, 1).Value)))
            p("name") = Trim(CStr(cfg.Cells(i, 2).Value))
            p("callsign") = Trim(CStr(cfg.Cells(i, 3).Value))
            result.Add p
        End If
    Next i
    Set GetOfflinePilots = result
End Function

' Reverse of GetCallsignForPilot -- pilot id from a callsign already shown
' on the sheet (used when harvesting a hand-typed roster back into
' pilot_ids). Returns 0 if not found/blank/"(vacio)".
Public Function PilotIdByCallsign(callsign As String) As Long
    Dim cfg As Worksheet
    Dim lr As Long, i As Long
    callsign = Trim(callsign)
    If callsign = "" Or callsign = "(vacio)" Then
        PilotIdByCallsign = 0
        Exit Function
    End If
    Set cfg = ThisWorkbook.Sheets("Config")
    lr = cfg.Cells(cfg.Rows.count, 1).End(xlUp).row
    For i = 2 To lr
        If Trim(CStr(cfg.Cells(i, 3).Value)) = callsign Then
            PilotIdByCallsign = CLng(Val(CStr(cfg.Cells(i, 1).Value)))
            Exit Function
        End If
    Next i
    PilotIdByCallsign = 0
End Function

' Same lookup RenderRaceTables.GetCallsign does (col A=id, C=callsign,
' B=name fallback) -- kept as its own small copy here since that one is
' Private to its own module.
Public Function GetCallsignForPilot(pid As Long) As String
    Dim cfg As Worksheet, lr As Long, i As Long, cs As String
    Set cfg = ThisWorkbook.Sheets("Config")
    lr = cfg.Cells(cfg.Rows.count, 1).End(xlUp).row
    For i = 2 To lr
        If CLng(Val(CStr(cfg.Cells(i, 1).Value))) = pid Then
            cs = Trim(CStr(cfg.Cells(i, 3).Value))
            If cs <> "" Then GetCallsignForPilot = cs Else GetCallsignForPilot = Trim(CStr(cfg.Cells(i, 2).Value))
            Exit Function
        End If
    Next i
    GetCallsignForPilot = "Pilot#" & pid
End Function

' ================================================================
'  SHUFFLE -- direct VBA port of manager_uc.generate_heats_logic (Python).
'  That function is pure (list/shuffle/pair-count only, no server or DB
'  calls), so this reproduces it exactly: rotate the roster per group,
'  chunk into heats, try up to 15 random shuffles per heat keeping the one
'  that repeats the fewest already-seen pairings.
' ================================================================

' pilots: Collection of callsign strings. Returns a Collection of groups,
' each a Collection of heats, each a Collection of callsign strings (same
' nested shape the Python version returns as a nested list).
Public Function ShuffleHeatsLocal(pilots As Collection, pilotsPerHeat As Long, numGroups As Long) As Collection
    Dim allGroups As New Collection
    Dim encounters As Object
    Dim g As Long, offset As Long, i As Long, j As Long
    Dim pilotArr() As String
    Dim n As Long
    Dim heatsInGroup As Collection
    Dim heatStart As Long

    Randomize

    n = pilots.count
    ReDim pilotArr(1 To n)
    For i = 1 To n
        pilotArr(i) = CStr(pilots(i))
    Next i

    Set encounters = CreateObject("Scripting.Dictionary")

    For g = 0 To numGroups - 1
        If n > 0 Then offset = (g * 3) Mod n Else offset = 0
        Dim rotated() As String
        ReDim rotated(1 To n)
        For i = 1 To n
            rotated(i) = pilotArr(((i - 1 + offset) Mod n) + 1)
        Next i

        Set heatsInGroup = New Collection
        For heatStart = 1 To n Step pilotsPerHeat
            Dim heatSize As Long
            heatSize = pilotsPerHeat
            If heatStart + heatSize - 1 > n Then heatSize = n - heatStart + 1

            Dim heatArr() As String
            ReDim heatArr(1 To heatSize)
            For i = 1 To heatSize
                heatArr(i) = rotated(heatStart + i - 1)
            Next i

            Dim best() As String
            Dim bestCost As Long, cost As Long, tryNum As Long
            ReDim best(1 To heatSize)
            For i = 1 To heatSize: best(i) = heatArr(i): Next i
            bestCost = 999999

            For tryNum = 1 To 15
                ShuffleArrayLocal heatArr
                cost = 0
                For i = 1 To heatSize
                    For j = i + 1 To heatSize
                        cost = cost + PairCount(encounters, heatArr(i), heatArr(j))
                    Next j
                Next i
                If cost < bestCost Then
                    bestCost = cost
                    For i = 1 To heatSize: best(i) = heatArr(i): Next i
                End If
            Next tryNum

            Dim heatColl As New Collection
            For i = 1 To heatSize
                heatColl.Add best(i)
            Next i
            heatsInGroup.Add heatColl

            For i = 1 To heatSize
                For j = i + 1 To heatSize
                    IncPairCount encounters, best(i), best(j)
                Next j
            Next i
        Next heatStart

        allGroups.Add heatsInGroup
    Next g

    Set ShuffleHeatsLocal = allGroups
End Function

' Fisher-Yates shuffle of a 1-based String array in place.
Private Sub ShuffleArrayLocal(arr() As String)
    Dim i As Long, j As Long, tmp As String
    Dim lo As Long, hi As Long
    lo = LBound(arr): hi = UBound(arr)
    For i = hi To lo + 1 Step -1
        j = lo + Int(Rnd * (i - lo + 1))
        tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
    Next i
End Sub

Private Function PairKey(a As String, b As String) As String
    If a < b Then
        PairKey = a & vbNullChar & b
    Else
        PairKey = b & vbNullChar & a
    End If
End Function

Private Function PairCount(encounters As Object, a As String, b As String) As Long
    Dim k As String
    k = PairKey(a, b)
    If encounters.Exists(k) Then
        PairCount = encounters(k)
    Else
        PairCount = 0
    End If
End Function

Private Sub IncPairCount(encounters As Object, a As String, b As String)
    Dim k As String
    k = PairKey(a, b)
    If encounters.Exists(k) Then
        encounters(k) = encounters(k) + 1
    Else
        encounters(k) = 1
    End If
End Sub

' ================================================================
'  BUILD -- assemble a class/heats Dictionary tree (matching the server's
'  get_class_results_grid shape) from a shuffle result, ready to hand to
'  RenderSingleClassInPlace.
' ================================================================

' Builds just the "heats" sub-dictionary -- shared by BuildLocalClassDict
' (brand new class) and GenerateHeatsLocal's "add to existing class" path
' (which only needs the new heats to merge into an already-harvested dict).
Private Function BuildHeatsDict(classId As Long, className As String, shuffled As Collection, groupOffset As Long) As Object
    Dim heatsDict As Object
    Dim g As Long, h As Long, s As Long
    Dim group As Collection, heat As Collection
    Dim heatId As Long, heatName As String
    Dim heatDict As Object, nodesColl As Collection, nodeDict As Object
    Dim pid As Long

    Set heatsDict = CreateObject("Scripting.Dictionary")
    For g = 1 To shuffled.count
        Set group = shuffled(g)
        For h = 1 To group.count
            Set heat = group(h)
            heatId = NextLocalId()
            heatName = className & " G" & (groupOffset + g - 1) & "-H" & h

            Set nodesColl = New Collection
            For s = 1 To heat.count
                Set nodeDict = CreateObject("Scripting.Dictionary")
                nodeDict("slot_id") = 0
                pid = PilotIdByCallsign(CStr(heat(s)))
                If pid > 0 Then nodeDict("pilot_id") = pid Else nodeDict("pilot_id") = Null
                nodeDict("node_index") = s - 1
                nodeDict("freq_label") = "N" & s
                Set nodeDict("rounds") = New Collection
                nodesColl.Add nodeDict
            Next s

            Set heatDict = CreateObject("Scripting.Dictionary")
            heatDict("id") = heatId
            heatDict("display_name") = heatName
            heatDict("class_id") = classId
            Set heatDict("nodes") = nodesColl

            Set heatsDict(CStr(heatId)) = heatDict
        Next h
    Next g

    Set BuildHeatsDict = heatsDict
End Function

' Full class dict for a brand-new offline class.
Public Function BuildLocalClassDict(classId As Long, className As String, winCondition As String, rounds As Long, shuffled As Collection, groupOffset As Long) As Object
    Dim clsDict As Object
    Set clsDict = CreateObject("Scripting.Dictionary")
    clsDict("id") = classId
    clsDict("display_name") = className
    clsDict("rounds") = rounds
    clsDict("win_condition") = winCondition
    Set clsDict("heats") = BuildHeatsDict(classId, className, shuffled, groupOffset)
    clsDict("standings") = Null
    Set BuildLocalClassDict = clsDict
End Function

' -- Full offline "Generar Heats" flow: shuffle pilots into groups/heats
' locally and, if extending an existing local class, merge the new heats
' into that class's current (harvested) heats so nothing already entered
' by hand gets lost -- then redraw via RenderSingleClassInPlace, same as
' the online success path. existingClassId = 0 means "create a new class".
Public Sub GenerateHeatsLocal(existingClassId As Long, stageName As String, winCondition As String, roundsPerGroup As Long, numGroups As Long, pilotsPerHeat As Long, pilotCallsigns As Collection)
    Dim classDict As Object
    Dim classId As Long
    Dim groupOffset As Long
    Dim shuffled As Collection

    If pilotCallsigns.count = 0 Then
        MsgBox "Selecciona al menos un piloto.", vbExclamation
        Exit Sub
    End If

    Set shuffled = ShuffleHeatsLocal(pilotCallsigns, pilotsPerHeat, numGroups)

    If existingClassId <> 0 Then
        Set classDict = HarvestClassFromSheet(existingClassId)
        If classDict Is Nothing Then
            MsgBox "No se encontro la clase local elegida.", vbExclamation
            Exit Sub
        End If
        classId = existingClassId
        groupOffset = NextGroupNumberFromHeats(classDict)

        Dim newHeats As Object, nk As Variant
        Set newHeats = BuildHeatsDict(classId, CStr(classDict("display_name")), shuffled, groupOffset)
        For Each nk In newHeats.keys
            Set classDict("heats")(nk) = newHeats(nk)
        Next nk
    Else
        classId = NextLocalId()
        groupOffset = 1
        Set classDict = BuildLocalClassDict(classId, stageName, winCondition, roundsPerGroup, shuffled, groupOffset)
    End If

    Set classDict("standings") = ComputeStandingsLocal(classDict)
    RenderSingleClassInPlace classDict
End Sub

' id+name pairs for every class currently drawn on the sheet (RaceMetaGroups),
' regardless of whether it was created online or offline -- once it's on the
' sheet, the sheet is the source of truth for "add heats to existing class"
' while offline. Used by frmCreateHeats's ComboBoxExistingClass offline.
Public Function GetLocalClassList() As Collection
    Dim result As New Collection
    Dim metaGroupWs As Worksheet, ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim item As Object

    On Error Resume Next
    Set metaGroupWs = ThisWorkbook.Sheets("RaceMetaGroups")
    On Error GoTo 0
    If metaGroupWs Is Nothing Then
        Set GetLocalClassList = result
        Exit Function
    End If

    Set ws = ThisWorkbook.Sheets("Carrera")
    lastRow = metaGroupWs.Cells(metaGroupWs.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        Dim cId As Long, hRow As Long, sC As Long
        cId = CLng(metaGroupWs.Cells(i, 1).Value)
        hRow = CLng(metaGroupWs.Cells(i, 2).Value)
        sC = CLng(metaGroupWs.Cells(i, 3).Value)
        Set item = CreateObject("Scripting.Dictionary")
        item("id") = cId
        item("name") = CStr(ws.Cells(hRow, sC).Value)
        result.Add item
    Next i
    Set GetLocalClassList = result
End Function

' Highest "G<n>" found in a class's current heat display_names, plus one --
' where new offline-added heats should continue numbering from.
Public Function NextGroupNumberFromHeats(classDict As Object) As Long
    Dim heats As Object, hk As Variant, heatObj As Object
    Dim maxG As Long, g As Long
    maxG = 0
    If classDict Is Nothing Then
        NextGroupNumberFromHeats = 1
        Exit Function
    End If
    Set heats = classDict("heats")
    For Each hk In heats.keys
        Set heatObj = heats(hk)
        g = ParseGroupNumber(CStr(heatObj("display_name")))
        If g > maxG Then maxG = g
    Next hk
    NextGroupNumberFromHeats = maxG + 1
End Function

' Parses the "G<n>" number out of a heat display_name like "MyClass G2-H1"
' -- returns 1 if the pattern isn't found (treats it as a single group).
Public Function ParseGroupNumber(displayName As String) As Long
    Dim p As Long, q As Long
    Dim numStr As String
    p = InStrRev(displayName, " G")
    If p = 0 Then
        ParseGroupNumber = 1
        Exit Function
    End If
    q = p + 2
    numStr = ""
    Do While q <= Len(displayName) And Mid(displayName, q, 1) >= "0" And Mid(displayName, q, 1) <= "9"
        numStr = numStr & Mid(displayName, q, 1)
        q = q + 1
    Loop
    If numStr = "" Then
        ParseGroupNumber = 1
    Else
        ParseGroupNumber = CLng(numStr)
    End If
End Function

' Groups (G1, G2, ...) for a class, derived from its heats' names -- same
' shape (Collection of Dictionary{group_id, label, heats}) the online
' GET /api/rhm/raceclass/<id>/groups endpoint returns, so RemixGroupAndRedraw
' and frmRemixGroups can use either source without caring which one it is.
Public Function GetLocalClassGroups(classId As Long) As Collection
    Dim classDict As Object
    Dim heats As Object, hk As Variant, heatObj As Object
    Dim groupsByNum As Object
    Dim result As New Collection
    Dim gNum As Long

    Set classDict = HarvestClassFromSheet(classId)
    If classDict Is Nothing Then
        Set GetLocalClassGroups = result
        Exit Function
    End If

    Set heats = classDict("heats")
    Set groupsByNum = CreateObject("Scripting.Dictionary")

    For Each hk In heats.keys
        Set heatObj = heats(hk)
        gNum = ParseGroupNumber(CStr(heatObj("display_name")))
        Dim hInfo As Object
        Set hInfo = CreateObject("Scripting.Dictionary")
        hInfo("id") = heatObj("id")
        hInfo("display_name") = heatObj("display_name")
        If Not groupsByNum.Exists(gNum) Then
            Set groupsByNum(gNum) = New Collection
        End If
        groupsByNum(gNum).Add hInfo
    Next hk

    Dim n As Long, i As Long, j As Long
    Dim gKeys() As Long
    n = groupsByNum.count
    If n = 0 Then
        Set GetLocalClassGroups = result
        Exit Function
    End If
    ReDim gKeys(1 To n)
    i = 0
    Dim gk As Variant
    For Each gk In groupsByNum.keys
        i = i + 1
        gKeys(i) = CLng(gk)
    Next gk
    Dim tmp As Long
    For i = 1 To n - 1
        For j = i + 1 To n
            If gKeys(j) < gKeys(i) Then
                tmp = gKeys(i): gKeys(i) = gKeys(j): gKeys(j) = tmp
            End If
        Next j
    Next i

    For i = 1 To n
        Dim gDict As Object
        Set gDict = CreateObject("Scripting.Dictionary")
        gDict("group_id") = gKeys(i) - 1 ' 0-based, matching the server's group_id
        gDict("label") = "G" & gKeys(i)
        Set gDict("heats") = groupsByNum(gKeys(i))
        result.Add gDict
    Next i

    Set GetLocalClassGroups = result
End Function

' ================================================================
'  HARVEST -- read back whatever the user typed into Pilot/Fr/R1..Rn cells
'  (via the RaceMeta row that already records which cell is which), and
'  reconstruct the SAME class/heats/nodes/rounds tree shape a server
'  response would have. This is what lets every offline flow (recompute
'  standings, remix, delete a heat, add heats) work off "whatever's on the
'  sheet right now" without caring whether that data ever came from RH.
' ================================================================

Public Function HarvestClassFromSheet(classId As Long) As Object
    Dim ws As Worksheet, metaWs As Worksheet, metaGroupWs As Worksheet
    Dim groupMetaRow As Long, i As Long, lastRow As Long, lastGroupRow As Long
    Dim sCol As Long, headerRow As Long
    Dim winCondition As String, rounds As Long
    Dim clsDict As Object, heatsDict As Object
    Dim nodesByHeat As Object, minRowByHeat As Object
    Dim className As String

    Set ws = ThisWorkbook.Sheets("Carrera")
    Set metaWs = ThisWorkbook.Sheets("RaceMeta")
    Set metaGroupWs = ThisWorkbook.Sheets("RaceMetaGroups")

    groupMetaRow = 0
    lastGroupRow = metaGroupWs.Cells(metaGroupWs.Rows.count, 1).End(xlUp).row
    For i = 2 To lastGroupRow
        If CLng(metaGroupWs.Cells(i, 1).Value) = classId Then
            groupMetaRow = i
            Exit For
        End If
    Next i
    If groupMetaRow = 0 Then
        Set HarvestClassFromSheet = Nothing
        Exit Function
    End If

    headerRow = CLng(metaGroupWs.Cells(groupMetaRow, 2).Value)
    sCol = CLng(metaGroupWs.Cells(groupMetaRow, 3).Value)
    className = CStr(ws.Cells(headerRow, sCol).Value)

    Set nodesByHeat = CreateObject("Scripting.Dictionary")
    Set minRowByHeat = CreateObject("Scripting.Dictionary")
    winCondition = ""
    rounds = 0

    lastRow = metaWs.Cells(metaWs.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If CLng(metaWs.Cells(i, 1).Value) = classId Then
            Dim hIdStr As String
            Dim rowNum As Long, frCol As Long, firstRoundCol As Long, nodeRounds As Long
            Dim pilotCol As Long, r As Long, nodeIndex As Long
            Dim callsignText As String, freqText As String
            Dim nodeDict As Object, roundsColl As Collection
            Dim cellVal As Variant

            hIdStr = CStr(metaWs.Cells(i, 2).Value)
            nodeIndex = CLng(metaWs.Cells(i, 3).Value)
            rowNum = CLng(metaWs.Cells(i, 4).Value)
            frCol = CLng(metaWs.Cells(i, 5).Value)
            firstRoundCol = CLng(metaWs.Cells(i, 6).Value)
            nodeRounds = CLng(metaWs.Cells(i, 7).Value)
            If winCondition = "" Then winCondition = CStr(metaWs.Cells(i, 8).Value)
            If rounds = 0 Then rounds = nodeRounds
            pilotCol = frCol - 1

            callsignText = Trim(CStr(ws.Cells(rowNum, pilotCol).Value))
            freqText = Trim(CStr(ws.Cells(rowNum, frCol).Value))

            Set roundsColl = New Collection
            For r = 1 To nodeRounds
                cellVal = ws.Cells(rowNum, firstRoundCol + r - 1).Value
                Dim rd As Object
                Set rd = CreateObject("Scripting.Dictionary")
                rd("round") = r
                rd("laps") = Null
                rd("time") = Null
                rd("position") = Null
                rd("points") = Null
                Select Case winCondition
                    Case "Cumulative_Points"
                        If IsNumeric(cellVal) Then rd("points") = CDbl(cellVal)
                    Case "Last_Heat_Position"
                        If IsNumeric(cellVal) Then rd("position") = CDbl(cellVal)
                    Case "Laps_Time__Best_X_Rounds"
                        Dim timeText As String
                        timeText = Trim(CStr(cellVal))
                        If timeText <> "" And timeText <> "-" Then rd("time") = timeText
                    Case Else
                        If IsNumeric(cellVal) Then rd("laps") = CDbl(cellVal)
                End Select
                roundsColl.Add rd
            Next r

            Set nodeDict = CreateObject("Scripting.Dictionary")
            nodeDict("slot_id") = 0
            If callsignText <> "" And callsignText <> "(vacio)" Then
                Dim pid As Long
                pid = PilotIdByCallsign(callsignText)
                If pid > 0 Then nodeDict("pilot_id") = pid Else nodeDict("pilot_id") = Null
            Else
                nodeDict("pilot_id") = Null
            End If
            nodeDict("node_index") = nodeIndex
            If freqText <> "" And freqText <> "?" Then nodeDict("freq_label") = freqText Else nodeDict("freq_label") = Null
            Set nodeDict("rounds") = roundsColl

            If Not nodesByHeat.Exists(hIdStr) Then
                Set nodesByHeat(hIdStr) = New Collection
                minRowByHeat(hIdStr) = rowNum
            ElseIf rowNum < minRowByHeat(hIdStr) Then
                minRowByHeat(hIdStr) = rowNum
            End If
            nodesByHeat(hIdStr).Add nodeDict
        End If
    Next i

    Set heatsDict = CreateObject("Scripting.Dictionary")
    Dim hk As Variant
    For Each hk In nodesByHeat.keys
        Dim heatDict As Object
        Dim titleRow As Long
        titleRow = CLng(minRowByHeat(hk)) - 1
        Set heatDict = CreateObject("Scripting.Dictionary")
        heatDict("id") = CLng(hk)
        heatDict("display_name") = CStr(ws.Cells(titleRow, sCol).Value)
        heatDict("class_id") = classId
        Set heatDict("nodes") = nodesByHeat(hk)
        Set heatsDict(hk) = heatDict
    Next hk

    Set clsDict = CreateObject("Scripting.Dictionary")
    clsDict("id") = classId
    clsDict("display_name") = className
    clsDict("rounds") = rounds
    clsDict("win_condition") = winCondition
    Set clsDict("heats") = heatsDict
    clsDict("standings") = Null

    Set HarvestClassFromSheet = clsDict
End Function

' Find which class a heat belongs to by scanning RaceMeta -- used by
' frmManageHeat offline (RenderRaceTables.bas has its own private copy of
' this same lookup for its own callers).
Public Function FindClassIdForHeatLocal(heatId As Long) As Long
    Dim metaWs As Worksheet
    Dim lastRow As Long, i As Long
    On Error Resume Next
    Set metaWs = ThisWorkbook.Sheets("RaceMeta")
    On Error GoTo 0
    If metaWs Is Nothing Then
        FindClassIdForHeatLocal = 0
        Exit Function
    End If
    lastRow = metaWs.Cells(metaWs.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If CLng(metaWs.Cells(i, 2).Value) = heatId Then
            FindClassIdForHeatLocal = CLng(metaWs.Cells(i, 1).Value)
            Exit Function
        End If
    Next i
    FindClassIdForHeatLocal = 0
End Function

' -- PUBLIC: write a new pilot callsign (or "" to clear it) directly into
' the Carrera cell for one heat/node, via its RaceMeta row -- offline
' equivalent of POST /api/rhm/heat/<id> for a single slot. Only works for a
' slot that's already occupied (has a RaceMeta row); a slot that was never
' occupied isn't tracked offline at all (empty slots aren't drawn/recorded
' -- see RenderHeat in RenderRaceTables.bas), so it can't be assigned a
' first pilot through this form while offline. --
Public Sub SetPilotForNodeOffline(heatId As Long, nodeIndex As Long, newCallsign As String)
    Dim metaWs As Worksheet, ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim rowNum As Long, frCol As Long, pilotCol As Long

    Set metaWs = ThisWorkbook.Sheets("RaceMeta")
    Set ws = ThisWorkbook.Sheets("Carrera")
    lastRow = metaWs.Cells(metaWs.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If CLng(metaWs.Cells(i, 2).Value) = heatId And CLng(metaWs.Cells(i, 3).Value) = nodeIndex Then
            rowNum = CLng(metaWs.Cells(i, 4).Value)
            frCol = CLng(metaWs.Cells(i, 5).Value)
            pilotCol = frCol - 1
            If Trim(newCallsign) = "" Then
                ws.Cells(rowNum, pilotCol).Value = "(vacio)"
            Else
                ws.Cells(rowNum, pilotCol).Value = newCallsign
            End If
            Exit Sub
        End If
    Next i
End Sub

' -- PUBLIC: remove one heat from a harvested class and redraw -- offline
' equivalent of DELETE /api/rhm/heat/<id> + redraw. Returns False if the
' class couldn't be found on the sheet at all. --
Public Function DeleteHeatLocal(classId As Long, heatId As Long) As Boolean
    Dim classDict As Object
    Set classDict = HarvestClassFromSheet(classId)
    If classDict Is Nothing Then
        DeleteHeatLocal = False
        Exit Function
    End If
    If classDict("heats").Exists(CStr(heatId)) Then
        classDict("heats").Remove CStr(heatId)
    End If
    Set classDict("standings") = ComputeStandingsLocal(classDict)
    RenderSingleClassInPlace classDict
    DeleteHeatLocal = True
End Function

' -- PUBLIC: offline remix -- reshuffle the pilots currently in the target
' group(s) (harvested straight off the sheet) using the same ShuffleHeatsLocal
' algorithm, clearing their round results (same as the online confirmation
' already warns), then redraw. groupIdsCsv is 0-based group ids, same as the
' online payload shape (RemixGroupAndRedraw/frmRemixGroups build this the
' same way regardless of online/offline). --
Public Function PerformRemixLocal(classId As Long, groupIdsCsv As String) As Boolean
    Dim classDict As Object
    Dim heats As Object, hk As Variant, heatObj As Object
    Dim targetGroupNums As Object
    Dim parts() As String, i As Long

    PerformRemixLocal = False

    Set classDict = HarvestClassFromSheet(classId)
    If classDict Is Nothing Then
        MsgBox "No se encontro la clase en la hoja.", vbExclamation
        Exit Function
    End If

    Set targetGroupNums = CreateObject("Scripting.Dictionary")
    parts = Split(groupIdsCsv, ",")
    For i = LBound(parts) To UBound(parts)
        targetGroupNums(CLng(Trim(parts(i))) + 1) = True ' group_id is 0-based, "G<n>" is 1-based
    Next i

    Set heats = classDict("heats")

    Dim groupHeatIds As New Collection
    Dim rosterCallsigns As New Collection
    Dim pilotsPerHeat As Long
    pilotsPerHeat = 0

    For Each hk In heats.keys
        Set heatObj = heats(hk)
        If targetGroupNums.Exists(ParseGroupNumber(CStr(heatObj("display_name")))) Then
            groupHeatIds.Add CStr(hk)
            Dim nd As Object, countOccupied As Long
            countOccupied = 0
            For Each nd In heatObj("nodes")
                If Not IsNull(nd("pilot_id")) Then
                    rosterCallsigns.Add GetCallsignForPilot(CLng(nd("pilot_id")))
                    countOccupied = countOccupied + 1
                End If
            Next nd
            If countOccupied > pilotsPerHeat Then pilotsPerHeat = countOccupied
        End If
    Next hk

    If rosterCallsigns.count = 0 Or groupHeatIds.count = 0 Then
        MsgBox "El/los grupo(s) seleccionado(s) no tienen pilotos para remixar.", vbExclamation
        Exit Function
    End If

    Dim shuffled As Collection, newHeatsFlat As Collection
    Set shuffled = ShuffleHeatsLocal(rosterCallsigns, pilotsPerHeat, 1)
    Set newHeatsFlat = shuffled(1)

    Dim orderedHeatIds() As String
    Dim cnt As Long
    cnt = groupHeatIds.count
    ReDim orderedHeatIds(1 To cnt)
    For i = 1 To cnt
        orderedHeatIds(i) = groupHeatIds(i)
    Next i
    SortHeatIdStrings orderedHeatIds, cnt

    For i = 1 To cnt
        Set heatObj = heats(orderedHeatIds(i))
        Dim newHeatPilots As Collection
        If i <= newHeatsFlat.count Then
            Set newHeatPilots = newHeatsFlat(i)
        Else
            Set newHeatPilots = New Collection
        End If

        Dim s As Long, nodes As Object, nd2 As Object
        Set nodes = heatObj("nodes")
        For s = 1 To nodes.count
            Set nd2 = nodes(s)
            If s <= newHeatPilots.count Then
                Dim pid2 As Long
                pid2 = PilotIdByCallsign(CStr(newHeatPilots(s)))
                If pid2 > 0 Then nd2("pilot_id") = pid2 Else nd2("pilot_id") = Null
            Else
                nd2("pilot_id") = Null
            End If
            Set nd2("rounds") = New Collection ' remixing clears past results
        Next s
    Next i

    Set classDict("standings") = ComputeStandingsLocal(classDict)
    RenderSingleClassInPlace classDict
    PerformRemixLocal = True
End Function

' Bubble sort a String array of (possibly negative) numeric-id strings.
Private Sub SortHeatIdStrings(arr() As String, count As Long)
    Dim i As Long, j As Long, tmp As String
    For i = 1 To count - 1
        For j = i + 1 To count
            If CLng(arr(i)) > CLng(arr(j)) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
End Sub

' ================================================================
'  STANDINGS -- pragmatic local equivalents of the 3 win_condition ranking
'  methods RotorHazard computes server-side. These are NOT byte-for-byte
'  reimplementations of RH's ranking engine (which has more configuration
'  and edge cases than a hand-entered sheet can express) -- they're a
'  faithful-enough approximation for manually-typed results:
'    Cumulative_Points        -> sum of a pilot's points, descending
'    Last_Heat_Position       -> a pilot's last recorded position, ascending
'    Laps_Time__Best_X_Rounds -> sum of a pilot's best Config!L7 round
'                                 times (seconds), ascending
' ================================================================

Public Function ComputeStandingsLocal(classDict As Object) As Object
    Dim winCondition As String, roundsToCount As Long
    Dim heats As Object, hk As Variant, heatObj As Object
    Dim nd As Object, rd As Object
    Dim pilots As Object
    Dim standingsColl As New Collection
    Dim st As Object

    winCondition = CStr(classDict("win_condition"))
    Set heats = classDict("heats")
    Set pilots = CreateObject("Scripting.Dictionary")

    For Each hk In heats.keys
        Set heatObj = heats(hk)
        For Each nd In heatObj("nodes")
            If Not IsNull(nd("pilot_id")) Then
                Dim pid As Long
                pid = CLng(nd("pilot_id"))
                If Not pilots.Exists(CStr(pid)) Then
                    Dim pinfo As Object
                    Set pinfo = CreateObject("Scripting.Dictionary")
                    pinfo("pilot_id") = pid
                    pinfo("callsign") = GetCallsignForPilot(pid)
                    Set pinfo("values") = New Collection
                    Set pilots(CStr(pid)) = pinfo
                End If
                For Each rd In nd("rounds")
                    Select Case winCondition
                        Case "Cumulative_Points"
                            If Not IsNull(rd("points")) Then pilots(CStr(pid))("values").Add CDbl(rd("points"))
                        Case "Last_Heat_Position"
                            If Not IsNull(rd("position")) Then pilots(CStr(pid))("values").Add CDbl(rd("position"))
                        Case "Laps_Time__Best_X_Rounds"
                            If Not IsNull(rd("time")) Then pilots(CStr(pid))("values").Add TimeTextToSeconds(CStr(rd("time")))
                    End Select
                Next rd
            End If
        Next nd
    Next hk

    roundsToCount = GetRoundsToCount()

    Dim scoreDict As Object
    Set scoreDict = CreateObject("Scripting.Dictionary")
    Dim pk As Variant
    For Each pk In pilots.keys
        Dim vals As Collection
        Set vals = pilots(pk)("values")
        Dim score As Double
        Select Case winCondition
            Case "Cumulative_Points"
                score = SumCollectionLocal(vals)
            Case "Last_Heat_Position"
                score = LastNonEmptyLocal(vals)
            Case "Laps_Time__Best_X_Rounds"
                score = SumBestNLocal(vals, roundsToCount)
            Case Else
                score = SumCollectionLocal(vals)
        End Select
        scoreDict(pk) = score
    Next pk

    Dim ascending As Boolean
    ascending = (winCondition <> "Cumulative_Points")

    Dim cnt As Long, idx As Long
    cnt = scoreDict.count
    Dim orderedKeys() As String
    ReDim orderedKeys(1 To cnt)
    idx = 0
    For Each pk In scoreDict.keys
        idx = idx + 1
        orderedKeys(idx) = CStr(pk)
    Next pk
    SortKeysByScore orderedKeys, scoreDict, ascending

    Dim pos As Long
    For pos = 1 To cnt
        Set st = CreateObject("Scripting.Dictionary")
        st("pilot_id") = pilots(orderedKeys(pos))("pilot_id")
        st("callsign") = pilots(orderedKeys(pos))("callsign")
        st("position") = pos
        Dim extra As Object
        Set extra = CreateObject("Scripting.Dictionary")
        Select Case winCondition
            Case "Cumulative_Points": extra("points") = scoreDict(orderedKeys(pos))
            Case "Last_Heat_Position": extra("last_position") = scoreDict(orderedKeys(pos))
            Case "Laps_Time__Best_X_Rounds": extra("best_time_sec") = Format(scoreDict(orderedKeys(pos)), "0.000")
        End Select
        Set st("extra") = extra
        standingsColl.Add st
    Next pos

    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("method_label") = MethodLabelFor(winCondition, roundsToCount)
    Set result("rank_fields") = New Collection
    Set result("standings") = standingsColl
    Set ComputeStandingsLocal = result
End Function

Private Function SumCollectionLocal(vals As Collection) As Double
    Dim v As Variant, total As Double
    total = 0
    For Each v In vals
        total = total + CDbl(v)
    Next v
    SumCollectionLocal = total
End Function

Private Function LastNonEmptyLocal(vals As Collection) As Double
    If vals.count = 0 Then
        LastNonEmptyLocal = 999999
    Else
        LastNonEmptyLocal = CDbl(vals(vals.count))
    End If
End Function

Private Function SumBestNLocal(vals As Collection, n As Long) As Double
    Dim arr() As Double
    Dim i As Long, cnt As Long
    cnt = vals.count
    If cnt = 0 Then
        SumBestNLocal = 999999
        Exit Function
    End If
    ReDim arr(1 To cnt)
    For i = 1 To cnt
        arr(i) = CDbl(vals(i))
    Next i
    Dim j As Long, tmp As Double
    For i = 1 To cnt - 1
        For j = i + 1 To cnt
            If arr(j) < arr(i) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
    Dim total As Double, take As Long
    take = n
    If take > cnt Then take = cnt
    total = 0
    For i = 1 To take
        total = total + arr(i)
    Next i
    SumBestNLocal = total
End Function

' Parses "M:SS.mmm" / "H:MM:SS.mmm" back to total seconds, using Val() (not
' CDbl) for the decimal parts -- Val always treats "." as the decimal point
' regardless of Windows locale, while CDbl would misparse "45.243" on a
' comma-decimal locale (e.g. Spanish/Colombian Windows). A blank/unparseable
' time gets a large sentinel so it sorts last instead of looking fastest.
Private Function TimeTextToSeconds(t As String) As Double
    Dim parts() As String
    Dim n As Long
    t = Trim(t)
    If t = "" Or t = "-" Then
        TimeTextToSeconds = 999999
        Exit Function
    End If
    parts = Split(t, ":")
    n = UBound(parts) - LBound(parts) + 1
    Select Case n
        Case 1
            TimeTextToSeconds = Val(parts(0))
        Case 2
            TimeTextToSeconds = Val(parts(0)) * 60 + Val(parts(1))
        Case 3
            TimeTextToSeconds = Val(parts(0)) * 3600 + Val(parts(1)) * 60 + Val(parts(2))
        Case Else
            TimeTextToSeconds = 999999
    End Select
End Function

Private Sub SortKeysByScore(keys() As String, scoreDict As Object, ascending As Boolean)
    Dim i As Long, j As Long, tmp As String, doSwap As Boolean
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If ascending Then
                doSwap = (scoreDict(keys(j)) < scoreDict(keys(i)))
            Else
                doSwap = (scoreDict(keys(j)) > scoreDict(keys(i)))
            End If
            If doSwap Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function MethodLabelFor(winCondition As String, roundsToCount As Long) As String
    Select Case winCondition
        Case "Cumulative_Points": MethodLabelFor = "Cumulative Points (local)"
        Case "Last_Heat_Position": MethodLabelFor = "Last Heat Position (local)"
        Case "Laps_Time__Best_X_Rounds": MethodLabelFor = "Best " & roundsToCount & " Rounds (local)"
        Case Else: MethodLabelFor = "Local"
    End Select
End Function
