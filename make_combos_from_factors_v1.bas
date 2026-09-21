'==============================================================
' make_combos_from_factors_v1.bas
'--------------------------------------------------------------
' Creates Nastran LOAD Combination sets in FEMAP from a FACTORS
' table saved out of Excel as CSV.
'
' ONLY INPUT: CSV_PATH below. Everything else is read live:
'   - combination rows and factors  -> from the CSV
'   - load set IDs                  -> from FEMAP's own load set
'                                      titles, matched by text
' Nothing is hardcoded. Add rows to the sheet, re-save as CSV,
' re-run. The macro re-reads both sources fresh every time.
'
' CSV FORMAT (the FACTORS sheet, File > Save As > CSV)
'   Case Name       -> becomes the combination title
'   "<X> Factor"    -> one column per load. 0 or blank = unused.
'                      The column heading minus " Factor" is the
'                      FEMAP load set title to look for, e.g.
'                      "Inx Factor" -> load set titled "Inx"
'   "<X> Case"      -> optional. If present, the load title is
'                      built as  X (<case text>)  e.g.
'                      Ec Case = "H case 1" -> "Ec (H case 1)"
'   Any other column (e.g. FEMAP Load Set No.) is ignored.
'
' TITLE MATCHING
'   Case-insensitive, and spaces are ignored, so "Ec (H case 1)"
'   matches a FEMAP title "Ec (H Case 1)". Nothing else is
'   loosened. A name that finds no load set is NEVER guessed:
'   the whole row is skipped and reported by name.
'   A title that exists on two load sets is refused as ambiguous.
'
' IDS
'   Row n of the CSV becomes load set FIRST_ID + n - 1, so the
'   first 23 rows land on 131..153. Mapping is fixed per row -
'   a skipped row leaves a gap rather than shifting the others.
'   Every run recreates everything. If the range is occupied the
'   macro refuses to run: delete the old combinations first.
'
' RUN_MODE
'   0 = read and report only. Nothing is created.
'   1 = create the FIRST resolvable combination, then stop.
'   2 = create all resolvable combinations.
'
' >>> SAVE A BACKUP BEFORE RUN_MODE 1 OR 2 <<<
'
' API
'   PutCombination and the load set creation are the same calls
'   that built the 24 SCR combinations - proven on 2022.2.2.
'==============================================================

Const RUN_MODE           As Integer = 0
Const BACKUP_DONE        As Integer = 0

' The ONLY path to edit. The log goes to an "audit" folder
' created next to this file.
Const CSV_PATH           As String = "C:\CHANGE_THIS_FOLDER\FACTORS.csv"

Const FIRST_ID           As Long = 131

' 1 = title "H case-1 = Es + Ec (H Case 1) - Inx"
' 0 = title "H case-1"
Const TITLE_WITH_FORMULA As Integer = 1

Const MAX_SCAN_ID        As Long = 3000
Const MAX_FAILS          As Integer = 3

Dim femap As Object
Dim logF As Integer
Dim failCount As Integer
Dim outDir As String
Dim headerOK As Integer

' current CSV line split into fields
Dim fld(0 To 60) As String
Dim nFld As Integer
Dim delim As String

' header mapping
Dim colTitle As Integer
Dim famName(0 To 40) As String
Dim famFacCol(0 To 40) As Integer
Dim famCaseCol(0 To 40) As Integer
Dim nFam As Integer

' live FEMAP load set table
Dim ldKey(0 To 3000) As String
Dim ldTitle(0 To 3000) As String
Dim ldID(0 To 3000) As Long
Dim ldDup(0 To 3000) As Integer
Dim nLd As Integer

' parsed rows
Dim rTitle(0 To 300) As String
Dim rFull(0 To 300) As String
Dim rOK(0 To 300) As Integer
Dim rReason(0 To 300) As String
Dim rLine(0 To 300) As Integer
Dim rN(0 To 300) As Integer
Dim rIDs(0 To 300, 0 To 40) As Long
Dim rFac(0 To 300, 0 To 40) As Double
Dim rTermTitle(0 To 300, 0 To 40) As String
Dim nRows As Integer

'==============================================================
Sub Main

    Dim i As Integer
    Dim p As Integer

    Set femap = GetObject(, "femap.model")

    If FileThere(CSV_PATH) = 0 Then
        femap.feAppMessage 2, "FACTORS CSV not found: " & CSV_PATH
        femap.feAppMessage 2, "Edit CSV_PATH at the top of the macro."
        End
    End If

    ' audit folder sits next to the CSV
    p = 0
    For i = 1 To Len(CSV_PATH)
        If Mid(CSV_PATH, i, 1) = "\" Then
            p = i
        End If
    Next i
    outDir = Left(CSV_PATH, p) & "audit\"

    If EnsureDir(outDir) = 0 Then
        femap.feAppMessage 2, "Cannot create " & outDir
        End
    End If

    If RUN_MODE > 0 Then
        If BACKUP_DONE <> 1 Then
            femap.feAppMessage 2, "REFUSING TO WRITE. Save a backup, then set BACKUP_DONE = 1."
            End
        End If
    End If

    failCount = 0

    logF = FreeFile
    Open outDir & "make_combos.log" For Output As #logF

    WL "=== make_combos_from_factors_v1 ==="
    WL "RUN_MODE = " & RUN_MODE
    WL "CSV      = " & CSV_PATH
    WL ""

    Call BuildLoadTable
    Call ReadCsv

    If headerOK = 0 Then
        Close #logF
        femap.feAppMessage 2, "CSV header not recognised. See " & outDir & "make_combos.log"
        End
    End If

    Call Report

    If RUN_MODE = 0 Then
        Call CheckIds(0)
        WL ""
        WL "  [report only] nothing created."
        WL "  Check every combination above against the sheet, then"
        WL "  back up, set BACKUP_DONE = 1 and RUN_MODE = 1."
    Else
        Call CheckIds(1)
        Call CreateAll
    End If

    Close #logF
    femap.feAppMessage 0, "Done. See " & outDir & "make_combos.log"

End Sub

'--------------------------------------------------------------
' Read every existing load set title from FEMAP into a lookup.
'--------------------------------------------------------------
Sub BuildLoadTable

    Dim oLS As Object
    Dim lid As Long
    Dim t As String
    Dim k As String
    Dim j As Integer
    Dim hit As Integer

    nLd = 0

    Set oLS = femap.feLoadSet                             'VERIFIED

    For lid = 1 To MAX_SCAN_ID

        If oLS.Get(lid) = -1 Then                         'VERIFIED

            t = oLS.title
            k = NormKey(t)

            hit = -1
            For j = 0 To nLd - 1
                If ldKey(j) = k Then
                    hit = j
                End If
            Next j

            If hit >= 0 Then
                ldDup(hit) = 1
            Else
                If nLd <= 3000 Then
                    ldKey(nLd) = k
                    ldTitle(nLd) = t
                    ldID(nLd) = lid
                    ldDup(nLd) = 0
                    nLd = nLd + 1
                End If
            End If

        End If

    Next lid

End Sub

'--------------------------------------------------------------
' Returns an index into the ld arrays, or -1 if no load set has
' that title.
'--------------------------------------------------------------
Function FindLoad(want As String) As Integer

    Dim k As String
    Dim j As Integer

    FindLoad = -1
    k = NormKey(want)

    For j = 0 To nLd - 1
        If ldKey(j) = k Then
            FindLoad = j
            Exit Function
        End If
    Next j

End Function

'--------------------------------------------------------------
Sub ReadCsv

    Dim f As Integer
    Dim hdr As String
    Dim ln As String
    Dim lineNo As Integer
    Dim t As String

    nRows = 0
    headerOK = 0

    f = FreeFile
    Open CSV_PATH For Input As #f

    If EOF(f) Then
        Close #f
        WL "*** CSV is empty."
        Exit Sub
    End If

    Line Input #f, hdr

    ' strip a UTF-8 byte-order mark if Excel wrote one
    Do While Len(hdr) > 0
        If Asc(Left(hdr, 1)) > 127 Then
            hdr = Mid(hdr, 2)
        Else
            Exit Do
        End If
    Loop

    ' some Excel locales save CSV with ; instead of ,
    delim = ","
    If InStr(hdr, ",") = 0 Then
        If InStr(hdr, ";") > 0 Then
            delim = ";"
        End If
    End If

    Call SplitLine(hdr)
    headerOK = MapHeader()

    If headerOK = 0 Then
        Close #f
        Exit Sub
    End If

    lineNo = 1

    Do While Not EOF(f)

        Line Input #f, ln
        lineNo = lineNo + 1

        If Len(Trim(ln)) > 0 Then

            Call SplitLine(ln)

            t = ""
            If colTitle < nFld Then
                t = Trim(fld(colTitle))
            End If

            ' Excel often writes empty ",,,," rows below the table
            If t <> "" Then
                If nRows <= 300 Then
                    Call ParseRow(nRows, lineNo)
                    nRows = nRows + 1
                End If
            End If

        End If

    Loop

    Close #f

End Sub

'--------------------------------------------------------------
' Works out which column holds the title, which hold factors,
' and which "<X> Case" column belongs to which load.
' Returns 1 if usable.
'--------------------------------------------------------------
Function MapHeader() As Integer

    Dim i As Integer
    Dim h As String
    Dim fam As String
    Dim idx As Integer
    Dim j As Integer

    MapHeader = 0
    colTitle = -1
    nFam = 0

    WL "--- CSV columns ---"

    For i = 0 To nFld - 1

        h = Trim(fld(i))

        If NormKey(h) = "CASENAME" Then
            colTitle = i
            WL "  col " & (i + 1) & "  '" & h & "'  -> combination title"
        ElseIf Len(h) > 7 And UCase(Right(h, 7)) = " FACTOR" Then
            fam = Trim(Left(h, Len(h) - 7))
            If nFam <= 40 Then
                famName(nFam) = fam
                famFacCol(nFam) = i
                famCaseCol(nFam) = -1
                nFam = nFam + 1
            End If
        End If

    Next i

    For i = 0 To nFld - 1

        h = Trim(fld(i))

        If Len(h) > 5 And UCase(Right(h, 5)) = " CASE" Then
            fam = Trim(Left(h, Len(h) - 5))
            idx = -1
            For j = 0 To nFam - 1
                If NormKey(famName(j)) = NormKey(fam) Then
                    idx = j
                End If
            Next j
            If idx >= 0 Then
                famCaseCol(idx) = i
            Else
                WL "  col " & (i + 1) & "  '" & h & "'  -> IGNORED, no matching '" & fam & " Factor' column"
            End If
        End If

    Next i

    For j = 0 To nFam - 1
        If famCaseCol(j) >= 0 Then
            WL "  col " & (famFacCol(j) + 1) & "  '" & famName(j) & " Factor'  -> load title  " & famName(j) & " (<col " & (famCaseCol(j) + 1) & ">)"
        Else
            WL "  col " & (famFacCol(j) + 1) & "  '" & famName(j) & " Factor'  -> load title  " & famName(j)
        End If
    Next j

    WL ""

    If colTitle < 0 Then
        WL "*** No 'Case Name' column found in the header."
        Exit Function
    End If

    If nFam = 0 Then
        WL "*** No '<load> Factor' columns found in the header."
        Exit Function
    End If

    MapHeader = 1

End Function

'--------------------------------------------------------------
' Resolve one CSV row into (load set ID, factor) terms.
'--------------------------------------------------------------
Sub ParseRow(r As Integer, lineNo As Integer)

    Dim fi As Integer
    Dim sFac As String
    Dim sCase As String
    Dim fac As Double
    Dim want As String
    Dim li As Integer
    Dim t As Integer
    Dim dupHit As Integer
    Dim desc As String
    Dim isFirst As Integer
    Dim full As String

    rTitle(r) = Trim(fld(colTitle))
    rLine(r) = lineNo
    rOK(r) = 1
    rReason(r) = ""
    rN(r) = 0
    desc = ""

    For fi = 0 To nFam - 1

        If rOK(r) = 1 Then

            sFac = ""
            If famFacCol(fi) < nFld Then
                sFac = Trim(fld(famFacCol(fi)))
            End If

            fac = 0
            If sFac <> "" Then
                If IsNumeric(sFac) Then
                    fac = CDbl(sFac)
                Else
                    rOK(r) = 0
                    rReason(r) = famName(fi) & " Factor is not a number: '" & sFac & "'"
                End If
            End If

            If rOK(r) = 1 And fac <> 0 Then

                want = famName(fi)

                If famCaseCol(fi) >= 0 Then
                    sCase = ""
                    If famCaseCol(fi) < nFld Then
                        sCase = Trim(fld(famCaseCol(fi)))
                    End If
                    If sCase = "" Then
                        rOK(r) = 0
                        rReason(r) = famName(fi) & " Factor is non-zero but " & famName(fi) & " Case is blank"
                    Else
                        want = famName(fi) & " (" & sCase & ")"
                    End If
                End If

                If rOK(r) = 1 Then

                    li = FindLoad(want)

                    If li = -1 Then
                        rOK(r) = 0
                        rReason(r) = "no load set titled '" & want & "' in FEMAP"
                    ElseIf ldDup(li) = 1 Then
                        rOK(r) = 0
                        rReason(r) = "more than one load set is titled '" & want & "' - ambiguous"
                    Else

                        dupHit = 0
                        For t = 0 To rN(r) - 1
                            If rIDs(r, t) = ldID(li) Then
                                dupHit = 1
                            End If
                        Next t

                        If dupHit = 1 Then
                            rOK(r) = 0
                            rReason(r) = "load set " & ldID(li) & " is referenced twice in one row"
                        ElseIf rN(r) > 40 Then
                            rOK(r) = 0
                            rReason(r) = "too many loads in one row"
                        Else
                            isFirst = 0
                            If rN(r) = 0 Then
                                isFirst = 1
                            End If
                            desc = desc & TermText(fac, ldTitle(li), isFirst)
                            rIDs(r, rN(r)) = ldID(li)
                            rFac(r, rN(r)) = fac
                            rTermTitle(r, rN(r)) = ldTitle(li)
                            rN(r) = rN(r) + 1
                        End If

                    End If

                End If

            End If

        End If

    Next fi

    If rOK(r) = 1 Then
        If rN(r) = 0 Then
            rOK(r) = 0
            rReason(r) = "every factor in the row is zero"
        End If
    End If

    If TITLE_WITH_FORMULA = 1 Then
        full = rTitle(r) & " = " & desc
    Else
        full = rTitle(r)
    End If

    If Len(full) > 79 Then
        full = Left(full, 79)
    End If

    rFull(r) = full

End Sub

Function TermText(fac As Double, nm As String, isFirst As Integer) As String

    Dim a As Double
    Dim mag As String

    a = Abs(fac)
    mag = ""
    If a <> 1 Then
        mag = CStr(a) & "*"
    End If

    If isFirst = 1 Then
        If fac < 0 Then
            TermText = "-" & mag & nm
        Else
            TermText = mag & nm
        End If
    Else
        If fac < 0 Then
            TermText = " - " & mag & nm
        Else
            TermText = " + " & mag & nm
        End If
    End If

End Function

'--------------------------------------------------------------
Sub Report

    Dim r As Integer
    Dim t As Integer
    Dim nOK As Integer
    Dim nBad As Integer
    Dim j As Integer
    Dim s As String

    WL "--- load sets found in FEMAP: " & nLd & " ---"
    For j = 0 To nLd - 1
        s = "  " & ldID(j) & "  " & ldTitle(j)
        If ldDup(j) = 1 Then
            s = s & "   <-- title used more than once"
        End If
        WL s
    Next j
    WL ""

    WL "--- combinations ---"

    nOK = 0
    nBad = 0

    For r = 0 To nRows - 1

        If rOK(r) = 1 Then

            nOK = nOK + 1
            WL "  " & (FIRST_ID + r) & "  " & rFull(r)

            For t = 0 To rN(r) - 1
                If rFac(r, t) >= 0 Then
                    s = "+" & CStr(rFac(r, t))
                Else
                    s = CStr(rFac(r, t))
                End If
                WL "          " & s & "  x  set " & rIDs(r, t) & "  " & rTermTitle(r, t)
            Next t

        Else

            nBad = nBad + 1
            WL "  SKIP  (csv line " & rLine(r) & ")  " & rTitle(r) & "  :  " & rReason(r)

        End If

    Next r

    WL ""
    WL "  rows read     : " & nRows
    WL "  resolvable    : " & nOK
    WL "  skipped       : " & nBad

    If nRows > 0 Then
        WL "  target IDs    : " & FIRST_ID & " to " & (FIRST_ID + nRows - 1)
    End If

    WL ""

End Sub

'--------------------------------------------------------------
' abortIfClash = 1 : stop the run if any target ID is occupied
' abortIfClash = 0 : report only
'--------------------------------------------------------------
Sub CheckIds(abortIfClash As Integer)

    Dim oLS As Object
    Dim r As Integer
    Dim clash As Integer

    WL "--- target ID check ---"

    clash = 0
    Set oLS = femap.feLoadSet

    For r = 0 To nRows - 1
        If rOK(r) = 1 Then
            If oLS.Get(FIRST_ID + r) = -1 Then
                clash = clash + 1
                WL "  ID " & (FIRST_ID + r) & " already used by: " & oLS.title
            End If
        End If
    Next r

    If clash = 0 Then
        WL "  all target IDs are free"
        WL ""
        Exit Sub
    End If

    WL "  " & clash & " target IDs are occupied."
    WL "  Delete the old combinations, or raise FIRST_ID."
    WL ""

    If abortIfClash = 1 Then
        WL "  *** Nothing was created."
        Close #logF
        femap.feAppMessage 2, "Target IDs already in use. Delete the old combinations first. See make_combos.log"
        End
    End If

End Sub

'--------------------------------------------------------------
Sub CreateAll

    Dim r As Integer
    Dim made As Integer
    Dim bad As Integer
    Dim stopNow As Integer

    WL "--- creating ---"

    made = 0
    bad = 0
    stopNow = 0

    For r = 0 To nRows - 1

        If rOK(r) = 1 Then

            If stopNow = 0 Then

                If CreateCombo(FIRST_ID + r, r) = 1 Then
                    made = made + 1
                    If RUN_MODE = 1 Then
                        WL "  [RUN_MODE 1] created set " & (FIRST_ID + r) & " only."
                        WL "  Open it in Model > Load > Combine and check the"
                        WL "  loads, the factors and that Set Type is"
                        WL "  Nastran LOAD Combination, then RUN_MODE = 2."
                        stopNow = 1
                    End If
                Else
                    bad = bad + 1
                    If failCount >= MAX_FAILS Then
                        stopNow = 1
                    End If
                End If

            End If

        End If

    Next r

    WL "  created : " & made
    WL "  failed  : " & bad
    WL ""

End Sub

'--------------------------------------------------------------
' Same creation sequence as make_load_combos_v2, which built
' the 24 SCR combinations successfully on this build.
'--------------------------------------------------------------
Function CreateCombo(newID As Long, r As Integer) As Integer

    Dim oLS As Object
    Dim vFac As Variant
    Dim vIDs As Variant
    Dim j As Integer
    Dim cnt As Integer

    CreateCombo = 0
    cnt = rN(r)

    On Error GoTo BailCombo

    ReDim vFac(0 To cnt - 1)
    ReDim vIDs(0 To cnt - 1)

    For j = 0 To cnt - 1
        vFac(j) = CDbl(rFac(r, j))
        vIDs(j) = rIDs(r, j)
    Next j

    Set oLS = femap.feLoadSet                             'VERIFIED

    On Error Resume Next
    oLS.title = rFull(r)
    Err = 0
    oLS.NastranCombination = True
    Err = 0
    oLS.CombinationType = 1
    Err = 0
    oLS.IsNastranCombination = True
    Err = 0
    On Error GoTo BailCombo

    oLS.PutCombination 1#, cnt, vFac, vIDs                'VERIFIED

    oLS.Put newID                                         'VERIFIED

    Set oLS = femap.feLoadSet
    If oLS.Get(newID) <> -1 Then
        Call NoteFail(newID, "load set was not created")
        Exit Function
    End If

    WL "  created " & newID & "  " & oLS.title

    CreateCombo = 1
    failCount = 0
    Exit Function

BailCombo:
    Call NoteFail(newID, "runtime error")

End Function

Sub NoteFail(lsID As Long, reason As String)

    failCount = failCount + 1
    WL "  FAILED set " & lsID & " : " & reason

    If failCount >= MAX_FAILS Then
        WL ""
        WL "  *** ABORTING after " & MAX_FAILS & " consecutive failures."
        WL ""
    End If

End Sub

'==============================================================
' UTILITY
'==============================================================

' Uppercase and drop all spaces, tabs and non-breaking spaces,
' so "Ec (H case 1)" and "Ec (H Case 1)" compare equal.
Function NormKey(s As String) As String

    Dim i As Integer
    Dim c As String
    Dim res As String

    res = ""

    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c <> " " And c <> Chr(9) And c <> Chr(160) Then
            res = res & c
        End If
    Next i

    NormKey = UCase(res)

End Function

' Splits one CSV line into fld(0..nFld-1). Handles quoted fields
' and doubled quotes, as Excel writes them.
Sub SplitLine(s As String)

    Dim i As Integer
    Dim c As String
    Dim cur As String
    Dim inQ As Integer

    nFld = 0
    cur = ""
    inQ = 0
    i = 1

    Do While i <= Len(s)

        c = Mid(s, i, 1)

        If inQ = 1 Then
            If c = Chr(34) Then
                If i < Len(s) Then
                    If Mid(s, i + 1, 1) = Chr(34) Then
                        cur = cur & Chr(34)
                        i = i + 1
                    Else
                        inQ = 0
                    End If
                Else
                    inQ = 0
                End If
            Else
                cur = cur & c
            End If
        Else
            If c = Chr(34) Then
                inQ = 1
            ElseIf c = delim Then
                If nFld <= 60 Then
                    fld(nFld) = cur
                    nFld = nFld + 1
                End If
                cur = ""
            Else
                cur = cur & c
            End If
        End If

        i = i + 1

    Loop

    If nFld <= 60 Then
        fld(nFld) = cur
        nFld = nFld + 1
    End If

End Sub

Function EnsureDir(p As String) As Integer

    Dim i As Integer
    Dim c As String
    Dim partial As String

    EnsureDir = 0
    partial = ""

    For i = 1 To Len(p)
        c = Mid(p, i, 1)
        partial = partial & c
        If c = "\" Then
            If Len(partial) > 3 Then
                On Error Resume Next
                MkDir partial
                On Error GoTo 0
            End If
        End If
    Next i

    On Error Resume Next
    If Dir(p, 16) <> "" Then
        EnsureDir = 1
    End If
    On Error GoTo 0

End Function

Function FileThere(p As String) As Integer

    FileThere = 0
    On Error Resume Next
    If Dir(p) <> "" Then
        FileThere = 1
    End If
    On Error GoTo 0

End Function

Sub WL(s As String)
    Print #logF, s
End Sub
