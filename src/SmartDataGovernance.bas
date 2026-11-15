Option Explicit
Option Compare Text

' ================================================================
'  Smart Data Governance — VBA Toolkit
'  Four data-governance modules for spreadsheet data:
'
'    Module 1 — Deduplicate + Merge
'         Same identifier+value → merge row content (CN/EN names, material,
'         batch IDs) with dedup on "/" and "+" delimiters, then delete
'         redundant rows bottom-up.
'
'    Module 2 — Calibrate By Identifier
'         Weighted frequency + recency scoring picks the best identifier
'         code for each product link, cleans common suffixes from links,
'         and auto-updates sibling rows that share the same batch ID +
'         similar product name.
'
'    Module 3 — Clean Text Fields
'         Intelligently deduplicate slash-separated CN names (prefer the
'         more detailed variant), normalize EN name casing (keep the one
'         with more capitals), clean material strings (mixed CN/EN,
'         normalize separators, strip lone "+" signs).
'
'  Author: ivy92888-glitch  |  License: MIT
'  Disclaimer: Unaffiliated with any spreadsheet vendor.
'              Generic clean-room implementation. Any resemblance to
'              specific business systems is coincidental.
' ================================================================

' --- SHEET NAME CONSTANTS ---
Private Const SHEET_LM_DATA     As String = "LM_Data"        ' 历史记录表（Module 1 & 2 用）
Private Const SHEET_REF_DATA    As String = "Reference"       ' 校准目标表（Module 1 & 2 & 3 用）

' --- COLUMN LETTERS ---
' LM_Data sheet columns
Private Const COL_LM_NAME_CN    As String = "A"   ' 中文名
Private Const COL_LM_NAME_EN    As String = "B"   ' 英文名
Private Const COL_LM_CODE       As String = "C"   ' 标识符/编码
Private Const COL_LM_LINK       As String = "H"   ' 链接
Private Const COL_LM_MATERIAL   As String = "I"   ' 材质
Private Const COL_LM_BATCH_ID   As String = "J"   ' 批次标识符

' Reference sheet columns
Private Const COL_REF_BATCH     As String = "A"   ' 批次标识符
Private Const COL_REF_NAME_CN   As String = "B"   ' 中文名
Private Const COL_REF_NAME_EN   As String = "C"   ' 英文名
Private Const COL_REF_CODE      As String = "D"   ' 标识符/编码（Module 2 校准目标）
Private Const COL_REF_LINK      As String = "I"   ' 链接

' --- SUFFIX LIST (strip from links before matching) ---
Private linkSuffixes As Variant
' Populated at runtime: Array("-stocklow", "-outofstock", "-default", "-backup",
'                            "-new", "-slow", "-clearance", "-recommended")

' --- SPECIAL WORDS (keep uppercase when formatting EN names) ---
Private specialWords As Variant
' Populated at runtime: Array("LED", "CD", "DVD", "USB", "HDMI", "CPU", "GPU",
'                            "RAM", "SSD", "HDD", "UK", "US", "EU", "DE", "PL",
'                            "FBA", "IP", "TV")

' --- SPECIAL MATERIALS (keep English word in mixed CN/EN material cells) ---
Private specialMaterials As Variant
' Populated at runtime: Array("PVC", "PP", "PE", "ABS", "LED", "USB", "HDMI")

' --- CALIBRATION WEIGHTS ---
Private Const FREQ_WEIGHT       As Double = 0.6   ' frequency weight (60%)
Private Const RECENCY_WEIGHT    As Double = 0.4   ' recency weight  (40%)

' ================================================================
'  MAIN ENTRY — Run all three modules on Reference sheet
' ================================================================
Sub RunAllGovernance()
    Call InitGovernance
    Application.ScreenUpdating = False

    Dim n1 As Long, n2 As Long, n3 As Long
    n1 = RunModule1_Deduplicate(ThisWorkbook)
    n2 = RunModule2_Calibrate(ThisWorkbook)
    n3 = RunModule3_CleanFields(ThisWorkbook)

    Application.ScreenUpdating = True
    MsgBox "=== Smart Data Governance ===" & vbCrLf & vbCrLf & _
           "Module 1 — Deduplicate+Merge: " & n1 & " rows merged into existing rows" & vbCrLf & _
           "Module 2 — CalibrateByIdentifier: " & n2 & " identifier codes calibrated" & vbCrLf & _
           "Module 3 — CleanFields: " & n3 & " fields cleaned", _
           vbInformation, "Governance Complete"
End Sub

Sub InitGovernance()
    linkSuffixes     = Array("-stocklow", "-outofstock", "-default", "-backup", _
                              "-new", "-slow", "-clearance", "-recommended")
    specialWords     = Array("LED", "CD", "DVD", "USB", "HDMI", "CPU", "GPU", _
                              "RAM", "ROM", "SSD", "HDD", "IP", "TV", _
                              "UK", "US", "EU", "DE", "PL", "FBA")
    specialMaterials = Array("PVC", "PP", "PE", "ABS", "PMMA", "PETG", "TPU", "TPE", _
                              "LED", "USB", "HDMI", "FRP", "GFRP")
End Sub

' ================================================================
'  MODULE 1 — Deduplicate + Merge
'  Works on LM_Data sheet: rows with same link + same identifier →
'  merge CN names, EN names, materials (with internal dedup on / and +),
'  then delete redundant rows bottom-up.
' ================================================================
Function RunModule1_Deduplicate(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(SHEET_LM_DATA)
    On Error GoTo 0
    If ws Is Nothing Then
        RunModule1_Deduplicate = 0: Exit Function
    End If

    Dim lastRow As Long, i As Long, j As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_LM_LINK).End(xlUp).Row
    If lastRow < 2 Then RunModule1_Deduplicate = 0: Exit Function

    Application.ScreenUpdating = False

    ' --- Step 1: Row-internal dedup + format on each column ---
    For i = 2 To lastRow
        ' CN names (col A) — dedup on "/" separator
        ws.Cells(i, COL_LM_NAME_CN).Value = _
            DedupAndJoin(ws.Cells(i, COL_LM_NAME_CN).Value, "/", False)

        ' EN names (col B) — dedup on "/" + format
        ws.Cells(i, COL_LM_NAME_EN).Value = _
            DedupAndFormatEN(ws.Cells(i, COL_LM_NAME_EN).Value)

        ' Materials (col I) — dedup on "+" separator
        ws.Cells(i, COL_LM_MATERIAL).Value = _
            DedupAndJoin(ws.Cells(i, COL_LM_MATERIAL).Value, "+", False)
    Next i

    ' --- Step 2: Build key groupings ---
    ' key = link & "|" & identifier → array of row numbers
    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")
    Dim key As String
    For i = 2 To lastRow
        Dim link As String, id As String
        link = Trim(ws.Cells(i, COL_LM_LINK).Value)
        id   = Trim(ws.Cells(i, COL_LM_CODE).Value)
        If Len(link) = 0 Or Len(id) = 0 Then GoTo M1_Next
        key = link & "|" & id
        If Not groups.Exists(key) Then
            groups.Add key, CreateObject("Scripting.Dictionary")
        End If
        groups(key)(i) = True   ' use inner dict as a set
M1_Next:
    Next i

    ' --- Step 3: For each group — merge content into one row ---
    Dim rowsToDelete As Object
    Set rowsToDelete = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    For Each k In groups.Keys
        Dim rowSet As Object
        Set rowSet = groups(k)
        If rowSet.Count <= 1 Then GoTo GroupDone

        ' Find max row num → keep this row, merge content from others into it
        Dim keepRow As Long, rn As Variant
        keepRow = 0
        For Each rn In rowSet.Keys
            If CLng(rn) > keepRow Then keepRow = CLng(rn)
        Next rn

        Dim mergedCN As String, mergedEN As String, mergedMat As String, mergedBatch As String
        Dim first As Boolean: first = True
        For Each rn In rowSet.Keys
            Dim r As Long: r = CLng(rn)
            If r = keepRow Then GoTo MergeNext
            If first Then
                mergedCN   = Trim(ws.Cells(r, COL_LM_NAME_CN).Value)
                mergedEN   = Trim(ws.Cells(r, COL_LM_NAME_EN).Value)
                mergedMat  = Trim(ws.Cells(r, COL_LM_MATERIAL).Value)
                mergedBatch = Trim(ws.Cells(r, COL_LM_BATCH_ID).Value)
                first = False
            Else
                mergedCN   = MergeValues(mergedCN,   Trim(ws.Cells(r, COL_LM_NAME_CN).Value), "/", True)
                mergedEN   = MergeValues(mergedEN,   Trim(ws.Cells(r, COL_LM_NAME_EN).Value), "/", False)
                mergedMat  = MergeValues(mergedMat,  Trim(ws.Cells(r, COL_LM_MATERIAL).Value), "+", True)
                mergedBatch = MergeValues(mergedBatch, Trim(ws.Cells(r, COL_LM_BATCH_ID).Value), "/", True)
            End If
            rowsToDelete(r) = True
MergeNext:
        Next rn

        ' Merge keep-row's own values into merged string, then write back
        If Not first Then
            mergedCN   = MergeValues(mergedCN,   Trim(ws.Cells(keepRow, COL_LM_NAME_CN).Value), "/", True)
            mergedEN   = MergeValues(mergedEN,   Trim(ws.Cells(keepRow, COL_LM_NAME_EN).Value), "/", False)
            mergedMat  = MergeValues(mergedMat,  Trim(ws.Cells(keepRow, COL_LM_MATERIAL).Value), "+", True)
            mergedBatch = MergeValues(mergedBatch, Trim(ws.Cells(keepRow, COL_LM_BATCH_ID).Value), "/", True)

            ws.Cells(keepRow, COL_LM_NAME_CN).Value   = mergedCN
            ws.Cells(keepRow, COL_LM_NAME_EN).Value   = mergedEN
            ws.Cells(keepRow, COL_LM_MATERIAL).Value  = mergedMat
            ws.Cells(keepRow, COL_LM_BATCH_ID).Value  = mergedBatch
        End If

GroupDone:
    Next k

    ' --- Step 4: Delete redundant rows bottom-up ---
    Dim delCount As Long
    delCount = 0
    Dim sorted() As Long: ReDim sorted(rowsToDelete.Count - 1)
    Dim si As Long: si = 0
    For Each rn In rowsToDelete.Keys
        sorted(si) = CLng(rn): si = si + 1
    Next rn
    QuickSort sorted, 0, UBound(sorted)

    For i = UBound(sorted) To 0 Step -1
        ws.Rows(sorted(i)).Delete
        delCount = delCount + 1
    Next i

    Application.ScreenUpdating = True
    RunModule1_Deduplicate = delCount
End Function

' ================================================================
'  MODULE 2 — Calibrate Identifier By Weighted Scoring
'  Reads LM_Data, for each stripped-back link builds a frequency+recency
'  score over all observed identifiers. Picks the highest-scoring one,
'  writes it to Reference sheet. Also auto-updates sibling rows that
'  share the same batch-ID suffix pattern + name inclusion rule.
' ================================================================
Function RunModule2_Calibrate(ByVal wb As Workbook) As Long
    Dim wsLM As Worksheet, wsRef As Worksheet
    On Error Resume Next
    Set wsLM  = wb.Sheets(SHEET_LM_DATA)
    Set wsRef = wb.Sheets(SHEET_REF_DATA)
    On Error GoTo 0
    If wsLM Is Nothing Or wsRef Is Nothing Then
        RunModule2_Calibrate = 0: Exit Function
    End If

    Dim lastRowLM As Long, lastRowRef As Long
    lastRowLM  = wsLM.Cells(wsLM.Rows.Count, COL_LM_LINK).End(xlUp).Row
    lastRowRef = wsRef.Cells(wsRef.Rows.Count, COL_REF_LINK).End(xlUp).Row

    Application.ScreenUpdating = False

    ' --- Step 1: Build score table from LM_Data ---
    ' key = strippedBackLink & "|" & identifier → Array(freq, rowSum)
    Dim scoreTable As Object
    Set scoreTable = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 2 To lastRowLM
        Dim link As String, id As String
        link = StripSuffix(Trim(wsLM.Cells(i, COL_LM_LINK).Value))
        id   = Trim(wsLM.Cells(i, COL_LM_CODE).Value)
        If Len(link) = 0 Or Len(id) = 0 Then GoTo M2_Next

        Dim sk As String: sk = link & "|" & id
        If Not scoreTable.Exists(sk) Then
            scoreTable.Add sk, Array(0, 0)
        End If
        scoreTable(sk)(0) = scoreTable(sk)(0) + 1           ' frequency
        scoreTable(sk)(1) = scoreTable(sk)(1) + i            ' row sum (recency proxy)
M2_Next:
    Next i

    ' --- Step 2: Aggregate per stripped-back link ---
    ' linkKey → Array of (identifier, score) pairs
    Dim linkCandidates As Object
    Set linkCandidates = CreateObject("Scripting.Dictionary")

    ' Find max frequency and max row-sum for normalization
    Dim maxFreq As Long, maxRowSum As Double
    maxFreq = 0: maxRowSum = 0
    Dim sk As Variant
    For Each sk In scoreTable.Keys
        Dim parts As Variant
        parts = Split(CStr(sk), "|")
        Dim lk As String: lk = parts(0)
        Dim pair As Variant: pair = scoreTable(sk)
        If pair(0) > maxFreq Then maxFreq = pair(0)
        If pair(1) > maxRowSum Then maxRowSum = pair(1)
    Next sk

    ' Compute weighted score
    For Each sk In scoreTable.Keys
        parts = Split(CStr(sk), "|")
        lk = parts(0)
        Dim iden As String: iden = parts(1)
        pair = scoreTable(sk)

        Dim freqScore As Double:  freqScore = (pair(0) / maxFreq) * 100
        Dim recencyScore As Double: recencyScore = (pair(1) / maxRowSum) * 100
        Dim totalScore As Double: totalScore = freqScore * FREQ_WEIGHT + recencyScore * RECENCY_WEIGHT

        If Not linkCandidates.Exists(lk) Then
            linkCandidates.Add lk, CreateObject("Scripting.Dictionary")
        End If
        If Not linkCandidates(lk).Exists(iden) Then
            linkCandidates(lk)(iden) = 0
        End If
        ' Keep max score per identifier (in case same identifier appears multiple times)
        If totalScore > linkCandidates(lk)(iden) Then
            linkCandidates(lk)(iden) = totalScore
        End If
    Next sk

    ' --- Step 3: For each link, pick the best-scoring identifier ---
    Dim bestCodeMap As Object   ' stripped link → best identifier
    Set bestCodeMap = CreateObject("Scripting.Dictionary")
    Dim lk As Variant
    For Each lk In linkCandidates.Keys
        Dim bestId As String, bestScore As Double
        bestId = "": bestScore = -1
        Dim cid As Variant
        For Each cid In linkCandidates(lk).Keys
            If linkCandidates(lk)(cid) > bestScore Then
                bestScore = linkCandidates(lk)(cid)
                bestId = CStr(cid)
            End If
        Next cid
        If Len(bestId) > 0 Then bestCodeMap(lk) = bestId
    Next lk

    ' --- Step 4: Calibrate Reference sheet ---
    Dim calibrated As Long
    calibrated = 0
    For i = 2 To lastRowRef
        Dim refLink As String, refCode As String
        refLink = StripSuffix(Trim(wsRef.Cells(i, COL_REF_LINK).Value))
        refCode = Trim(wsRef.Cells(i, COL_REF_CODE).Value)

        If bestCodeMap.Exists(refLink) Then
            Dim targetCode As String
            targetCode = bestCodeMap(refLink)
            If targetCode <> refCode Then
                wsRef.Cells(i, COL_REF_CODE).Value = targetCode
                calibrated = calibrated + 1

                ' --- Step 4a: Auto-update sibling rows ---
                Call UpdateSiblings(wsRef, i, targetCode, refCode)
            End If
        End If
    Next i

    Application.ScreenUpdating = True
    RunModule2_Calibrate = calibrated
End Function

' --- Sibling row updater ---
' Batch-ID suffix rule: current row batch = sibling batch + "备用" (or local suffix)
' Name-inclusion rule: sibling CN name includes current CN name
Sub UpdateSiblings(ByVal wsRef As Worksheet, ByVal anchorRow As Long, _
                   ByVal newCode As String, ByVal oldCode As String)
    Dim anchorBatch As String, anchorName As String
    anchorBatch = Trim(wsRef.Cells(anchorRow, COL_REF_BATCH).Value)
    anchorName  = Trim(wsRef.Cells(anchorRow, COL_REF_NAME_CN).Value)
    If Len(anchorBatch) = 0 Or Len(anchorName) = 0 Then Exit Sub

    Dim lastRow As Long
    lastRow = wsRef.Cells(wsRef.Rows.Count, COL_REF_BATCH).End(xlUp).Row

    Dim j As Long
    For j = 2 To lastRow
        If j = anchorRow Then GoTo SibNext
        Dim sibBatch As String, sibName As String, sibCode As String
        sibBatch = Trim(wsRef.Cells(j, COL_REF_BATCH).Value)
        sibName  = Trim(wsRef.Cells(j, COL_REF_NAME_CN).Value)
        sibCode  = Trim(wsRef.Cells(j, COL_REF_CODE).Value)

        ' Check suffix rule + name inclusion
        If (sibBatch = anchorBatch & "备用" Or sibBatch & "备用" = anchorBatch) Then
            If InStr(1, sibName, anchorName, vbTextCompare) > 0 _
               Or InStr(1, anchorName, sibName, vbTextCompare) > 0 Then
                If sibCode = oldCode Or Len(sibCode) = 0 Then
                    wsRef.Cells(j, COL_REF_CODE).Value = newCode
                End If
            End If
        End If
SibNext:
    Next j
End Sub

' --- Strip suffixes from a URL ---
Function StripSuffix(ByVal url As String) As String
    Dim s As String: s = url
    Dim suf As Variant
    For Each suf In linkSuffixes
        If InStr(1, s, suf, vbTextCompare) > 0 Then
            s = Replace(s, suf, "", 1, -1, vbTextCompare)
            Exit For   ' strip first matching suffix
        End If
    Next suf
    StripSuffix = s
End Function

' --- Weighted scoring helper ---
Function ScoreWeighted(ByVal freq As Long, ByVal rowSum As Double, _
                        ByVal maxFreq As Long, ByVal maxRowSum As Double) As Double
    If maxFreq = 0 Or maxRowSum = 0 Then ScoreWeighted = 0: Exit Function
    Dim fs As Double: fs = (freq / maxFreq) * 100
    Dim rs As Double: rs = (rowSum / maxRowSum) * 100
    ScoreWeighted = fs * FREQ_WEIGHT + rs * RECENCY_WEIGHT
End Function

' ================================================================
'  MODULE 3 — Clean Text Fields
'  Cleans CN names (prefer detailed "-" variant), EN names
'  (keep variant with most capital letters), and material strings
'  (mixed CN/EN handling, normalize separators, strip lone "+").
' ================================================================
Function RunModule3_CleanFields(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(SHEET_REF_DATA)
    On Error GoTo 0
    If ws Is Nothing Then RunModule3_CleanFields = 0: Exit Function

    Dim lastRow As Long, i As Long, cleaned As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_REF_NAME_EN).End(xlUp).Row
    Application.ScreenUpdating = False

    For i = 2 To lastRow
        ' --- CN names (col B) ---
        Dim rawCN As String, cleanCN As String
        rawCN = Trim(ws.Cells(i, COL_REF_NAME_CN).Value)
        If Len(rawCN) > 0 Then
            cleanCN = CleanCNNames(rawCN)
            If cleanCN <> rawCN Then
                ws.Cells(i, COL_REF_NAME_CN).Value = cleanCN
                cleaned = cleaned + 1
            End If
        End If

        ' --- EN names (col C) ---
        Dim rawEN As String, cleanEN As String
        rawEN = Trim(ws.Cells(i, COL_REF_NAME_EN).Value)
        If Len(rawEN) > 0 Then
            cleanEN = CleanENNames(rawEN)
            If cleanEN <> rawEN Then
                ws.Cells(i, COL_REF_NAME_EN).Value = cleanEN
                cleaned = cleaned + 1
            End If
        End If

        ' --- Materials (col J — adjust to your layout) ---
        Dim rawMat As String, cleanMat As String
        rawMat = Trim(ws.Cells(i, "J").Value)
        If Len(rawMat) > 0 Then
            cleanMat = CleanMaterial(rawMat)
            If cleanMat <> rawMat Then
                ws.Cells(i, "J").Value = cleanMat
                cleaned = cleaned + 1
            End If
        End If
    Next i

    Application.ScreenUpdating = True
    RunModule3_CleanFields = cleaned
End Function

' --- Clean CN names: dedup + prefer more detailed (has "-") ---
Function CleanCNNames(ByVal raw As String) As String
    Dim parts As Variant
    parts = Split(raw, "/")
    If UBound(parts) < 1 Then CleanCNNames = Trim(raw): Exit Function

    Dim chosen() As String
    ReDim chosen(0 To UBound(parts))
    Dim ci As Long: ci = 0
    Dim j As Long

    For j = LBound(parts) To UBound(parts)
        Dim p As String: p = Trim(parts(j))
        If Len(p) = 0 Then GoTo CN_Next

        ' Skip if base-name already has a more detailed variant
        Dim hasBetter As Boolean: hasBetter = False
        Dim k As Long
        For k = LBound(parts) To UBound(parts)
            If k = j Then GoTo CN_InnerNext
            Dim p2 As String: p2 = Trim(parts(k))
            If Len(p2) > 0 And InStr(p2, "-") > 0 Then
                ' p has no "-" but p2 has "-" and p is contained in p2 → p2 is better
                If InStr(p2, p) > 0 And InStr(p, "-") = 0 Then
                    hasBetter = True: Exit For
                End If
            End If
CN_InnerNext:
        Next k
        If hasBetter Then GoTo CN_Next

        chosen(ci) = p: ci = ci + 1
CN_Next:
    Next j

    If ci = 0 Then
        CleanCNNames = Trim(raw)
    Else
        ReDim Preserve chosen(0 To ci - 1)
        ' Dedup (case-insensitive substring)
        CleanCNNames = JoinUniqueStrings(chosen, "/")
    End If
End Function

' --- Clean EN names: keep variant with most capital letters, dedup, format ---
Function CleanENNames(ByVal raw As String) As String
    Dim parts As Variant
    parts = Split(raw, "/")
    If UBound(parts) < 1 Then
        CleanENNames = FormatENName(raw): Exit Function
    End If

    ' For each, count capitals after lowercasing — pick one with most capitals
    Dim bestIdx As Long, bestCapCount As Long
    bestIdx = -1: bestCapCount = -1

    Dim j As Long
    For j = LBound(parts) To UBound(parts)
        Dim p As String: p = Trim(parts(j))
        If Len(p) = 0 Then GoTo EN_Next
        Dim capCount As Long: capCount = CountCapitals(p)
        If capCount > bestCapCount Then
            bestCapCount = capCount: bestIdx = j
        End If
EN_Next:
    Next j

    If bestIdx >= 0 Then
        CleanENNames = FormatENName(Trim(parts(bestIdx)))
    Else
        CleanENNames = FormatENName(Trim(raw))
    End If
End Function

Function CountCapitals(ByVal s As String) As Long
    Dim c As Long, i As Long
    For i = 1 To Len(s)
        Dim ch As String: ch = Mid(s, i, 1)
        If ch >= "A" And ch <= "Z" Then c = c + 1
    Next i
    CountCapitals = c
End Function

' --- Clean material strings: mixed CN/EN handling + separator normalize + strip lone "+" ---
Function CleanMaterial(ByVal raw As String) As String
    Dim s As String: s = raw

    ' Step 1: Normalize separators (Chinese "加", space → "+")
    s = Replace(s, "加", "+")
    s = Replace(s, " ", "+")

    ' Step 2: Split on "/", process each part
    Dim parts As Variant
    parts = Split(s, "/")
    Dim cleaned() As String
    ReDim cleaned(0 To UBound(parts))
    Dim ci As Long: ci = 0
    Dim j As Long

    For j = LBound(parts) To UBound(parts)
        Dim p As String: p = Trim(parts(j))
        If Len(p) = 0 Then GoTo Mat_Next

        ' Step 2a: Handle CN+EN mixed — keep EN if special material, else strip EN
        p = CleanMixedMaterial(p)

        ' Step 2b: Split on "+", dedup, rejoin, strip lone "+"
        p = DedupAndJoin(p, "+", True)
        If Len(p) = 0 Then GoTo Mat_Next

        ' Step 2c: Remove lone leading/trailing "+"
        Do While Left(p, 1) = "+": p = Mid(p, 2): Loop
        Do While Right(p, 1) = "+": p = Left(p, Len(p) - 1): Loop

        If Len(p) > 0 Then
            cleaned(ci) = p: ci = ci + 1
        End If
Mat_Next:
    Next j

    If ci = 0 Then
        CleanMaterial = Trim(raw)
    Else
        ReDim Preserve cleaned(0 To ci - 1)
        CleanMaterial = Join(cleaned, "/")
    End If
End Function

Function CleanMixedMaterial(ByVal s As String) As String
    ' Split on "+" for mixed CN+EN
    Dim tokens As Variant
    tokens = Split(s, "+")
    Dim keep() As String
    ReDim keep(0 To UBound(tokens))
    Dim ci As Long: ci = 0
    Dim j As Long

    For j = LBound(tokens) To UBound(tokens)
        Dim t As String: t = Trim(tokens(j))
        If Len(t) = 0 Then GoTo MM_Next

        ' Check if this token has English letters
        Dim hasEN As Boolean: hasEN = False
        Dim k As Long
        For k = 1 To Len(t)
            Dim ch As String: ch = Mid(t, k, 1)
            If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
                hasEN = True: Exit For
            End If
        Next k

        If hasEN Then
            ' Keep only if it's a special material (PVC, PP, etc.)
            Dim isSpecial As Boolean: isSpecial = False
            Dim sm As Variant
            For Each sm In specialMaterials
                If InStr(1, t, CStr(sm), vbTextCompare) > 0 Then
                    isSpecial = True: Exit For
                End If
            Next sm
            If isSpecial Then
                keep(ci) = t: ci = ci + 1
            End If
        Else
            ' Pure Chinese material — keep as-is
            keep(ci) = t: ci = ci + 1
        End If
MM_Next:
    Next j

    If ci = 0 Then
        CleanMixedMaterial = Trim(s)
    Else
        ReDim Preserve keep(0 To ci - 1)
        CleanMixedMaterial = Join(keep, "+")
    End If
End Function

' ================================================================
'  SHARED UTILITIES
' ================================================================

' --- Dedup strings split by separator, rejoin ---
Function DedupAndJoin(ByVal raw As String, ByVal sep As String, _
                       ByVal preserveCase As Boolean) As String
    If Len(Trim(raw)) = 0 Then DedupAndJoin = "": Exit Function
    Dim parts As Variant
    parts = Split(raw, sep)
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim out() As String
    ReDim out(0 To UBound(parts))
    Dim ci As Long: ci = 0
    Dim j As Long
    For j = LBound(parts) To UBound(parts)
        Dim p As String: p = Trim(parts(j))
        If Len(p) = 0 Then GoTo DJ_Next
        Dim key As String
        If preserveCase Then key = p Else key = LCase(p)
        If Not seen.Exists(key) Then
            seen.Add key, True
            out(ci) = p: ci = ci + 1
        End If
DJ_Next:
    Next j
    If ci = 0 Then DedupAndJoin = "": Exit Function
    ReDim Preserve out(0 To ci - 1)
    DedupAndJoin = Join(out, sep)
End Function

' --- Dedup EN names AND format (first-letter cap, special words all caps) ---
Function DedupAndFormatEN(ByVal raw As String) As String
    If Len(Trim(raw)) = 0 Then DedupAndFormatEN = "": Exit Function
    Dim parts As Variant
    parts = Split(raw, "/")
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim out() As String
    ReDim out(0 To UBound(parts))
    Dim ci As Long: ci = 0
    Dim j As Long
    For j = LBound(parts) To UBound(parts)
        Dim p As String: p = FormatENName(Trim(parts(j)))
        If Len(p) = 0 Then GoTo DF_Next
        If Not seen.Exists(LCase(p)) Then
            seen.Add LCase(p), True
            out(ci) = p: ci = ci + 1
        End If
DF_Next:
    Next j
    If ci = 0 Then DedupAndFormatEN = "": Exit Function
    ReDim Preserve out(0 To ci - 1)
    DedupAndFormatEN = Join(out, "/")
End Function

' --- Format one EN name: first letter cap per word, special words all caps ---
Function FormatENName(ByVal s As String) As String
    If Len(Trim(s)) = 0 Then FormatENName = "": Exit Function
    ' Split on space, format each token
    Dim tokens As Variant
    tokens = Split(Trim(s), " ")
    Dim out() As String
    ReDim out(0 To UBound(tokens))
    Dim j As Long
    For j = LBound(tokens) To UBound(tokens)
        Dim t As String: t = Trim(tokens(j))
        If Len(t) = 0 Then GoTo FN_Next

        ' Check special words first (case-insensitive)
        Dim isSpecial As Boolean: isSpecial = False
        Dim sw As Variant
        For Each sw In specialWords
            If UCase(CStr(sw)) = UCase(t) Then
                t = UCase(CStr(sw))
                isSpecial = True: Exit For
            End If
        Next sw

        If Not isSpecial Then
            ' First letter cap, rest lowercase
            If Len(t) > 1 Then
                t = UCase(Left(t, 1)) & LCase(Mid(t, 2))
            Else
                t = UCase(t)
            End If
        End If
        out(j) = t
FN_Next:
    Next j
    FormatENName = Join(out, " ")
End Function

' --- Merge two values with dedup separator-aware ---
Function MergeValues(ByVal a As String, ByVal b As String, ByVal sep As String, _
                      ByVal preserveCase As Boolean) As String
    a = Trim(a): b = Trim(b)
    If Len(a) = 0 Then MergeValues = b: Exit Function
    If Len(b) = 0 Then MergeValues = a: Exit Function
    MergeValues = DedupAndJoin(a & sep & b, sep, preserveCase)
End Function

' --- Join an array of strings with dedup (substring-safe) ---
Function JoinUniqueStrings(ByVal arr() As String, ByVal sep As String) As String
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim out() As String
    ReDim out(0 To UBound(arr))
    Dim ci As Long: ci = 0
    Dim j As Long
    For j = LBound(arr) To UBound(arr)
        Dim p As String: p = Trim(arr(j))
        If Len(p) = 0 Then GoTo JU_Next
        If Not seen.Exists(LCase(p)) Then
            seen.Add LCase(p), True
            out(ci) = p: ci = ci + 1
        End If
JU_Next:
    Next j
    If ci = 0 Then JoinUniqueStrings = "": Exit Function
    ReDim Preserve out(0 To ci - 1)
    JoinUniqueStrings = Join(out, sep)
End Function

' --- QuickSort helper ---
Private Sub QuickSort(ByRef arr() As Long, ByVal first As Long, ByVal last As Long)
    If first >= last Then Exit Sub
    Dim pivot As Long: pivot = arr((first + last) \ 2)
    Dim i As Long: i = first
    Dim j As Long: j = last
    Dim tmp As Long
    Do While i <= j
        Do While arr(i) > pivot: i = i + 1: Loop
        Do While arr(j) < pivot: j = j - 1: Loop
        If i <= j Then
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    QuickSort arr, first, j
    QuickSort arr, i, last
End Sub
