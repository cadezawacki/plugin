Attribute VB_Name = "EptToolkit"
'===============================================================================
' EptToolkit.bas - convenience wrappers for the ExcelPerfToolkit XLL (EPT.*)
'
' Import: VBE > File > Import File... (or drag this file into a project's
' Modules folder). Requires the ExcelPerfToolkit add-in (the packed .xll) to
' be loaded; call EptIsAvailable to check at runtime.
'
' Three things live here:
'
'   1. FAST-MODE GUARDS - EptFastModeBegin/EptFastModeEnd wrap any macro in
'      the standard suspension of screen updates, automatic calculation,
'      events, and the status bar. Re-entrant: nested Begin/End pairs only
'      restore Excel's state when the outermost pair closes, so wrappers can
'      call other wrappers safely.
'
'   2. AUTO-OPTIMIZATION HOOKS - EptAutoOptimizeOpen/EptAutoOptimizeClose,
'      designed to be called from workbook events. Paste this into
'      ThisWorkbook:
'
'          Private Sub Workbook_Open()
'              EptAutoOptimizeOpen
'          End Sub
'
'          Private Sub Workbook_BeforeClose(Cancel As Boolean)
'              EptAutoOptimizeClose
'          End Sub
'
'   3. ONE-LINE WRAPPERS for the common EPT use cases. Every wrapper follows
'      the toolkit's one-crossing rule: the range is read into a Variant in
'      ONE bulk crossing, transformed in managed memory by the XLL, and
'      written back in ONE bulk crossing - never cell-by-cell.
'===============================================================================
Option Explicit

' ---------------------------------------------------------------------------
' Fast-mode state (module-level so Begin/End can nest).
' ---------------------------------------------------------------------------
Private mFastModeDepth As Long
Private mSavedCalculation As XlCalculation
Private mSavedScreenUpdating As Boolean
Private mSavedEnableEvents As Boolean
Private mSavedDisplayStatusBar As Boolean

' ===========================================================================
' 0. Availability
' ===========================================================================

' True when the ExcelPerfToolkit XLL is loaded and answering calls.
Public Function EptIsAvailable() As Boolean
    Dim v As Variant
    On Error Resume Next
    v = Application.Run("EPT.SIMD.CAPABILITIES")
    EptIsAvailable = (Err.Number = 0) And Not IsError(v)
    Err.Clear
    On Error GoTo 0
End Function

' ===========================================================================
' 1. Fast-mode guards
' ===========================================================================

' Suspend the four big macro slowdowns. Always pair with EptFastModeEnd
' (use an error handler so an exception cannot leave Excel suspended):
'
'   Sub MyMacro()
'       EptFastModeBegin
'       On Error GoTo Cleanup
'       ' ... work ...
'   Cleanup:
'       EptFastModeEnd
'   End Sub
Public Sub EptFastModeBegin()
    If mFastModeDepth = 0 Then
        With Application
            mSavedCalculation = .Calculation
            mSavedScreenUpdating = .ScreenUpdating
            mSavedEnableEvents = .EnableEvents
            mSavedDisplayStatusBar = .DisplayStatusBar
            .Calculation = xlCalculationManual
            .ScreenUpdating = False
            .EnableEvents = False
            .DisplayStatusBar = False
        End With
    End If
    mFastModeDepth = mFastModeDepth + 1
End Sub

' Restore Excel's state saved by the outermost EptFastModeBegin. When the
' saved mode was automatic calculation, a recalc runs so the sheet is
' consistent before control returns to the user (pass False to skip it).
Public Sub EptFastModeEnd(Optional ByVal recalcIfAutomatic As Boolean = True)
    If mFastModeDepth > 0 Then mFastModeDepth = mFastModeDepth - 1
    If mFastModeDepth = 0 Then
        With Application
            .Calculation = mSavedCalculation
            .ScreenUpdating = mSavedScreenUpdating
            .EnableEvents = mSavedEnableEvents
            .DisplayStatusBar = mSavedDisplayStatusBar
        End With
        If recalcIfAutomatic And mSavedCalculation = xlCalculationAutomatic Then
            Application.Calculate
        End If
    End If
End Sub

' ===========================================================================
' 2. Auto-optimization hooks (wire to Workbook_Open / Workbook_BeforeClose)
' ===========================================================================

' One-time, low-risk host tuning at workbook open. Deliberately does NOT
' force manual calculation - that belongs to explicit fast-mode sections.
Public Sub EptAutoOptimizeOpen()
    On Error Resume Next
    With Application
        ' Make sure the multithreaded recalc engine is on with one thread per
        ' core, so every IsThreadSafe EPT.* UDF actually fans out.
        .MultiThreadedCalculation.Enabled = True
        .MultiThreadedCalculation.ThreadMode = xlThreadModeAutomatic
        ' UI animation only costs time during bulk writes.
        .EnableAnimations = False
    End With
    ' Warm the XLL so the first real call doesn't pay assembly-load latency.
    If EptIsAvailable() Then Application.Run "EPT.SIMD.CAPABILITIES"
    Err.Clear
    On Error GoTo 0
End Sub

' Cleanup at workbook close: drop the EPT session cache so memoized blocks
' never outlive the workbook that built them, and restore the UI defaults.
Public Sub EptAutoOptimizeClose()
    On Error Resume Next
    Application.Run "EPT.CACHE.CLEAR"
    Application.EnableAnimations = True
    Err.Clear
    On Error GoTo 0
End Sub

' ===========================================================================
' 3. Bulk-transfer primitives (the one-crossing rule in VBA form)
' ===========================================================================

' Read a whole range as a 2-D Variant in ONE crossing. Single cells are
' boxed into a 1x1 array so callers can always index (1, 1).
Public Function EptReadRange(ByVal source As Range) As Variant
    If source.Cells.CountLarge = 1 Then
        Dim one(1 To 1, 1 To 1) As Variant
        one(1, 1) = source.Value
        EptReadRange = one
    Else
        EptReadRange = source.Value
    End If
End Function

' Write a 2-D array in ONE crossing, sized from the anchor's top-left cell.
Public Sub EptWriteArray(ByVal anchor As Range, ByRef data As Variant)
    Dim rowCount As Long, colCount As Long
    rowCount = UBound(data, 1) - LBound(data, 1) + 1
    colCount = UBound(data, 2) - LBound(data, 2) + 1
    anchor.Cells(1, 1).Resize(rowCount, colCount).Value = data
End Sub

' Read a range, run any same-shape EPT.* block function over it, and write
' the result back in place: two crossings total, regardless of cell count.
'
'   EptTransformRange Range("A1:A100000"), "EPT.TRIMBLOCK"
'   EptTransformRange Range("A1:A100000"), "EPT.SUBSTITUTE", "N/A", ""
Public Sub EptTransformRange(ByVal target As Range, ByVal eptFunction As String, ParamArray extraArgs() As Variant)
    Dim arr As Variant
    arr = EptReadRange(target)
    Select Case UBound(extraArgs) - LBound(extraArgs) + 1
        Case 0: arr = Application.Run(eptFunction, arr)
        Case 1: arr = Application.Run(eptFunction, arr, extraArgs(0))
        Case 2: arr = Application.Run(eptFunction, arr, extraArgs(0), extraArgs(1))
        Case 3: arr = Application.Run(eptFunction, arr, extraArgs(0), extraArgs(1), extraArgs(2))
        Case 4: arr = Application.Run(eptFunction, arr, extraArgs(0), extraArgs(1), extraArgs(2), extraArgs(3))
        Case Else: Err.Raise vbObjectError + 513, "EptTransformRange", "At most 4 extra arguments are supported."
    End Select
    If IsError(arr) Then Err.Raise vbObjectError + 514, "EptTransformRange", eptFunction & " returned an error."
    target.Value = arr
End Sub

' ===========================================================================
' 4. Common use-case wrappers
' ===========================================================================

' Trim/collapse whitespace, then convert numeric-looking text to real
' numbers - the standard two-step cleanup after any paste or import.
Public Sub EptCleanRange(ByVal target As Range)
    EptFastModeBegin
    On Error GoTo Cleanup
    Dim arr As Variant
    arr = EptReadRange(target)
    arr = Application.Run("EPT.TRIMBLOCK", arr)
    arr = Application.Run("EPT.COERCENUMERIC", arr)
    target.Value = arr
Cleanup:
    EptFastModeEnd
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Whole-column lookup in one O(M+R) pass: resolves every cell of
' lookupValues against the FIRST column of tableRange and writes the
' column(s) named by colIndex (1-based; pass Array(2, 5) for several) at
' destination. Exact match by default - note this is the OPPOSITE of native
' VLOOKUP's default, because exact is what almost everyone means.
Public Sub EptVLookupFast(ByVal lookupValues As Range, ByVal tableRange As Range, _
                          ByVal colIndex As Variant, ByVal destination As Range, _
                          Optional ByVal approximateMatch As Boolean = False)
    Dim result As Variant
    result = Application.Run("EPT.VLOOKUPB", EptReadRange(lookupValues), _
                             EptReadRange(tableRange), colIndex, CVErr(xlErrNA), approximateMatch)
    If IsError(result) Then Err.Raise vbObjectError + 515, "EptVLookupFast", "EPT.VLOOKUPB returned an error."
    EptWriteArray destination, result
End Sub

' Apply a whole find/replace pair table to a range in one pass - the bulk
' replacement for chained SUBSTITUTE formulas or repeated Range.Replace.
Public Sub EptSubstituteAll(ByVal target As Range, ByVal findValues As Range, ByVal replaceValues As Range)
    EptTransformRange target, "EPT.SUBSTITUTEALL", EptReadRange(findValues), EptReadRange(replaceValues)
End Sub

' One-pass GROUP BY: distinct keys + one aggregate, written at destination.
' operation: sum|count|average|min|max|median|stdev|... (see EPT.GROUPBY).
Public Sub EptGroupBy(ByVal keys As Range, ByVal values As Range, _
                      ByVal operation As String, ByVal destination As Range)
    Dim result As Variant
    result = Application.Run("EPT.GROUPBY", EptReadRange(keys), EptReadRange(values), operation)
    If IsError(result) Then Err.Raise vbObjectError + 516, "EptGroupBy", "EPT.GROUPBY returned an error."
    EptWriteArray destination, result
End Sub

' Cumulative totals down each column in one O(N) pass, written at
' destination. operation: sum|count|average|min|max|product.
Public Sub EptRunningTotals(ByVal source As Range, ByVal destination As Range, _
                            Optional ByVal operation As String = "sum")
    Dim result As Variant
    result = Application.Run("EPT.RUNNINGAGG", EptReadRange(source), operation)
    If IsError(result) Then Err.Raise vbObjectError + 517, "EptRunningTotals", "EPT.RUNNINGAGG returned an error."
    EptWriteArray destination, result
End Sub

' Stream a delimited file straight into the grid - no Workbooks.Open, no
' object model, one bulk write at destination.
Public Sub EptImportCsv(ByVal filePath As String, ByVal destination As Range, _
                        Optional ByVal delimiter As String = ",")
    Dim arr As Variant
    arr = Application.Run("EPT.READCSV", filePath, delimiter)
    If IsError(arr) Then Err.Raise vbObjectError + 518, "EptImportCsv", "EPT.READCSV failed for: " & filePath
    EptWriteArray destination, arr
End Sub

' Stream a range to a delimited file - one bulk read, direct managed I/O.
Public Sub EptExportCsv(ByVal source As Range, ByVal filePath As String, _
                        Optional ByVal delimiter As String = ",")
    Dim written As Variant
    written = Application.Run("EPT.WRITECSV", filePath, EptReadRange(source), delimiter)
    If IsError(written) Then Err.Raise vbObjectError + 519, "EptExportCsv", "EPT.WRITECSV failed for: " & filePath
End Sub

' Content hash of a range (XXH3, hex) - cheap change detection for caching
' patterns built on EPT.MEMOIZE / EPT.DISKCACHE.*.
Public Function EptRangeHash(ByVal source As Range) As String
    EptRangeHash = Application.Run("EPT.HASHBLOCK", EptReadRange(source))
End Function
