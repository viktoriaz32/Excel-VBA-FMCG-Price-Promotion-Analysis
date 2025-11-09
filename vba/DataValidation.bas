Attribute VB_Name = "DataValidation"
Option Explicit

' ============================================================
' Excel_VBA_Price_Promotion_Analysis
' Data Validation Module – validates sheets 1–8 + writes report
' Run: RunDataValidation
' ============================================================

' === Config: Sheet Names ===
Private Const SH_SALES As String = "Sales"
Private Const SH_PRODUCTS As String = "Products"
Private Const SH_STORES As String = "Stores"
Private Const SH_CAL As String = "Calendar"
Private Const SH_PROMOS As String = "Promos"
Private Const SH_PRICELIST As String = "Pricelist"
Private Const SH_COMP As String = "Competitor"
Private Const SH_MEDIA As String = "Media"
Private Const SH_VALIDATION As String = "DataValidation"   ' output log

' === Styling (colors) ===
Private Const CLR_ERR As Long = 13434879   ' light red
Private Const CLR_WARN As Long = 10092543  ' light yellow

' === Types / Severities ===
Private Enum Sev
    svERROR = 1
    svWARNING = 2
End Enum

' === Entry point ===
Public Sub RunDataValidation()
    Dim t0 As Double: t0 = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ClearOldHighlights
    PrepareValidationSheet
    
    ' Build lookup indexes for referential checks
    Dim dictSKU As Object, dictStore As Object, dictYearWeek As Object, dictWeekStart As Object
    Set dictSKU = BuildIndexDict(SH_PRODUCTS, "SKU")
    Set dictStore = BuildIndexDict(SH_STORES, "StoreID")
    Set dictYearWeek = BuildIndexDict(SH_CAL, "YearWeek")
    Set dictWeekStart = BuildIndexDictDate(SH_CAL, "WeekStart")
    
    ' Validate each sheet
    ValidateProducts
    ValidateStores
    ValidateCalendar
    ValidateSales dictSKU, dictStore, dictYearWeek, dictWeekStart
    ValidatePromos dictSKU, dictStore, dictYearWeek, dictWeekStart
    ValidatePricelist dictSKU, dictStore, dictYearWeek
    ValidateCompetitor dictYearWeek
    ValidateMedia dictYearWeek
    
    ' Autoformat the log
    FinalizeValidationSheet
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Dim msg As String
    msg = "Data validation finished in " & Format(Timer - t0, "0.00") & "s." & vbCrLf & _
          "See the 'DataValidation' sheet for details. Invalid cells are highlighted."
    MsgBox msg, vbInformation, "Validation complete"
End Sub

' ============================================================
' Sheet-specific validators
' ============================================================

Private Sub ValidateSales(ByVal dictSKU As Object, ByVal dictStore As Object, _
                          ByVal dictYW As Object, ByVal dictWS As Object)
    Dim sh As Worksheet: Set sh = SheetByName(SH_SALES)
    If sh Is Nothing Then Exit Sub
    
    Dim needCols As Variant
    needCols = Array("YearWeek", "WeekStart", "StoreID", "SKU", "Units", "NetPrice_LCU", _
                     "NetRevenue_LCU", "PromoFlag", "FeatureDisplayFlag", _
                     "OnInvoiceDiscount_Pct", "OffInvoiceRebate_Pct", "Returns_Units")
    Dim map As Object: Set map = HeaderMap(sh, needCols)
    If map Is Nothing Then Exit Sub
    
    Dim lastRow As Long: lastRow = LastDataRow(sh)
    Dim r As Long
    For r = 2 To lastRow
        ' Required presence
        Req sh, r, map("YearWeek")
        Req sh, r, map("WeekStart")
        Req sh, r, map("StoreID")
        Req sh, r, map("SKU")
        
        ' Types & ranges
        IsDateCell sh, r, map("WeekStart")
        IsNonNegInt sh, r, map("Units")
        IsNonNeg sh, r, map("NetPrice_LCU")
        IsNonNeg sh, r, map("NetRevenue_LCU")
        IsFlag01 sh, r, map("PromoFlag")
        IsFlag01 sh, r, map("FeatureDisplayFlag")
        IsPct01 sh, r, map("OnInvoiceDiscount_Pct")
        IsPct01 sh, r, map("OffInvoiceRebate_Pct")
        IsNonNegInt sh, r, map("Returns_Units")
        
        ' Foreign keys
        FKExists sh, r, map("SKU"), dictSKU, "SKU not found in Products."
        FKExists sh, r, map("StoreID"), dictStore, "StoreID not found in Stores."
        FKExists sh, r, map("YearWeek"), dictYW, "YearWeek not found in Calendar."
        FKExistsDate sh, r, map("WeekStart"), dictWS, "WeekStart not found in Calendar."
    Next r
End Sub

Private Sub ValidateProducts()
    Dim sh As Worksheet: Set sh = SheetByName(SH_PRODUCTS)
    If sh Is Nothing Then Exit Sub
    
    Dim needCols As Variant
    needCols = Array("SKU", "Brand", "Category", "Segment", "PackSize_ml", "UnitsPerCase", _
                     "LaunchDate", "Status", "StdUnitCost_LCU")
    Dim map As Object: Set map = HeaderMap(sh, needCols)
    If map Is Nothing Then Exit Sub
    
    Dim lastRow As Long: lastRow = LastDataRow(sh)
    
    ' Uniqueness of SKU
    UniqueCheck sh, "SKU", map("SKU")
    
    Dim r As Long
    For r = 2 To lastRow
        Req sh, r, map("SKU")
        Req sh, r, map("Brand")
        Req sh, r, map("Category")
        Req sh, r, map("Segment")
        IsPos sh, r, map("PackSize_ml")
        IsPosInt sh, r, map("UnitsPerCase")
        IsDateCell sh, r, map("LaunchDate")
        Req sh, r, map("Status")  ' (Optionally enforce set e.g., Active/Delisted/Dormant)
    Next r
End Sub

Private Sub ValidateStores()
    Dim sh As Worksheet: Set sh = SheetByName(SH_STORES)
    If sh Is Nothing Then Exit Sub
    
    Dim needCols As Variant
    needCols = Array("StoreID", "Retailer", "Channel", "Region", "Format")
    Dim map As Object: Set map = HeaderMap(sh, needCols)
    If map Is Nothing Then Exit Sub
    
    Dim lastRow As Long: lastRow = LastDataRow(sh)
    UniqueCheck sh, "StoreID", map("StoreID")
    
    Dim r As Long
    For r = 2 To lastRow
        Req sh, r, map("StoreID")
        Req sh, r, map("Retailer")
        Req sh, r, map("Channel")
        Req sh, r, map("Region")
        Req sh, r, map("Format")
    Next r
End Sub

Private Sub ValidateCalendar()
    Dim sh As Worksheet: Set sh = SheetByName(SH_CAL)
    If sh Is Nothing Then Exit Sub
    
    Dim needCols As Variant
    needCols = Array("YearWeek", "WeekStart", "WeekEnd", "Month", "Quarter", _
                     "FiscalPeriod", "HolidayFlag", "Season", "ISOWeek", "ISOYear")
    Dim map As Object: Set map = HeaderMap(sh, needCols)
    If map Is Nothing Then Exit Sub
    
    Dim lastRow As Long: lastRow = LastDataRow(sh)
    UniqueCheck sh, "YearWeek", map("YearWeek")
    
    Dim r As Long
    For r = 2 To lastRow
        Req sh, r, map("YearWeek")
        IsDateCell sh, r, map("WeekStart")
        IsDateCell sh, r, map("WeekEnd")
        If IsDateVal(sh.Cells(r, map("WeekStart"))) And IsDateVal(sh.Cells(r, map("WeekEnd"))) Then
            If sh.Cells(r, map("WeekEnd")).Value < sh.Cells(r, map("WeekStart")).Value Then
                LogIssue sh, r, map("WeekEnd"), svERROR, "WeekEnd < WeekStart."
            End If
        End If
        
        InRangeInt sh, r, map("Month"), 1, 12, "Month must be 1–12."
        Dim qv As Variant: qv = QuarterToInt(sh.Cells(r, map("Quarter")).Value)
If IsEmpty(qv) Then
    LogIssue sh, r, map("Quarter"), svERROR, "Quarter not recognized (accepts Q1..Q4 or 1–4)."
End If
        Req sh, r, map("FiscalPeriod")
        IsFlag01 sh, r, map("HolidayFlag")
        Req sh, r, map("Season")
        InRangeInt sh, r, map("ISOWeek"), 1, 53, "ISOWeek must be 1–53."
        IsNonNegInt sh, r, map("ISOYear")
    Next r
End Sub

Private Sub ValidatePromos(ByVal dictSKU As Object, ByVal dictStore As Object, _
                           ByVal dictYW As Object, ByVal dictWS As Object)
    Dim sh As Worksheet: Set sh = SheetByName(SH_PROMOS)
    If sh Is Nothing Then Exit Sub
    
    Dim needCols As Variant
    needCols = Array("PromoID", "SKU", "StoreID", "WeekStart", "WeekEnd", "Mechanic", _
                     "Depth_Pct", "FeatureDisplayFlag", "CoopFunding_LCU", "Comments")
    Dim map As Object: Set map = HeaderMap(sh, needCols)
    If map Is Nothing Then Exit Sub
    
    Dim lastRow As Long: lastRow = LastDataRow(sh)
    UniqueCheck sh, "PromoID", map("PromoID")
    
    Dim r As Long
    For r = 2 To lastRow
        Req sh, r, map("PromoID")
        Req sh, r, map("SKU")
        Req sh, r, map("StoreID")
        IsDateCell sh, r, map("WeekStart")
        IsDateCell sh, r, map("WeekEnd")
        If IsDateVal(sh.Cells(r, map("WeekStart"))) And IsDateVal(sh.Cells(r, map("WeekEnd"))) Then
            If sh.Cells(r, map("WeekEnd")).Value < sh.Cells(r, map("WeekStart")).Value Then
                LogIssue sh, r, map("WeekEnd"), svERROR, "WeekEnd < WeekStart."
            End If
        End If
        Req sh, r, map("Mechanic")
        IsPct01 sh, r, map("Depth_Pct")
        IsFlag01 sh, r, map("FeatureDisplayFlag")
        IsNonNeg sh, r, map("CoopFunding_LCU")
        
        ' FKs
        FKExists sh, r, map("SKU"), dictSKU, "SKU not found in Products."
        FKExists sh, r, map("StoreID"), dictStore, "StoreID not found in Stores."
        FKExistsDate sh, r, map("WeekStart"), dictWS, "WeekStart not found in Calendar."
        ' Not all promos map to single YearWeek; window is validated by dates present.
    Next r
End Sub

Private Sub ValidatePricelist(ByVal dictSKU As Object, ByVal dictStore As Object, ByVal dictYW As Object)
    Dim sh As Worksheet: Set sh = SheetByName(SH_PRICELIST)
    If sh Is Nothing Then Exit Sub
    
    Dim needCols As Variant
    needCols = Array("YearWeek", "SKU", "StoreID", "ListPrice_LCU", "AvgNetPrice_LCU", "AvgUnitCost_LCU")
    Dim map As Object: Set map = HeaderMap(sh, needCols)
    If map Is Nothing Then Exit Sub
    
    Dim lastRow As Long: lastRow = LastDataRow(sh)
    ' Composite uniqueness: YearWeek+SKU+StoreID
    CompositeUniqueCheck sh, Array(map("YearWeek"), map("SKU"), map("StoreID")), "YearWeek+SKU+StoreID"
    
    Dim r As Long
    For r = 2 To lastRow
        Req sh, r, map("YearWeek")
        Req sh, r, map("SKU")
        Req sh, r, map("StoreID")
        IsNonNeg sh, r, map("ListPrice_LCU")
        IsNonNeg sh, r, map("AvgNetPrice_LCU")
        IsNonNeg sh, r, map("AvgUnitCost_LCU")
        
        FKExists sh, r, map("YearWeek"), dictYW, "YearWeek not found in Calendar."
        FKExists sh, r, map("SKU"), dictSKU, "SKU not found in Products."
        FKExists sh, r, map("StoreID"), dictStore, "StoreID not found in Stores."
    Next r
End Sub

Private Sub ValidateCompetitor(ByVal dictYW As Object)
    Dim sh As Worksheet: Set sh = SheetByName(SH_COMP)
    If sh Is Nothing Then Exit Sub
    
    Dim needCols As Variant
    needCols = Array("YearWeek", "CompetitorBrand", "SKU_Comp", "AvgPrice_LCU", "PromoFlag")
    Dim map As Object: Set map = HeaderMap(sh, needCols)
    If map Is Nothing Then Exit Sub
    
    Dim lastRow As Long: lastRow = LastDataRow(sh)
    Dim r As Long
    For r = 2 To lastRow
        Req sh, r, map("YearWeek")
        Req sh, r, map("CompetitorBrand")
        ' SKU_Comp may be blank if unknown; no strict req
        IsNonNeg sh, r, map("AvgPrice_LCU")
        IsFlag01 sh, r, map("PromoFlag")
        FKExists sh, r, map("YearWeek"), dictYW, "YearWeek not found in Calendar."
    Next r
End Sub

Private Sub ValidateMedia(ByVal dictYW As Object)
    Dim sh As Worksheet: Set sh = SheetByName(SH_MEDIA)
    If sh Is Nothing Then Exit Sub
    
    Dim needCols As Variant
    needCols = Array("YearWeek", "Channel", "Spend_LCU", "Impressions", "GRPs")
    Dim map As Object: Set map = HeaderMap(sh, needCols)
    If map Is Nothing Then Exit Sub
    
    Dim lastRow As Long: lastRow = LastDataRow(sh)
    Dim r As Long
    For r = 2 To lastRow
        Req sh, r, map("YearWeek")
        Req sh, r, map("Channel")
        IsNonNeg sh, r, map("Spend_LCU")
        IsNonNegInt sh, r, map("Impressions")
        IsNonNeg sh, r, map("GRPs")
        FKExists sh, r, map("YearWeek"), dictYW, "YearWeek not found in Calendar."
    Next r
End Sub

' ============================================================
' Helpers – logging, headers, indices, checks
' ============================================================

Private Sub PrepareValidationSheet()
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(SH_VALIDATION)
    On Error GoTo 0
    
    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        sh.name = SH_VALIDATION
    Else
        sh.Cells.Clear
    End If
    
    With sh
        .Range("A1:E1").EntireRow.Font.Bold = True
        .Range("A1").Value = "Sheet"
        .Range("B1").Value = "Row"
        .Range("C1").Value = "Column"
        .Range("D1").Value = "Severity"
        .Range("E1").Value = "Message"
        .Columns("A:E").ColumnWidth = 22
    End With
End Sub

Private Sub FinalizeValidationSheet()
    Dim sh As Worksheet: Set sh = SheetByName(SH_VALIDATION)
    If sh Is Nothing Then Exit Sub
    With sh
        .Rows(1).Font.Bold = True
        .Columns.AutoFit
        .Activate
    End With
End Sub

Private Sub ClearOldHighlights()
    Dim nm As name
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Clear previous coloring without nuking real formatting
        On Error Resume Next
        ws.Cells.FormatConditions.Delete
        On Error GoTo 0
        ' Also clear our simple color fills
        ws.Cells.Interior.ColorIndex = xlNone
    Next ws
End Sub

Private Function SheetByName(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If SheetByName Is Nothing Then
        LogIssueNothing nm, "Sheet is missing.", svERROR
    End If
End Function

Private Function HeaderMap(ByVal sh As Worksheet, ByVal reqCols As Variant) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lastCol As Long: lastCol = sh.Cells(1, sh.Columns.Count).End(xlToLeft).Column
    Dim c As Long, name As String
    
    Dim idx As Object: Set idx = CreateObject("Scripting.Dictionary")
    For c = 1 To lastCol
        name = Trim(CStr(sh.Cells(1, c).Value))
        If Len(name) > 0 Then
            idx(LCase$(name)) = c
        End If
    Next c
    
    Dim i As Long, need As String
    For i = LBound(reqCols) To UBound(reqCols)
        need = CStr(reqCols(i))
        If idx.Exists(LCase$(need)) Then
            dict(need) = idx(LCase$(need))
        Else
            LogIssue sh, 1, 1, svERROR, "Missing required column '" & need & "'."
        End If
    Next i
    
    Set HeaderMap = dict
End Function

Private Function LastDataRow(ByVal sh As Worksheet) As Long
    Dim r As Long
    r = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    If r < 2 Then r = 1
    LastDataRow = r
End Function

' Build a dictionary of values from a single column (used for FK existence checks)
Private Function BuildIndexDict(ByVal sheetName As String, ByVal colName As String) As Object
    Dim sh As Worksheet: Set sh = SheetByName(sheetName)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    If sh Is Nothing Then
        Set BuildIndexDict = dict
        Exit Function
    End If
    Dim map As Object: Set map = HeaderMap(sh, Array(colName))
    If map Is Nothing Or Not map.Exists(colName) Then
        Set BuildIndexDict = dict
        Exit Function
    End If
    Dim lastRow As Long: lastRow = LastDataRow(sh)
    Dim r As Long, v
    For r = 2 To lastRow
        v = sh.Cells(r, map(colName)).Value
        If Len(Trim(CStr(v))) > 0 Then dict(CStr(v)) = True
    Next r
    Set BuildIndexDict = dict
End Function

' === Logging ===
Private Sub LogIssueNothing(ByVal sheetName As String, ByVal msg As String, ByVal severity As Sev)
    Dim sh As Worksheet: Set sh = SheetByName(SH_VALIDATION)
    If sh Is Nothing Then Exit Sub
    Dim nr As Long: nr = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row + 1
    sh.Cells(nr, 1).Value = sheetName
    sh.Cells(nr, 2).Value = "-"
    sh.Cells(nr, 3).Value = "-"
    sh.Cells(nr, 4).Value = IIf(severity = svERROR, "ERROR", "WARNING")
    sh.Cells(nr, 5).Value = msg
End Sub

Private Sub LogIssue(ByVal sh As Worksheet, ByVal rowIdx As Long, ByVal colIdx As Long, _
                     ByVal severity As Sev, ByVal msg As String)
    Dim log As Worksheet: Set log = SheetByName(SH_VALIDATION)
    If log Is Nothing Then Exit Sub
    Dim nr As Long: nr = log.Cells(log.Rows.Count, 1).End(xlUp).Row + 1
    log.Cells(nr, 1).Value = sh.name
    log.Cells(nr, 2).Value = rowIdx
    log.Cells(nr, 3).Value = sh.Cells(1, colIdx).Value
    log.Cells(nr, 4).Value = IIf(severity = svERROR, "ERROR", "WARNING")
    log.Cells(nr, 5).Value = msg
    
    ' Color the offending cell
    On Error Resume Next
    If severity = svERROR Then
        sh.Cells(rowIdx, colIdx).Interior.Color = CLR_ERR
    Else
        sh.Cells(rowIdx, colIdx).Interior.Color = CLR_WARN
    End If
    On Error GoTo 0
End Sub

' === Generic checks ===
Private Sub Req(ByVal sh As Worksheet, ByVal r As Long, ByVal c As Long)
    If Len(Trim(CStr(sh.Cells(r, c).Value))) = 0 Then
        LogIssue sh, r, c, svERROR, "Required value is blank."
    End If
End Sub

Private Sub IsDateCell(ByVal sh As Worksheet, ByVal r As Long, ByVal c As Long)
    If Not IsDateVal(sh.Cells(r, c)) Then
        LogIssue sh, r, c, svERROR, "Invalid date."
    End If
End Sub

Private Function IsDateVal(ByVal cel As Range) As Boolean
    On Error GoTo EH
    If IsDate(cel.Value) Then
        IsDateVal = True
    Else
        IsDateVal = False
    End If
    Exit Function
EH:
    IsDateVal = False
End Function


Private Sub IsPos(ByVal sh As Worksheet, ByVal r As Long, ByVal c As Long)
    Dim v As Variant: v = sh.Cells(r, c).Value
    If Not IsNumeric(v) Then
        LogIssue sh, r, c, svERROR, "Not numeric."
    ElseIf v <= 0 Then
        LogIssue sh, r, c, svERROR, "Must be > 0."
    End If
End Sub

Private Sub IsNonNegInt(ByVal sh As Worksheet, ByVal r As Long, ByVal c As Long)
    Dim v As Variant: v = sh.Cells(r, c).Value
    If Not IsNumeric(v) Then
        LogIssue sh, r, c, svERROR, "Not numeric."
    ElseIf v < 0 Or v <> Fix(v) Then
        LogIssue sh, r, c, svERROR, "Must be an integer >= 0."
    End If
End Sub

Private Sub IsPosInt(ByVal sh As Worksheet, ByVal r As Long, ByVal c As Long)
    Dim v As Variant: v = sh.Cells(r, c).Value
    If Not IsNumeric(v) Then
        LogIssue sh, r, c, svERROR, "Not numeric."
    ElseIf v <= 0 Or v <> Fix(v) Then
        LogIssue sh, r, c, svERROR, "Must be an integer > 0."
    End If
End Sub

Private Sub IsFlag01(ByVal sh As Worksheet, ByVal r As Long, ByVal c As Long)
    Dim v As Variant: v = sh.Cells(r, c).Value
    If VarType(v) = vbBoolean Then Exit Sub
    If IsNumeric(v) Then
        If v = 0 Or v = 1 Then Exit Sub
    End If
    Dim s As String: s = UCase$(Trim(CStr(v)))
    If s = "0" Or s = "1" Or s = "TRUE" Or s = "FALSE" Or s = "Y" Or s = "N" Then
        If s = "Y" Then sh.Cells(r, c).Value = 1
        If s = "N" Then sh.Cells(r, c).Value = 0
        Exit Sub
    End If
    LogIssue sh, r, c, svERROR, "Flag must be 0/1/TRUE/FALSE (Y/N mapped)."
End Sub

' Accept 0–1 as OK; warn if 1–100 looks like percent typed as whole number
Private Sub IsPct01(ByVal sh As Worksheet, ByVal r As Long, ByVal c As Long)
    Dim v As Variant: v = sh.Cells(r, c).Value
    If Not IsNumeric(v) Then
        LogIssue sh, r, c, svERROR, "Percent not numeric."
    ElseIf v < 0 Then
        LogIssue sh, r, c, svERROR, "Percent < 0."
    ElseIf v <= 1 Then
        ' OK
    ElseIf v <= 100 Then
        LogIssue sh, r, c, svWARNING, "Percent appears as 1–100; expected 0–1."
    Else
        LogIssue sh, r, c, svERROR, "Percent > 100."
    End If
End Sub

' ---- Integer in range, with safe coercion ----
Private Sub InRangeInt(ByVal sh As Worksheet, ByVal r As Long, ByVal c As Long, _
                       ByVal lo As Long, ByVal hi As Long, ByVal msg As String)
    Dim v As Variant: v = sh.Cells(r, c).Value

    If IsError(v) Then
        LogIssue sh, r, c, svERROR, "Not an integer."
        Exit Sub
    End If
    If Not IsNumeric(v) Then
        LogIssue sh, r, c, svERROR, "Not an integer."
        Exit Sub
    End If

    Dim d As Double
    On Error GoTo ConvertFail
    d = CDbl(v)
    On Error GoTo 0

    If d <> Fix(d) Then
        LogIssue sh, r, c, svERROR, "Not an integer."
        Exit Sub
    End If

    If d < lo Or d > hi Then
        LogIssue sh, r, c, svERROR, msg
    End If
    Exit Sub

ConvertFail:
    On Error GoTo 0
    LogIssue sh, r, c, svERROR, "Not an integer."
End Sub


' Foreign key exists in dictionary
Private Sub FKExists(ByVal sh As Worksheet, ByVal r As Long, ByVal c As Long, _
                     ByVal dict As Object, ByVal errMsg As String)
    Dim k As String: k = Trim(CStr(sh.Cells(r, c).Value))
    If Len(k) = 0 Then Exit Sub ' handled by Req if required
    If dict Is Nothing Then Exit Sub
    If Not dict.Exists(k) Then
        LogIssue sh, r, c, svERROR, errMsg
    End If
End Sub

' Uniqueness for single column
Private Sub UniqueCheck(ByVal sh As Worksheet, ByVal keyName As String, ByVal colIdx As Long)
    Dim lastRow As Long: lastRow = LastDataRow(sh)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As String
    For r = 2 To lastRow
        k = Trim(CStr(sh.Cells(r, colIdx).Value))
        If Len(k) > 0 Then
            If dict.Exists(k) Then
                LogIssue sh, r, colIdx, svERROR, keyName & " duplicated (also in row " & dict(k) & ")."
            Else
                dict(k) = r
            End If
        End If
    Next r
End Sub

' Uniqueness for composite key
Private Sub CompositeUniqueCheck(ByVal sh As Worksheet, ByVal cols As Variant, ByVal keyLabel As String)
    Dim lastRow As Long: lastRow = LastDataRow(sh)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long, keyVal As String, i As Long, v
    For r = 2 To lastRow
        keyVal = ""
        For i = LBound(cols) To UBound(cols)
            v = sh.Cells(r, cols(i)).Value
            keyVal = keyVal & "|" & Trim(CStr(v))
        Next i
        If dict.Exists(keyVal) Then
            LogIssue sh, r, cols(LBound(cols)), svERROR, keyLabel & " duplicated (also in row " & dict(keyVal) & ")."
        Else
            dict(keyVal) = r
        End If
    Next r
End Sub

' Accepts Q1..Q4 or 1..4 (text/number). Returns 1..4 if valid; otherwise Empty.
Private Function QuarterToInt(v As Variant) As Variant
    Dim s As String: s = UCase$(Trim$(v & ""))
    If s = "" Then QuarterToInt = Empty: Exit Function
    s = Replace(s, "Q", "")
    If IsNumeric(s) Then
        If CLng(s) >= 1 And CLng(s) <= 4 Then
            QuarterToInt = CLng(s)
            Exit Function
        End If
    End If
    QuarterToInt = Empty
End Function

' === Checks for numeric >= 0 (used by NetPrice_LCU, NetRevenue_LCU, etc.) ===
Private Sub IsNonNeg(ByVal sh As Worksheet, ByVal r As Long, ByVal c As Long)
    Dim v As Variant: v = sh.Cells(r, c).Value
    If IsError(v) Then
        LogIssue sh, r, c, svERROR, "Not numeric."
        Exit Sub
    End If
    If Not IsNumeric(v) Then
        LogIssue sh, r, c, svERROR, "Not numeric."
    ElseIf CDbl(v) < 0 Then
        LogIssue sh, r, c, svERROR, "Must be >= 0."
    End If
End Sub
' Build a dictionary keyed by DATE SERIALS (CLng(CDate(...))) for FK checks on date columns
Private Function BuildIndexDictDate(ByVal sheetName As String, ByVal colName As String) As Object
    Dim sh As Worksheet: Set sh = SheetByName(sheetName)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    If sh Is Nothing Then Set BuildIndexDictDate = dict: Exit Function

    Dim map As Object: Set map = HeaderMap(sh, Array(colName))
    If map Is Nothing Or Not map.Exists(colName) Then Set BuildIndexDictDate = dict: Exit Function

    Dim lastRow As Long: lastRow = LastDataRow(sh)
    Dim r As Long, v
    For r = 2 To lastRow
        v = sh.Cells(r, map(colName)).Value
        If IsDate(v) Then dict(CLng(CDate(v))) = True
    Next r
    Set BuildIndexDictDate = dict
End Function

' Date-aware FK check: compares by serial, so text "2024-04-01" equals a true Date 4/1/2024
Private Sub FKExistsDate(ByVal sh As Worksheet, ByVal r As Long, ByVal c As Long, _
                         ByVal dict As Object, ByVal errMsg As String)
    Dim v As Variant: v = sh.Cells(r, c).Value
    If Not IsDate(v) Then Exit Sub   ' your IsDateCell already logs invalid dates
    If dict Is Nothing Then Exit Sub
    If Not dict.Exists(CLng(CDate(v))) Then
        LogIssue sh, r, c, svERROR, errMsg
    End If
End Sub


