Option Explicit

Public gCancelFlag As Boolean
Public gChartsCount As Long
Public gSilentMode As Boolean

' === GLOBAL COUNTERS ===
Public gLinksCount As Long
Public gTablesCount As Long
Public gColorCount As Long
Public gDeltaCount As Long

' ============================================
'   NEW GLOBAL TO SHARE EXCEL PATH BETWEEN
'   TABLE + CHART UPDATES
' ============================================
Public gSelectedExcelFile As String  ' holds the file chosen in Step_1a
Public gBatchMode As Boolean      ' True when running the full STEP_1_RUN_ALL sequence

' ============================================
'   SHARED EXCEL INSTANCE (batch mode)
'   Step_1b keeps Excel alive; Step_1c reuses it
' ============================================
Public gExcelApp As Object
Public gOpenedWorkbooks As Object



' ============================================
'         NEW COLOR PRESERVATION TOGGLE
' ============================================
Public PreserveFormattedColor As Boolean
' (Kept from code_v7, but used only for reuse logic on \"ntbl_\" or forced false for \"htmp_\".)

' ==================================================
'       COLOR CONSTANTS (for Excel 3-color scale)
' ==================================================
Public Const COLOR_MINIMUM As String = "#F8696B"   ' Default Red
Public Const COLOR_MIDPOINT As String = "#FFEB84"  ' Default Yellow
Public Const COLOR_MAXIMUM As String = "#63BE7B"   ' Default Green

' Font color constants for contrast text
Public Const COLOR_DARKFONT As String = "#000000"  ' Default black
Public Const COLOR_LIGHTFONT As String = "#FFFFFF" ' Default white

' ==================================================
'       DATA STRUCTURE: TABLE FORMATTING DATA
' ==================================================
Type TableFormattingData
    ' Layout and text
    RowHeights() As Single
    ColumnWidths() As Single
    FontNames() As Variant
    FontSizes() As Variant
    FontBolds() As Variant
    FontItalics() As Variant
    FontColors() As Variant
    FontUnderline() As Variant
    FontShadow() As Variant
    HorizontalAlignments() As Variant
    VerticalAlignments() As Variant
    MarginLefts() As Variant
    MarginRights() As Variant
    MarginTops() As Variant
    MarginBottoms() As Variant
    ShapeLeft As Single
    ShapeTop As Single
    ShapeWidth As Single
    ShapeHeight As Single
    
    ' Borders
    BorderVisible() As Variant
    BorderWeight() As Variant
    BorderDashStyle() As Variant
    BorderColor() As Variant
    
    ' NEW: Store each cell's fill color (Long/RGB)
    CellFillColors() As Variant
    CellFillVisible() As Boolean
    CellFillType() As Long
    
    ' NEW: Transparency
    CellFillTransparency() As Single
End Type

Public Const xlConditionValueLowestValue As Long = 1
Public Const xlConditionValueHighestValue As Long = 2
Public Const xlConditionValuePercentile As Long = 5

Public borderTypes As Variant

' ==================================================
'        INITIALIZATION & UTILITIES
' ==================================================
Private Sub InitializeBorderTypes()
    borderTypes = Array(ppBorderLeft, ppBorderRight, ppBorderTop, ppBorderBottom)
End Sub

Public Function ExtractFileName(fullName As String) As String
    Dim parts() As String
    parts = Split(fullName, "!")
    If UBound(parts) >= 0 Then
        ExtractFileName = parts(0)
    Else
        ExtractFileName = "Not Specified"
    End If
End Function

Public Function ExtractSheetName(fullName As String) As String
    Dim parts() As String
    parts = Split(fullName, "!")
    If UBound(parts) >= 1 Then
        ExtractSheetName = parts(1)
    Else
        ExtractSheetName = "Not Specified"
    End If
End Function

Public Function ExtractAndConvertRange(fullName As String) As String
    Dim parts() As String
    If InStr(fullName, "!") > 0 Then
        parts = Split(fullName, "!")
        If UBound(parts) >= 2 Then
            ExtractAndConvertRange = ConvertR1C1ToA1(parts(2))
            Exit Function
        End If
    End If
    ExtractAndConvertRange = "Not Specified"
End Function

Public Function ConvertR1C1ToA1(r1c1Range As String) As String
    Dim parts() As String
    Dim startCell As String, endCell As String
    
    parts = Split(r1c1Range, ":")
    If UBound(parts) = 0 Then
        ConvertR1C1ToA1 = ConvertSingleR1C1ToA1(parts(0))
        Exit Function
    End If
    
    startCell = parts(0)
    endCell = parts(1)
    ConvertR1C1ToA1 = ConvertSingleR1C1ToA1(startCell) & ":" & ConvertSingleR1C1ToA1(endCell)
End Function

Private Function ConvertSingleR1C1ToA1(r1c1Cell As String) As String
    Dim rowNum As Long, colNum As Long
    Dim colLetter As String
    
    rowNum = CLng(Mid(r1c1Cell, 2, InStr(r1c1Cell, "C") - 2))
    colNum = CLng(Mid(r1c1Cell, InStr(r1c1Cell, "C") + 1))
    
    colLetter = ""
    Do While colNum > 0
        colLetter = Chr((colNum - 1) Mod 26 + 65) & colLetter
        colNum = (colNum - 1) \ 26
    Loop
    
    ConvertSingleR1C1ToA1 = colLetter & rowNum
End Function

Public Function HexToRGB(hexColor As String) As Long
    Dim r As Long, g As Long, b As Long
    hexColor = Replace(hexColor, "#", "")
    r = CLng("&H" & Mid(hexColor, 1, 2))
    g = CLng("&H" & Mid(hexColor, 3, 2))
    b = CLng("&H" & Mid(hexColor, 5, 2))
    HexToRGB = RGB(r, g, b)
End Function

' ==================================================
' FIND EXISTING TABLES - UPDATED WITH FLEXIBLE MATCHING
'   Now checks if table name CONTAINS both prefix and linkedName
'   Order doesn't matter: "ntbl_object_5" and "ccst_object_5_ntbl_" both work
' ==================================================
Public Function FindExistingNtblTable(ByVal pptSlide As Slide, ByVal linkedName As String) As Shape
    Dim shp As Shape
    For Each shp In pptSlide.Shapes
        If shp.HasTable Then
            ' Check if shape name contains BOTH "ntbl_" AND linkedName (exact token match)
            If InStr(shp.Name, "ntbl_") > 0 And IsExactTokenMatch(shp.Name, linkedName) Then
                Set FindExistingNtblTable = shp
                Exit Function
            End If
        End If
    Next shp
End Function

Public Function FindExistingHtmpTable(ByVal pptSlide As Slide, ByVal linkedName As String) As Shape
    Dim shp As Shape
    For Each shp In pptSlide.Shapes
        If shp.HasTable Then
            ' Check if shape name contains BOTH "htmp_" AND linkedName (exact token match)
            If InStr(shp.Name, "htmp_") > 0 And IsExactTokenMatch(shp.Name, linkedName) Then
                Set FindExistingHtmpTable = shp
                Exit Function
            End If
        End If
    Next shp
End Function

Public Function FindExistingTrnsTable(ByVal pptSlide As Slide, ByVal linkedName As String) As Shape
    Dim shp As Shape
    For Each shp In pptSlide.Shapes
        If shp.HasTable Then
            ' Check if shape name contains BOTH "trns_" AND linkedName (exact token match)
            If InStr(shp.Name, "trns_") > 0 And IsExactTokenMatch(shp.Name, linkedName) Then
                Set FindExistingTrnsTable = shp
                Exit Function
            End If
        End If
    Next shp
End Function

Public Function FindExistingDeltShape(ByVal pptSlide As Slide, ByVal linkedName As String) As Shape
    Dim shp As Shape
    For Each shp In pptSlide.Shapes
        ' delt_ shapes are NOT tables — search all shapes
        If InStr(shp.Name, "delt_") > 0 And IsExactTokenMatch(shp.Name, linkedName) Then
            Set FindExistingDeltShape = shp
            Exit Function
        End If
    Next shp
End Function

Private Function FindTemplateShape(ByVal templateName As String) As Shape
    Dim shp As Shape
    For Each shp In ActivePresentation.Slides(1).Shapes
        If shp.Name = templateName Then
            Set FindTemplateShape = shp
            Exit Function
        End If
    Next shp
End Function

' ==================================================
' HELPER FUNCTION: EXACT TOKEN MATCHING
' Ensures linkedName appears as complete token, not substring
' ==================================================
Private Function IsExactTokenMatch(ByVal shapeName As String, ByVal linkedName As String) As Boolean
    Dim pos As Long
    Dim nameLen As Long
    Dim linkedLen As Long
    Dim beforeChar As String
    Dim afterChar As String
    
    nameLen = Len(shapeName)
    linkedLen = Len(linkedName)
    pos = InStr(shapeName, linkedName)
    
    ' If not found at all, return False
    If pos = 0 Then
        IsExactTokenMatch = False
        Exit Function
    End If
    
    ' Check all occurrences of linkedName in shapeName
    Do While pos > 0
        ' Check character before (if exists)
        If pos = 1 Then
            beforeChar = ""  ' At start of string
        Else
            beforeChar = Mid(shapeName, pos - 1, 1)
        End If
        
        ' Check character after (if exists)
        If pos + linkedLen > nameLen Then
            afterChar = ""  ' At end of string
        Else
            afterChar = Mid(shapeName, pos + linkedLen, 1)
        End If
        
        ' Check if both before and after are word boundaries (non-alphanumeric)
        If IsWordBoundary(beforeChar) And IsWordBoundary(afterChar) Then
            IsExactTokenMatch = True
            Exit Function
        End If
        
        ' Look for next occurrence
        pos = InStr(pos + 1, shapeName, linkedName)
    Loop
    
    IsExactTokenMatch = False
End Function

Private Function IsWordBoundary(ByVal char As String) As Boolean
    ' Empty string (start/end), underscore, space, or non-alphanumeric characters are boundaries
    If char = "" Then
        IsWordBoundary = True
    ElseIf char = "_" Or char = " " Or char = "-" Then
        IsWordBoundary = True
    ElseIf Not ((char >= "A" And char <= "Z") Or (char >= "a" And char <= "z") Or (char >= "0" And char <= "9")) Then
        IsWordBoundary = True
    Else
        IsWordBoundary = False
    End If
End Function

' ==================================================
'       FORMATTING EXTRACTION / APPLICATION
'   (unchanged from code_v7)
' ==================================================

Public Function ExtractTableFormatting(tblShape As Shape) As TableFormattingData
    Dim tbl As Table
    Dim i As Long, j As Long, k As Long
    Dim numRows As Long, numCols As Long
    Dim tData As TableFormattingData
    
    If Not tblShape.HasTable Then Exit Function
    Set tbl = tblShape.Table
    
    numRows = tbl.Rows.Count
    numCols = tbl.Columns.Count
    
    tData.ShapeLeft = tblShape.Left
    tData.ShapeTop = tblShape.Top
    tData.ShapeWidth = tblShape.Width
    tData.ShapeHeight = tblShape.Height
    
    ReDim tData.RowHeights(1 To numRows)
    ReDim tData.ColumnWidths(1 To numCols)
    ReDim tData.FontNames(1 To numRows, 1 To numCols)
    ReDim tData.FontSizes(1 To numRows, 1 To numCols)
    ReDim tData.FontBolds(1 To numRows, 1 To numCols)
    ReDim tData.FontItalics(1 To numRows, 1 To numCols)
    ReDim tData.FontColors(1 To numRows, 1 To numCols)
    ReDim tData.FontUnderline(1 To numRows, 1 To numCols)
    ReDim tData.FontShadow(1 To numRows, 1 To numCols)
    ReDim tData.HorizontalAlignments(1 To numRows, 1 To numCols)
    ReDim tData.VerticalAlignments(1 To numRows, 1 To numCols)
    ReDim tData.MarginLefts(1 To numRows, 1 To numCols)
    ReDim tData.MarginRights(1 To numRows, 1 To numCols)
    ReDim tData.MarginTops(1 To numRows, 1 To numCols)
    ReDim tData.MarginBottoms(1 To numRows, 1 To numCols)
    ReDim tData.BorderVisible(1 To numRows, 1 To numCols, 1 To 4)
    ReDim tData.BorderWeight(1 To numRows, 1 To numCols, 1 To 4)
    ReDim tData.BorderDashStyle(1 To numRows, 1 To numCols, 1 To 4)
    ReDim tData.BorderColor(1 To numRows, 1 To numCols, 1 To 4)
    
    ' NEW: Store cell fill colors & transparency
    ReDim tData.CellFillColors(1 To numRows, 1 To numCols)
    ReDim tData.CellFillVisible(1 To numRows, 1 To numCols)
    ReDim tData.CellFillType(1 To numRows, 1 To numCols)
    ReDim tData.CellFillTransparency(1 To numRows, 1 To numCols)
    
    Dim shpFill As FillFormat

    For i = 1 To numRows
        tData.RowHeights(i) = tbl.Rows(i).Height
    Next i
    
    For j = 1 To numCols
        tData.ColumnWidths(j) = tbl.Columns(j).Width
    Next j
    
    For i = 1 To numRows
        For j = 1 To numCols
            ' Font properties
            With tbl.Cell(i, j).Shape.TextFrame.TextRange.Font
                tData.FontNames(i, j) = .Name
                tData.FontSizes(i, j) = .Size
                tData.FontBolds(i, j) = .Bold
                tData.FontItalics(i, j) = .Italic
                tData.FontColors(i, j) = .Color.RGB
                tData.FontUnderline(i, j) = .Underline
                tData.FontShadow(i, j) = .Shadow
            End With
            
            ' Paragraph / margins
            With tbl.Cell(i, j).Shape.TextFrame
                tData.MarginLefts(i, j) = .MarginLeft
                tData.MarginRights(i, j) = .MarginRight
                tData.MarginTops(i, j) = .MarginTop
                tData.MarginBottoms(i, j) = .MarginBottom
                tData.HorizontalAlignments(i, j) = .TextRange.ParagraphFormat.Alignment
                tData.VerticalAlignments(i, j) = .VerticalAnchor
            End With
            
            ' fill
            Set shpFill = tbl.Cell(i, j).Shape.Fill
            
            tData.CellFillVisible(i, j) = (shpFill.Visible <> msoFalse)
            tData.CellFillType(i, j) = shpFill.Type ' e.g. msoFillSolid, msoFillGradient, etc.
            tData.CellFillColors(i, j) = shpFill.ForeColor.RGB
            
            ' NEW: Transparency
            tData.CellFillTransparency(i, j) = shpFill.Transparency
            
            ' Border properties
            Dim bType As Long
            For k = LBound(borderTypes) To UBound(borderTypes)
                bType = borderTypes(k)
                On Error Resume Next
                With tbl.Cell(i, j).Borders(bType)
                    If Err.Number = 0 Then
                        tData.BorderVisible(i, j, k + 1) = .Visible
                        tData.BorderWeight(i, j, k + 1) = .Weight
                        tData.BorderDashStyle(i, j, k + 1) = .DashStyle
                        tData.BorderColor(i, j, k + 1) = .ForeColor.RGB
                    Else
                        Err.Clear
                        tData.BorderVisible(i, j, k + 1) = msoFalse
                        tData.BorderWeight(i, j, k + 1) = 0
                        tData.BorderDashStyle(i, j, k + 1) = msoLineSolid
                        tData.BorderColor(i, j, k + 1) = RGB(255, 255, 255)
                    End If
                End With
                On Error GoTo 0
            Next k
        Next j
    Next i
    
    ExtractTableFormatting = tData
End Function

Public Sub ApplyOldFormattingToNewTable(newTblShape As Shape, oldData As TableFormattingData)
    Dim tbl As Table
    Dim i As Long, j As Long, k As Long
    
    Set tbl = newTblShape.Table
    
    ' ============= Apply row & column sizing =============
    For i = 1 To UBound(oldData.RowHeights)
        On Error Resume Next
        tbl.Rows(i).Height = oldData.RowHeights(i)
        On Error GoTo 0
    Next i
    
    For j = 1 To UBound(oldData.ColumnWidths)
        On Error Resume Next
        tbl.Columns(j).Width = oldData.ColumnWidths(j)
        On Error GoTo 0
    Next j
    
    Dim numRows As Long, numCols As Long
    numRows = tbl.Rows.Count
    numCols = tbl.Columns.Count
    
    ' ============= Apply per-cell formatting =============
    For i = 1 To numRows
        For j = 1 To numCols
            ' --- Fill color ---
            If PreserveFormattedColor Then
                ' Check if old cell fill was visible
                If oldData.CellFillVisible(i, j) = False Then
                    tbl.Cell(i, j).Shape.Fill.Visible = msoFalse
                Else
                    tbl.Cell(i, j).Shape.Fill.Visible = msoTrue
                    Select Case oldData.CellFillType(i, j)
                        Case msoFillSolid
                            tbl.Cell(i, j).Shape.Fill.Solid
                            tbl.Cell(i, j).Shape.Fill.ForeColor.RGB = oldData.CellFillColors(i, j)
                            
                            ' NEW: Transparency for solid
                            tbl.Cell(i, j).Shape.Fill.Transparency = oldData.CellFillTransparency(i, j)
                        
                        Case msoFillGradient
                            ' Minimal gradient approach
                            tbl.Cell(i, j).Shape.Fill.TwoColorGradient msoGradientDiagonalUp, 1
                            tbl.Cell(i, j).Shape.Fill.ForeColor.RGB = oldData.CellFillColors(i, j)
                            
                            ' NEW: Transparency
                            tbl.Cell(i, j).Shape.Fill.Transparency = oldData.CellFillTransparency(i, j)
                        
                        Case Else
                            ' default to solid with the stored color
                            tbl.Cell(i, j).Shape.Fill.Solid
                            tbl.Cell(i, j).Shape.Fill.ForeColor.RGB = oldData.CellFillColors(i, j)
                            
                            ' NEW: Transparency
                            tbl.Cell(i, j).Shape.Fill.Transparency = oldData.CellFillTransparency(i, j)
                    End Select
                End If
            End If
            
            ' --- Font props ---
            With tbl.Cell(i, j).Shape.TextFrame.TextRange.Font
                On Error Resume Next
                .Name = oldData.FontNames(i, j)
                .Size = oldData.FontSizes(i, j)
                .Bold = oldData.FontBolds(i, j)
                .Italic = oldData.FontItalics(i, j)
                .Color.RGB = oldData.FontColors(i, j)
                .Underline = oldData.FontUnderline(i, j)
                .Shadow = oldData.FontShadow(i, j)
                On Error GoTo 0
            End With
            
            ' --- Paragraph & margins ---
            With tbl.Cell(i, j).Shape.TextFrame
                On Error Resume Next
                .MarginLeft = oldData.MarginLefts(i, j)
                .MarginRight = oldData.MarginRights(i, j)
                .MarginTop = oldData.MarginTops(i, j)
                .MarginBottom = oldData.MarginBottoms(i, j)
                .TextRange.ParagraphFormat.Alignment = oldData.HorizontalAlignments(i, j)
                .VerticalAnchor = oldData.VerticalAlignments(i, j)
                On Error GoTo 0
            End With
            
            ' --- Borders ---
            Dim bType As Long
            For k = LBound(borderTypes) To UBound(borderTypes)
                bType = borderTypes(k)
                With tbl.Cell(i, j).Borders(bType)
                    On Error Resume Next
                    If oldData.BorderVisible(i, j, k + 1) = msoFalse Then
                        .Visible = msoFalse
                        .Weight = 0
                        .DashStyle = msoLineSolid
                        .ForeColor.RGB = RGB(255, 255, 255)
                    Else
                        .Visible = oldData.BorderVisible(i, j, k + 1)
                        .Weight = oldData.BorderWeight(i, j, k + 1)
                        .DashStyle = oldData.BorderDashStyle(i, j, k + 1)
                        .ForeColor.RGB = oldData.BorderColor(i, j, k + 1)
                    End If
                    On Error GoTo 0
                End With
            Next k
        Next j
    Next i
End Sub

' ==================================================
'             APPLY 3-COLOR SCALE
'      (same code_v7 logic for "htmp_")
' ==================================================
Private Sub ApplyColorScale(ByVal targetRange As Object)
    ' Delete old formatting
    targetRange.FormatConditions.Delete
    
    ' 3-color scale
    With targetRange.FormatConditions.AddColorScale(ColorScaleType:=3)
        With .ColorScaleCriteria(1)
            .Type = xlConditionValueLowestValue
            .FormatColor.Color = HexToRGB(COLOR_MINIMUM)
        End With
        With .ColorScaleCriteria(2)
            .Type = xlConditionValuePercentile
            .Value = 50
            .FormatColor.Color = HexToRGB(COLOR_MIDPOINT)
        End With
        With .ColorScaleCriteria(3)
            .Type = xlConditionValueHighestValue
            .FormatColor.Color = HexToRGB(COLOR_MAXIMUM)
        End With
    End With
End Sub

' ==================================================
'   GET CONTRAST FONT COLOR (BASED ON BRIGHTNESS)
' ==================================================
Private Function GetContrastFontColor(ByVal cellBackgroundColor As Long, _
                                      ByVal darkFontColorRGB As Long, _
                                      ByVal lightFontColorRGB As Long) As Long
    Dim r As Long, g As Long, b As Long
    Dim brightness As Double
    
    ' Convert from long to R/G/B
    r = cellBackgroundColor Mod 256
    g = (cellBackgroundColor \ 256) Mod 256
    b = (cellBackgroundColor \ 65536) Mod 256
    
    ' Weighted brightness (common formula)
    brightness = 0.299 * r + 0.587 * g + 0.114 * b
    
    ' Threshold
    If brightness < 128 Then
        GetContrastFontColor = lightFontColorRGB
    Else
        GetContrastFontColor = darkFontColorRGB
    End If
End Function

' ==================================================
'  PROCESS A SINGLE LINKED SHAPE (MODULARIZED)
' ==================================================
Private Sub ProcessLinkedShape(ByVal pptSlide As Slide, _
                               ByVal pptShape As Shape, _
                               ByVal excelApp As Object, _
                               ByVal openedWorkbooks As Object)

    Dim filePath As String, sheetName As String, rangeAddress As String
    Dim existingTable As Shape
    Dim oldFormatting As TableFormattingData
    Dim hasOldFormatting As Boolean
    Dim excelWorkbook As Object
    Dim excelSheet As Object
    Dim cellRange As Object
    Dim tableShape As Shape
    
    ' <<< ADDED: handle "trns_" >>>
    Dim existingHtmp As Shape
    Dim existingTrns As Shape
    
    ' localPreserve logic from code_v7
    Dim localPreserve As Boolean
    
    ' -------------------------------
    ' Extract workbook / sheet / range
    ' -------------------------------
    filePath = ExtractFileName(pptShape.LinkFormat.SourceFullName)
    sheetName = ExtractSheetName(pptShape.LinkFormat.SourceFullName)
    rangeAddress = ExtractAndConvertRange(pptShape.LinkFormat.SourceFullName)
    
    ' Validate the range before proceeding
    If rangeAddress = "Not Specified" Then Exit Sub
    If Dir(filePath) = "" Then Exit Sub
    
    ' ------------------------------------------------
    ' 1) Check for "ntbl_<>" or "htmp_<>" or "trns_<>"
    ' ------------------------------------------------
    hasOldFormatting = False
    
    Set existingTable = FindExistingNtblTable(pptSlide, pptShape.Name)  ' ntbl_
    If existingTable Is Nothing Then
        Set existingHtmp = FindExistingHtmpTable(pptSlide, pptShape.Name)  ' htmp_
        If Not existingHtmp Is Nothing Then
            Set existingTable = existingHtmp
        Else
            ' try "trns_"
            Set existingTrns = FindExistingTrnsTable(pptSlide, pptShape.Name)
            If Not existingTrns Is Nothing Then
                Set existingTable = existingTrns
            End If
        End If
    End If
    
    If Not existingTable Is Nothing Then
        oldFormatting = ExtractTableFormatting(existingTable)
        hasOldFormatting = True
    Else
        ' If no ntbl_/htmp_/trns_ exists but a delt_ shape does,
        ' this OLE is delt_-only — skip creating an ntbl_ table.
        If Not FindExistingDeltShape(pptSlide, pptShape.Name) Is Nothing Then
            Exit Sub
        End If
    End If

    ' -------------------------------
    ' Open or get the Excel workbook
    ' -------------------------------
WorkbookError:
    On Error GoTo WorkbookError
    If Not openedWorkbooks.Exists(filePath) Then
        Set excelWorkbook = excelApp.Workbooks.Open(filePath, ReadOnly:=False)
        openedWorkbooks.Add filePath, excelWorkbook
    Else
        Set excelWorkbook = openedWorkbooks(filePath)
    End If
    On Error GoTo 0
    
    Set excelSheet = excelWorkbook.Sheets(sheetName)
    Set cellRange = excelSheet.Range(rangeAddress)
    
    ' ----------------------------------------------------
    ' Check the prefix to decide if we do heatmap or transpose
    ' (Same code_v7 approach: if shape starts with "htmp_", localPreserve=False => re-extract color
    '   if shape starts with "ntbl_", localPreserve=True => preserve
    '   if shape starts with "trns_", localPreserve=True => preserve but transpose data
    '   if no shape found => brand-new => preserve color
    ' ----------------------------------------------------
    Dim doTranspose As Boolean
    doTranspose = False
    
    If Not existingTable Is Nothing Then
        If Left(existingTable.Name, 5) = "ntbl_" Then
            localPreserve = True
        ElseIf Left(existingTable.Name, 5) = "htmp_" Then
            localPreserve = False
        ElseIf Left(existingTable.Name, 5) = "trns_" Then
            localPreserve = True
            doTranspose = True
        End If
    Else
        ' brand new => same logic as code_v7
        localPreserve = True
    End If
    
    ' If no old table or localPreserve=False => re-extract color scale + font contrast
    If (Not hasOldFormatting) Or (localPreserve = False) Then
        ApplyColorScale cellRange
        
        ' Force recalc
        excelApp.Calculate
        excelSheet.Calculate
        
        ' Adjust Font Contrast
        Dim targetCell As Object
        Dim cellColor As Long
        Dim darkFontColorRGB As Long, lightFontColorRGB As Long
        
        darkFontColorRGB = HexToRGB(COLOR_DARKFONT)
        lightFontColorRGB = HexToRGB(COLOR_LIGHTFONT)
        
        For Each targetCell In cellRange
            cellColor = targetCell.DisplayFormat.Interior.Color
            targetCell.Font.Color = GetContrastFontColor(cellColor, darkFontColorRGB, lightFontColorRGB)
        Next targetCell
    End If
    
    ' ---------------------------------
    ' Reuse existing shape or create a new "ntbl_" shape
    ' ---------------------------------
    Dim rowIndex As Long, colIndex As Long
    Dim cellFormattedColor As Long
    
    If hasOldFormatting Then
        ' Reuse existing shape
        Set tableShape = existingTable
        ' keep old layout
        ApplyOldFormattingToNewTable tableShape, oldFormatting
    Else
        ' brand new => create
        If doTranspose Then
            ' If we want to physically create a transposed dimension table, we can swap # rows/columns
            ' But for minimal changes, code_v7 always used cellRange.Rows.Count x cellRange.Columns.Count
            ' We'll do the same, and rely on the data fill for transposing. If you prefer actually
            ' flipping the shape's row/col, do it here. For now, keep code_v7 approach:
        End If
        
        Set tableShape = pptSlide.Shapes.AddTable(cellRange.Rows.Count, cellRange.Columns.Count, 100, 100, 400, 200)
        tableShape.Name = "ntbl_" & pptShape.Name
        tableShape.AlternativeText = "Linked to Excel File: " & filePath & vbCrLf & _
                                     "Sheet: " & sheetName & vbCrLf & _
                                     "Range: " & rangeAddress
    End If
    
    ' ---------------------------------
    ' Fill table cells from Excel
    ' If doTranspose => swap rowIndex/colIndex
    ' ---------------------------------
    Dim PPTrow As Long, PPTcol As Long
    Dim totalRows As Long, totalCols As Long
    totalRows = tableShape.Table.Rows.Count
    totalCols = tableShape.Table.Columns.Count
    
    Dim maxRows As Long, maxCols As Long
    maxRows = cellRange.Rows.Count
    maxCols = cellRange.Columns.Count
    
    For rowIndex = 1 To maxRows
        For colIndex = 1 To maxCols
            
            If doTranspose Then
                PPTrow = colIndex
                PPTcol = rowIndex
            Else
                PPTrow = rowIndex
                PPTcol = colIndex
            End If
            
            ' Check boundaries
            If PPTrow <= totalRows And PPTcol <= totalCols Then
                With tableShape.Table.Cell(PPTrow, PPTcol).Shape
                    ' Always copy text
                    .TextFrame.TextRange.Text = cellRange.Cells(rowIndex, colIndex).Text
                    
                    ' If no old formatting or localPreserve=False => re-pull fill & font
                    If (Not hasOldFormatting) Or (localPreserve = False) Then
                        On Error Resume Next
                        cellFormattedColor = cellRange.Cells(rowIndex, colIndex).DisplayFormat.Interior.Color
                        On Error GoTo 0
                        .Fill.ForeColor.RGB = cellFormattedColor
                        
                        With .TextFrame.TextRange.Font
                            .Name = cellRange.Cells(rowIndex, colIndex).Font.Name
                            .Size = cellRange.Cells(rowIndex, colIndex).Font.Size
                            .Bold = cellRange.Cells(rowIndex, colIndex).Font.Bold
                            .Italic = cellRange.Cells(rowIndex, colIndex).Font.Italic
                            .Color = cellRange.Cells(rowIndex, colIndex).Font.Color
                        End With
                    End If
                    
                    '.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter '"commented out so I can mannually align the new tranposed column or row"
                End With
            End If
        Next colIndex
    Next rowIndex
    
    Exit Sub

CouldNotOpenWB:
'    MsgBox "Could not open workbook: " & filePath & vbCrLf & Err.Description, vbExclamation
End Sub

' ==================================================
'           MAIN PROCEDURE
' ==================================================
Sub Step_1b_UpdateTableContent()

    ' By default, let’s preserve any existing color (set to True).
    PreserveFormattedColor = True
    
    Dim pptSlide As Slide
    Dim pptShape As Shape
    
    Dim excelApp As Object
    Dim openedWorkbooks As Object
    
    InitializeBorderTypes
    
    On Error GoTo ErrorHandler
    
    ' Create or attach to Excel
    On Error Resume Next
    Set excelApp = GetObject(, "Excel.Application")
    On Error GoTo ErrorHandler
    If excelApp Is Nothing Then
        Set excelApp = CreateObject("Excel.Application")
    End If
    
    ' Turn off some Excel overhead for performance
    Dim initScreenUpdating As Boolean
    Dim initEnableEvents As Boolean
    Dim initDisplayAlerts As Boolean
    
    initScreenUpdating = excelApp.ScreenUpdating
    initEnableEvents = excelApp.EnableEvents
    initDisplayAlerts = excelApp.DisplayAlerts
    
    excelApp.ScreenUpdating = False
    excelApp.EnableEvents = False
    excelApp.DisplayAlerts = False
    
    Set openedWorkbooks = CreateObject("Scripting.Dictionary")

    ' Share Excel instance with Step_1c in batch mode
    If gBatchMode Then
        Set gExcelApp = excelApp
        Set gOpenedWorkbooks = openedWorkbooks
    End If

    ' Loop slides & shapes
    Dim pres As Presentation
    Dim cntTablesUpdated As Long
    Set pres = ActivePresentation
    
    For Each pptSlide In pres.Slides
        For Each pptShape In pptSlide.Shapes
            If pptShape.Type = msoLinkedOLEObject Then
                If InStr(1, pptShape.OLEFormat.ProgID, "Excel.Sheet") > 0 Then
                    ' Process each shape separately
                    ProcessLinkedShape pptSlide, pptShape, excelApp, openedWorkbooks
                    ' Only count if a table was actually created/updated (not delt_-only)
                    If Not (FindExistingNtblTable(pptSlide, pptShape.Name) Is Nothing _
                       And FindExistingHtmpTable(pptSlide, pptShape.Name) Is Nothing _
                       And FindExistingTrnsTable(pptSlide, pptShape.Name) Is Nothing) Then
                        cntTablesUpdated = cntTablesUpdated + 1
                    End If
                End If
            End If
        Next pptShape
    Next pptSlide
    
    gTablesCount = cntTablesUpdated

    ' In batch mode, keep Excel alive for Step_1c to reuse
    If gBatchMode Then
        If Not gSilentMode Then
            MsgBox "All linked Excel objects processed and tables created.", vbInformation
        End If
        Exit Sub
    End If

    ' Close all opened workbooks (standalone mode only)
    Dim wbKey As Variant
    For Each wbKey In openedWorkbooks.Keys
        On Error Resume Next
        openedWorkbooks(wbKey).Close False
        On Error GoTo 0
    Next wbKey

    ' Restore Excel settings
    excelApp.ScreenUpdating = initScreenUpdating
    excelApp.EnableEvents = initEnableEvents
    excelApp.DisplayAlerts = initDisplayAlerts

    If excelApp.Workbooks.Count = 0 Then excelApp.Quit
    Set excelApp = Nothing

    If Not gSilentMode Then
    MsgBox "All linked Excel objects processed and tables created.", vbInformation
    End If
    Exit Sub

ErrorHandler:
    If Not gSilentMode Then
        MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    End If

    ' In batch mode, leave Excel alive for Step_1c to clean up
    If gBatchMode Then Exit Sub

    On Error Resume Next

    ' Attempt to close any open workbooks
    Dim errKey As Variant
    If Not openedWorkbooks Is Nothing Then
        For Each errKey In openedWorkbooks.Keys
            openedWorkbooks(errKey).Close False
        Next errKey
    End If
    
    ' Restore and quit Excel if we had it open
        '--------------------------------------------------------------------------
        ' NOTE:
        '   We previously hit Run‑time error 462 ("The remote server machine does not
        '   exist or is unavailable") whenever the code tried to reset screen updating
        '   or other properties after Excel had already been closed/crashed.
        '   The small wrapper below fixes that by
        '     • running the restore calls under On Error Resume Next, so any failed
        '       COM call is silently ignored, and
        '     • wrapping the calls in a With excelApp block that is skipped if the
        '       reference has gone "zombie".
        '   This prevents the 462 exception and lets the clean‑up finish gracefully.
        '--------------------------------------------------------------------------
    If Not excelApp Is Nothing Then
        ' Attempt to reset settings *only* if the Excel instance is still alive
        On Error Resume Next ' ignore failures if Excel has already gone away
        With excelApp
            .ScreenUpdating = initScreenUpdating
            .EnableEvents = initEnableEvents
            .DisplayAlerts = initDisplayAlerts
            .Quit
        End With
        On Error GoTo 0
        Set excelApp = Nothing
    End If
End Sub


Sub Step_1a_UpdateTableLinks()
    Dim newFilePath As String
    Dim pptPresentation As Presentation
    Dim pptSlide As Slide
    Dim pptShape As Shape
    
    Dim linkedRangeShapes As New Collection
    Dim updatedCount As Long
    
    '====================================================================
    '                 TOGGLES / SETTINGS
    '====================================================================
    ' 1) If True, do one bulk 'UpdateLinks' at the end
    '    If False, no forced update (faster).
    Dim doSinglePassUpdate As Boolean
    doSinglePassUpdate = True
    
    ' 2) If True, set every link's AutoUpdate mode to Manual
    '    If False, leave them as-is (could remain Automatic).
    Dim setLinksToManual As Boolean
    setLinksToManual = True
    '====================================================================
    
    '-----------------------------------------
    ' 1) Prompt user: pick new Excel file
    '-----------------------------------------
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select the NEW Excel file for linked ranges"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls", 1
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            newFilePath = .SelectedItems(1)
            gSelectedExcelFile = newFilePath    ' store for chart update
        Else
    gLinksCount = updatedCount
    If Not gSilentMode Then
        MsgBox "Operation cancelled. No file selected.", vbInformation
    End If
    gCancelFlag = True
    Exit Sub
        End If
    End With
    
    '-----------------------------------------
    ' 2) Collect all linked OLE objects
    '    that are pasted Excel ranges
    '-----------------------------------------
    Set pptPresentation = ActivePresentation
    
    Dim sld As Slide
    For Each sld In pptPresentation.Slides
        For Each pptShape In sld.Shapes
            CollectLinkedOLERangeShape pptShape, linkedRangeShapes
        Next pptShape
    Next sld
    
    If linkedRangeShapes.Count = 0 Then
    If Not gSilentMode Then
        MsgBox "No linked Excel ranges found in this presentation.", vbInformation, "No Linked Ranges"
    End If
        Exit Sub
    End If
    
    '-----------------------------------------
    ' 3) Open Excel once (invisible)
    '-----------------------------------------
    Dim excelApp As Object
    Dim wb As Object
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    
    ' Optional: open the workbook once (to preload).
    ' If you find you don't need it actually open, you can comment this out
    Set wb = excelApp.Workbooks.Open(newFilePath)
    
    '-----------------------------------------
    ' 4) Re-point each link, preserving sheet+range
    '-----------------------------------------
    Dim objShp As Shape
    Dim oldLink As String
    Dim bangPos As Long
    Dim linkTail As String
    
    For Each objShp In linkedRangeShapes
        
        oldLink = objShp.LinkFormat.SourceFullName
        bangPos = InStr(oldLink, "!")
        
        If bangPos > 0 Then
            ' everything after first "!" is "Sheet1!R1C1:R5C5", etc.
            linkTail = Mid(oldLink, bangPos + 1)
            
            ' new link = newFilePath + "!" + old sheet+range
            objShp.LinkFormat.SourceFullName = newFilePath & "!" & linkTail
            
            ' If toggle is True, set link to Manual
            If setLinksToManual Then
                objShp.LinkFormat.AutoUpdate = ppUpdateOptionManual
            End If
            
            ' If you want all links to be manual, uncomment:
            ' objShp.LinkFormat.AutoUpdate = ppUpdateOptionManual
            
            ' Per-shape update — refreshes each OLE visual from the new Excel source
            objShp.LinkFormat.Update
            
            updatedCount = updatedCount + 1
        End If
    Next objShp
    
    '-----------------------------------------
    ' 5) (Optional) Single pass update at end
    '-----------------------------------------
    ' Bulk UpdateLinks removed — per-shape LinkFormat.Update above already
    ' refreshes each OLE. This second pass was causing a redundant double-refresh.
    ' If doSinglePassUpdate Then
    '     pptPresentation.UpdateLinks
    ' End If
    
    '-----------------------------------------
    ' 6) Close Excel
    '-----------------------------------------
    wb.Close SaveChanges:=False
    excelApp.Quit
    
    Set wb = Nothing
    Set excelApp = Nothing
    
    '-----------------------------------------
    ' 7) Final message
    '-----------------------------------------
    If Not gSilentMode Then
        MsgBox updatedCount & " linked Excel range(s) updated." & vbCrLf & _
           "New path: " & newFilePath & vbCrLf & _
           IIf(doSinglePassUpdate, "(Updated in one bulk pass)", "(No forced update performed)"), _
           vbInformation, "Linked Range Update"
    End If
End Sub


Private Sub CollectLinkedOLERangeShape(ByVal pptShape As Shape, _
                                       ByRef linkedRangeShapes As Collection)
    Dim subShp As Shape
    
    If pptShape.Type = msoGroup Then
        ' Recurse into group items if shape is a group
        For Each subShp In pptShape.GroupItems
            CollectLinkedOLERangeShape subShp, linkedRangeShapes
        Next subShp
        
    ElseIf pptShape.Type = msoLinkedOLEObject Then
        ' Check if LinkFormat is valid
        If Not pptShape.LinkFormat Is Nothing Then
            ' This shape is a linked Excel OLE object (like a pasted range)
            linkedRangeShapes.Add pptShape
        End If
    End If
End Sub

Sub Step_1d_ColorNumbersInTables()

    '=======================================================================
    ' USER‑DEFINED SETTINGS
    '-----------------------------------------------------------------------
    Dim positiveHex As String: positiveHex = "#33CC33"    ' Green for >0
    Dim negativeHex As String: negativeHex = "#ED0590"    ' Red   for <0
    Dim neutralHex  As String: neutralHex  = "#595959"    ' Grey  for 0 / text

    ' Add a prefix (e.g., "+") in front of every positive number/percentage.
    ' Leave this blank ("") if you don’t want any extra symbol.
    Dim positivePrefix As String: positivePrefix = "+"

    ' Remove symbol(s) from the final numeric text **after** colour formatting.
    ' Supply any combination of "+", "-", "%" (e.g. "%+", "%-", "+-", "%+-").
    ' Leave blank ("") to skip removal entirely.
    Dim symbolRemoval As String: symbolRemoval = "%"   ' example removes both % and +
    '=======================================================================

    ' Validate symbolRemoval – only allow the approved characters
    Dim i As Long, ch As String
    For i = 1 To Len(symbolRemoval)
        ch = Mid$(symbolRemoval, i, 1)
        Select Case ch
            Case "+", "-", "%"
                ' accepted chars
            Case Else
                MsgBox "Invalid symbolRemoval setting: only '+', '-' and '%' allowed.", vbCritical
                Exit Sub
        End Select
    Next i

    Dim positiveColor As Long: positiveColor = HexToLong(positiveHex)
    Dim negativeColor As Long: negativeColor = HexToLong(negativeHex)
    Dim neutralColor  As Long: neutralColor  = HexToLong(neutralHex)

    Dim sld As Slide, shp As Shape, tbl As Table
    Dim rowIndex As Long, colIndex As Long
    Dim cellText As String, testVal As String, newText As String
    Dim cellValue As Double, hadPercent As Boolean
    Dim foundShapes As Long, foundTables As Long, cellsColored As Long, totalTables As Long

    For Each sld In ActivePresentation.Slides
        Debug.Print "Processing slide #" & sld.SlideIndex
        foundShapes = 0: foundTables = 0: cellsColored = 0

        For Each shp In sld.Shapes
            If InStr(shp.Name, "_ccst") > 0 Then      ' Check if shape name contains "_ccst" anywhere
                foundShapes = foundShapes + 1

                If shp.HasTable Then
                    foundTables = foundTables + 1: totalTables = totalTables + 1
                    Set tbl = shp.Table

                    For rowIndex = 1 To tbl.Rows.Count
                        For colIndex = 1 To tbl.Columns.Count

                            cellText = Trim$(tbl.Cell(rowIndex, colIndex).Shape.TextFrame.TextRange.Text)
                            testVal = cellText
                            hadPercent = False

                            ' Strip trailing % for numeric test
                            If Right$(testVal, 1) = "%" Then
                                hadPercent = True
                                testVal = Trim$(Left$(testVal, Len(testVal) - 1))
                            End If

                            If IsNumeric(testVal) Then
                                cellValue = CDbl(testVal)

                                ' ------------------ add prefix for positives ------------------
                                If cellValue > 0 And positivePrefix <> "" Then
                                    If Left$(cellText, Len(positivePrefix)) <> positivePrefix Then
                                        If hadPercent Then
                                            newText = positivePrefix & Trim$(testVal) & "%"
                                        Else
                                            newText = positivePrefix & Trim$(testVal)
                                        End If
                                        tbl.Cell(rowIndex, colIndex).Shape.TextFrame.TextRange.Text = newText
                                        cellText = newText
                                    End If
                                End If
                                ' ----------------------------------------------------------------

                                ' Colour based on sign
                                If cellValue > 0 Then
                                    tbl.Cell(rowIndex, colIndex).Shape.TextFrame.TextRange.Font.Color.RGB = positiveColor
                                ElseIf cellValue < 0 Then
                                    tbl.Cell(rowIndex, colIndex).Shape.TextFrame.TextRange.Font.Color.RGB = negativeColor
                                Else
                                    tbl.Cell(rowIndex, colIndex).Shape.TextFrame.TextRange.Font.Color.RGB = neutralColor
                                End If
                                cellsColored = cellsColored + 1

                                ' ------------- optional symbol removal -------------------------
                                If symbolRemoval <> "" Then
                                    ' remove trailing % if requested
                                    If InStr(symbolRemoval, "%") > 0 Then
                                        If Right$(cellText, 1) = "%" Then
                                            cellText = Left$(cellText, Len(cellText) - 1)
                                        End If
                                    End If
                                    ' remove leading + if requested
                                    If InStr(symbolRemoval, "+") > 0 Then
                                        If Left$(cellText, 1) = "+" Then
                                            cellText = Mid$(cellText, 2)
                                        End If
                                    End If
                                    ' remove leading - if requested
                                    If InStr(symbolRemoval, "-") > 0 Then
                                        If Left$(cellText, 1) = "-" Then
                                            cellText = Mid$(cellText, 2)
                                        End If
                                    End If
                                    ' write back any change
                                    tbl.Cell(rowIndex, colIndex).Shape.TextFrame.TextRange.Text = cellText
                                End If
                                ' ----------------------------------------------------------------
                            Else
                                ' Non‑numeric – neutral colour
                                tbl.Cell(rowIndex, colIndex).Shape.TextFrame.TextRange.Font.Color.RGB = neutralColor
                            End If

                        Next colIndex
                    Next rowIndex
                End If
            End If
        Next shp

        Debug.Print "   Found shapes (*_ccst): " & foundShapes
        Debug.Print "   Found tables (*_ccst): " & foundTables
        Debug.Print "   Cells coloured: " & cellsColored
        Debug.Print "--------------"
    Next sld

    gColorCount = totalTables
    If Not gSilentMode Then
        MsgBox totalTables & " table(s) were updated. Check the Immediate Window (Ctrl+G) for details.", vbInformation
    End If

End Sub

'=======================================================================
' Convert a #RRGGBB hex string to a Long color value using VBA's RGB()
'=======================================================================
Private Function HexToLong(ByVal hexColor As String) As Long
    Dim r As Long, g As Long, b As Long
    
    ' Remove '#' if present
    hexColor = Replace(hexColor, "#", "")
    
    ' Safety check: if not 6 hex digits, return black
    If Len(hexColor) <> 6 Then
        HexToLong = RGB(0, 0, 0)
        Exit Function
    End If
    
    ' Extract R, G, B from the hex string
    r = CLng("&H" & Mid(hexColor, 1, 2))
    g = CLng("&H" & Mid(hexColor, 3, 2))
    b = CLng("&H" & Mid(hexColor, 5, 2))
    
    ' Convert to a valid color value
    HexToLong = RGB(r, g, b)
End Function

' ==================================================
'   STEP 1d: APPLY DELTA INDICATOR ARROWS
'   Reads single-cell delt_ objects, determines
'   positive/negative/no change, and swaps the
'   indicator shape from slide-1 templates.
' ==================================================
Sub Step_1c_ApplyDeltaIndicators()

    ' --- Locate the 3 template shapes on Slide 1 ---
    Dim tmplPos As Shape, tmplNeg As Shape, tmplNone As Shape
    Set tmplPos = FindTemplateShape("tmpl_delta_pos")
    Set tmplNeg = FindTemplateShape("tmpl_delta_neg")
    Set tmplNone = FindTemplateShape("tmpl_delta_none")

    If tmplPos Is Nothing Or tmplNeg Is Nothing Or tmplNone Is Nothing Then
        If Not gSilentMode Then
            MsgBox "Missing template shapes on Slide 1." & vbCrLf & _
                   "Expected: tmpl_delta_pos, tmpl_delta_neg, tmpl_delta_none", _
                   vbCritical, "Delta Indicators"
        End If
        Exit Sub
    End If

    ' --- Excel: use shared instance from Step_1b (batch mode) or lazy-init ---
    Dim excelApp As Object
    Dim openedWorkbooks As Object
    Dim ownsExcel As Boolean   ' True if we created our own instance (standalone)

    ' Try shared globals from Step_1b first
    Set excelApp = gExcelApp
    Set openedWorkbooks = gOpenedWorkbooks
    ownsExcel = False

    On Error GoTo DeltaErrorHandler

    ' =====================================================================
    '   TWO-PASS APPROACH
    '   Pass 1: Collect OLE shapes that have a matching delt_ shape
    '           (no shapes are modified — safe for For Each iteration)
    '   Pass 2: Process collected items (delete old delt_, paste new)
    ' =====================================================================

    ' Data structure for collected items
    Dim collSlideIndex() As Long
    Dim collOLESourceFull() As String
    Dim collOLEName() As String
    Dim collDeltName() As String
    Dim collDeltLeft() As Single
    Dim collDeltTop() As Single
    Dim collDeltWidth() As Single
    Dim collDeltHeight() As Single
    Dim collCount As Long
    collCount = 0

    Dim pres As Presentation
    Dim pptSlide As Slide
    Dim pptShape As Shape
    Dim deltShape As Shape
    Dim cntDelta As Long

    Set pres = ActivePresentation

    ' --- PASS 1: Collect ---
    Dim sldIndex As Long
    For sldIndex = 2 To pres.Slides.Count   ' skip slide 1 (templates)
        Set pptSlide = pres.Slides(sldIndex)

        For Each pptShape In pptSlide.Shapes
            If pptShape.Type = msoLinkedOLEObject Then
                If InStr(1, pptShape.OLEFormat.ProgID, "Excel.Sheet") > 0 Then

                    Set deltShape = FindExistingDeltShape(pptSlide, pptShape.Name)

                    If Not deltShape Is Nothing Then
                        ' Grow arrays
                        collCount = collCount + 1
                        ReDim Preserve collSlideIndex(1 To collCount)
                        ReDim Preserve collOLESourceFull(1 To collCount)
                        ReDim Preserve collOLEName(1 To collCount)
                        ReDim Preserve collDeltName(1 To collCount)
                        ReDim Preserve collDeltLeft(1 To collCount)
                        ReDim Preserve collDeltTop(1 To collCount)
                        ReDim Preserve collDeltWidth(1 To collCount)
                        ReDim Preserve collDeltHeight(1 To collCount)

                        collSlideIndex(collCount) = sldIndex
                        collOLESourceFull(collCount) = pptShape.LinkFormat.SourceFullName
                        collOLEName(collCount) = pptShape.Name
                        collDeltName(collCount) = deltShape.Name
                        collDeltLeft(collCount) = deltShape.Left
                        collDeltTop(collCount) = deltShape.Top
                        collDeltWidth(collCount) = deltShape.Width
                        collDeltHeight(collCount) = deltShape.Height

                        Set deltShape = Nothing
                    End If

                End If
            End If
        Next pptShape
    Next sldIndex

    ' --- PASS 2: Process collected items ---
    Dim idx As Long
    Dim cellValue As String, testVal As String
    Dim numVal As Double
    Dim deltaSign As String
    Dim srcTemplate As Shape
    Dim oldDelt As Shape
    Dim newShape As Shape
    Dim gotValue As Boolean

    For idx = 1 To collCount
        Set pptSlide = pres.Slides(collSlideIndex(idx))
        cellValue = ""
        gotValue = False

        ' ---------------------------------------------------------
        ' PRIMARY: Try reading from existing ntbl_/htmp_/trns_ table
        ' (fast — pure PowerPoint, no Excel needed)
        ' ---------------------------------------------------------
        Dim tblShape As Shape
        Set tblShape = FindExistingNtblTable(pptSlide, collOLEName(idx))
        If tblShape Is Nothing Then Set tblShape = FindExistingHtmpTable(pptSlide, collOLEName(idx))
        If tblShape Is Nothing Then Set tblShape = FindExistingTrnsTable(pptSlide, collOLEName(idx))

        If Not tblShape Is Nothing Then
            If tblShape.HasTable Then
                On Error Resume Next
                cellValue = Trim$(tblShape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text)
                If Err.Number = 0 And Len(cellValue) > 0 Then gotValue = True
                Err.Clear
                On Error GoTo DeltaErrorHandler
            End If
        End If

        ' ---------------------------------------------------------
        ' FALLBACK: Read from Excel (delt_-only OLE, no table exists)
        ' ---------------------------------------------------------
        If Not gotValue Then
            Dim filePath As String, sheetName As String, rangeAddress As String
            filePath = ExtractFileName(collOLESourceFull(idx))
            sheetName = ExtractSheetName(collOLESourceFull(idx))
            rangeAddress = ExtractAndConvertRange(collOLESourceFull(idx))

            If rangeAddress = "Not Specified" Or Dir(filePath) = "" Then GoTo NextItem

            ' Lazy-init Excel (only once, only if needed — standalone mode)
            If excelApp Is Nothing Then
                Set excelApp = CreateObject("Excel.Application")
                excelApp.Visible = False
                excelApp.ScreenUpdating = False
                excelApp.EnableEvents = False
                excelApp.DisplayAlerts = False
                Set openedWorkbooks = CreateObject("Scripting.Dictionary")
                ownsExcel = True
            End If

            ' Open or reuse workbook
            Dim excelWorkbook As Object
            If Not openedWorkbooks.Exists(filePath) Then
                On Error Resume Next
                Set excelWorkbook = excelApp.Workbooks.Open(filePath, ReadOnly:=True)
                On Error GoTo DeltaErrorHandler
                If Not excelWorkbook Is Nothing Then
                    openedWorkbooks.Add filePath, excelWorkbook
                End If
            Else
                Set excelWorkbook = openedWorkbooks(filePath)
            End If

            If excelWorkbook Is Nothing Then GoTo NextItem

            On Error Resume Next
            cellValue = Trim$(excelWorkbook.Sheets(sheetName).Range(rangeAddress).Text)
            If Err.Number = 0 And Len(cellValue) > 0 Then gotValue = True
            Err.Clear
            On Error GoTo DeltaErrorHandler
        End If

        If Not gotValue Then GoTo NextItem

        ' --- Determine sign ---
        testVal = cellValue
        If Right$(testVal, 1) = "%" Then testVal = Left$(testVal, Len(testVal) - 1)
        testVal = Trim$(testVal)

        If IsNumeric(testVal) Then
            numVal = CDbl(testVal)
            If numVal > 0 Then
                deltaSign = "pos"
            ElseIf numVal < 0 Then
                deltaSign = "neg"
            Else
                deltaSign = "none"
            End If
        Else
            deltaSign = "none"
        End If

        ' Pick the correct template
        Select Case deltaSign
            Case "pos":  Set srcTemplate = tmplPos
            Case "neg":  Set srcTemplate = tmplNeg
            Case Else:   Set srcTemplate = tmplNone
        End Select

        ' Find and delete the old delt_ shape by name
        ' (it still exists — Pass 1 didn't modify it)
        Set oldDelt = Nothing
        Dim shp As Shape
        For Each shp In pptSlide.Shapes
            If shp.Name = collDeltName(idx) Then
                Set oldDelt = shp
                Exit For
            End If
        Next shp

        If Not oldDelt Is Nothing Then
            oldDelt.Delete
            Set oldDelt = Nothing
        End If

        ' Copy template to this slide
        srcTemplate.Copy
        pptSlide.Shapes.Paste

        ' The pasted shape is the last shape on the slide
        Set newShape = pptSlide.Shapes(pptSlide.Shapes.Count)

        ' Restore position, size & name
        newShape.Left = collDeltLeft(idx)
        newShape.Top = collDeltTop(idx)
        newShape.Width = collDeltWidth(idx)
        newShape.Height = collDeltHeight(idx)
        newShape.Name = collDeltName(idx)

        cntDelta = cntDelta + 1

NextItem:
    Next idx

    ' --- Cleanup: close workbooks, quit Excel, clear globals ---
    GoSub DeltaCleanupExcel

    gDeltaCount = cntDelta
    If Not gSilentMode Then
        MsgBox cntDelta & " delta indicator(s) updated.", vbInformation, "Delta Indicators"
    End If
    Exit Sub

DeltaCleanupExcel:
    On Error Resume Next
    If Not openedWorkbooks Is Nothing Then
        Dim wbKey As Variant
        For Each wbKey In openedWorkbooks.Keys
            openedWorkbooks(wbKey).Close False
        Next wbKey
    End If
    If Not excelApp Is Nothing Then
        excelApp.Quit
    End If
    On Error GoTo 0
    Set excelApp = Nothing
    Set openedWorkbooks = Nothing
    Set gExcelApp = Nothing
    Set gOpenedWorkbooks = Nothing
    Return

DeltaErrorHandler:
    If Not gSilentMode Then
        MsgBox "Delta indicator error: " & Err.Description, vbCritical, "Error"
    End If
    GoSub DeltaCleanupExcel
End Sub

' =========================================================
' MASTER WRAPPER: Step_1_Update_All_Linked_Tables_Combined
' =========================================================
Sub Step_1_RUN_ALL()
    Dim tStart As Single: tStart = Timer
    gSilentMode = True
    gBatchMode = True     ' indicate batch run
    gCancelFlag = False
    gChartsCount = 0

    ' Step 1a: Relink Excel ranges
    Step_1a_UpdateTableLinks
    If gCancelFlag Then
        MsgBox "Operation cancelled. No actions were performed.", vbInformation, "Update"
    gBatchMode = False    ' reset batch flag
        Exit Sub
    End If

    ' Step 1b: Refresh / rebuild tables (keeps Excel alive in batch mode)
    Step_1b_UpdateTableContent

    ' Step 1c: Apply delta indicator arrows (uses shared Excel, then cleans up)
    Step_1c_ApplyDeltaIndicators

    ' Step 1d: Apply +/- colour coding (no Excel needed)
    Step_1d_ColorNumbersInTables

    ' Step 2: Update chart links
    Step_2_UpdateChartLinks

    gSilentMode = False

    Dim msg As String
    msg = "Run complete!" & vbCrLf & _
          gTablesCount & " table(s) refreshed" & vbCrLf & _
          gColorCount & " table(s) color‑formatted" & vbCrLf & _
          gDeltaCount & " delta indicator(s) updated" & vbCrLf & _
          gChartsCount & " chart(s) updated" & vbCrLf & _
          "Elapsed: " & Format(Timer - tStart, "0.00") & " sec"

    gBatchMode = False    ' reset batch flag
    MsgBox msg, vbInformation, "Update Tables & Charts"
End Sub

Sub Step_2_UpdateChartLinks()

    '====================================================================
    '                 TOGGLES / SETTINGS
    '====================================================================
    ' 1) When running **within STEP_1_RUN_ALL** you can reuse the file picked in
    '    Step 1a for chart links. Set True to reuse, False to always show the
    '    picker in the batch run.
    Dim reuseFileInBatch As Boolean
    reuseFileInBatch = True            '<< default for batch runs
    '====================================================================

    Dim excelFilePath As String

    '-------------------------------------------------
    ' 1) Decide which Excel file to use
    '-------------------------------------------------
    If gBatchMode And reuseFileInBatch Then
        ' Running as part of STEP_1_RUN_ALL and reuse is ON
        If Len(gSelectedExcelFile) > 0 And Dir(gSelectedExcelFile) <> "" Then
            excelFilePath = gSelectedExcelFile
        Else
            ' No stored file (or it disappeared) – fall back to dialog
            GoTo PickFile
        End If
    Else
PickFile:
        With Application.FileDialog(msoFileDialogFilePicker)
            .Title = "Select the Excel file for linked charts"
            .Filters.Clear
            .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls", 1
            .AllowMultiSelect = False
            If .Show = -1 Then
                excelFilePath = .SelectedItems(1)
                ' Remember it so subsequent batch runs can reuse it
                gSelectedExcelFile = excelFilePath
            Else
                If Not gSilentMode Then _
                    MsgBox "Operation cancelled. No file selected.", vbInformation
                Exit Sub
            End If
        End With
    End If

    '-------------------------------------------------
    ' 2) Gather linked charts
    '-------------------------------------------------
    Dim pptPresentation As Presentation
    Dim sld As Slide
    Dim pptShape As Shape
    Dim linkedCharts As New Collection
    Dim updatedChartsCount As Long

    Set pptPresentation = ActivePresentation
    For Each sld In pptPresentation.Slides
        For Each pptShape In sld.Shapes
            CollectLinkedCharts pptShape, linkedCharts
        Next pptShape
    Next sld

    If linkedCharts.Count = 0 Then
        If Not gSilentMode Then _
            MsgBox "No linked charts found in this presentation.", vbInformation, "No Linked Charts"
        Exit Sub
    End If

    '-------------------------------------------------
    ' 3) Open workbook and update charts
    '-------------------------------------------------
    Dim excelApp As Object
    Dim workbook As Object

    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    Set workbook = excelApp.Workbooks.Open(excelFilePath)

    Dim chartShape As Shape
    For Each chartShape In linkedCharts
        chartShape.LinkFormat.SourceFullName = excelFilePath
        chartShape.LinkFormat.Update
        updatedChartsCount = updatedChartsCount + 1
    Next chartShape

    workbook.Close False
    excelApp.Quit
    Set workbook = Nothing
    Set excelApp = Nothing

    gChartsCount = updatedChartsCount
    If Not gSilentMode Then _
        MsgBox updatedChartsCount & " linked chart(s) have been updated.", _
             vbInformation, "Charts Update Summary"

End Sub

Sub CollectLinkedCharts(pptShape As Shape, linkedCharts As Collection)
    Dim groupedShape As Shape
    If pptShape.Type = msoGroup Then
        ' Loop through grouped shapes and collect linked charts
        For Each groupedShape In pptShape.GroupItems
            CollectLinkedCharts groupedShape, linkedCharts
        Next groupedShape
    ElseIf pptShape.HasChart Then
        If pptShape.Chart.ChartData.IsLinked Then
            ' Add linked chart to the collection
            linkedCharts.Add pptShape
        End If
    End If
End Sub
