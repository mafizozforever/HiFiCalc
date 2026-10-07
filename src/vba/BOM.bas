Attribute VB_Name = "BOM"
Option Explicit

' Builds the eight-column spill range used by "Ведомомость закупки".
' Source must contain DATA_Ведомости!A2:O40 through the PurchaseSource name.
' Numeric indexes refer to DATA columns A:O; inline letters are legacy labels.
' The function is intentionally pure: it reads Source and returns an array;
' it never selects, activates, or writes cells.
Public Function PurchaseSummary(ByVal Source As Range) As Variant
    Const COL_ENABLED As Long = 1       ' U  - Включить
    Const COL_COMPONENT As Long = 5     ' Y  - Компонент
    Const COL_NOMINAL As Long = 6       ' Z  - Номинал
    Const COL_UNIT As Long = 7          ' AA - Ед.
    Const COL_QUANTITY As Long = 8      ' AB - Кол-во
    Const COL_RATING As Long = 9        ' AC - Рейтинг
    Const COL_WIRE_DIAMETER As Long = 10 ' AD - Диаметр провода
    Const COL_TURNS As Long = 11        ' AE - Витков
    Const COL_BOBBIN As Long = 12       ' AF - Каркас
    Const COL_COMPONENT_KEY As Long = 14 ' AH - Составной ключ
    Const COL_WIRE_LENGTH As Long = 15  ' AJ - Общий метраж
    Const REQUIRED_COLUMNS As Long = 15

    On Error GoTo CalculationError

    If Source Is Nothing Then
        PurchaseSummary = CVErr(xlErrRef)
        Exit Function
    End If

    If Source.Areas.Count <> 1 Or Source.Columns.Count < REQUIRED_COLUMNS Then
        PurchaseSummary = CVErr(xlErrRef)
        Exit Function
    End If

    Dim sourceValues As Variant
    sourceValues = Source.Value2

    Dim components As Object
    Set components = CreateObject("Scripting.Dictionary")
    components.CompareMode = vbBinaryCompare

    Dim wireDiameters As Object
    Set wireDiameters = CreateObject("Scripting.Dictionary")
    wireDiameters.CompareMode = vbBinaryCompare

    Dim rowIndex As Long
    Dim componentKey As String
    Dim part As Variant
    Dim diameter As Double
    Dim diameterKey As String

    ' First pass: aggregate components and collect enabled wire diameters.
    For rowIndex = 1 To UBound(sourceValues, 1)
        AssertNotError sourceValues(rowIndex, COL_COMPONENT_KEY)

        componentKey = CStr(sourceValues(rowIndex, COL_COMPONENT_KEY))
        If Len(componentKey) > 0 Then
            If components.Exists(componentKey) Then
                part = components(componentKey)
            Else
                AssertNotError sourceValues(rowIndex, COL_COMPONENT)
                AssertNotError sourceValues(rowIndex, COL_NOMINAL)
                AssertNotError sourceValues(rowIndex, COL_UNIT)
                AssertNotError sourceValues(rowIndex, COL_RATING)
                AssertNotError sourceValues(rowIndex, COL_WIRE_DIAMETER)
                AssertNotError sourceValues(rowIndex, COL_TURNS)
                AssertNotError sourceValues(rowIndex, COL_BOBBIN)

                part = Array( _
                    sourceValues(rowIndex, COL_COMPONENT), _
                    sourceValues(rowIndex, COL_NOMINAL), _
                    sourceValues(rowIndex, COL_UNIT), _
                    ZeroToBlank(sourceValues(rowIndex, COL_RATING)), _
                    0#, _
                    ZeroToBlank(sourceValues(rowIndex, COL_WIRE_DIAMETER)), _
                    ZeroToBlank(sourceValues(rowIndex, COL_TURNS)), _
                    ZeroToBlank(sourceValues(rowIndex, COL_BOBBIN)))
            End If

            part(4) = AddNumeric(part(4), sourceValues(rowIndex, COL_QUANTITY))
            components(componentKey) = part
        End If

        AssertNotError sourceValues(rowIndex, COL_ENABLED)
        AssertNotError sourceValues(rowIndex, COL_WIRE_DIAMETER)
        If IsEnabled(sourceValues(rowIndex, COL_ENABLED)) _
                And IsUsableNumber(sourceValues(rowIndex, COL_WIRE_DIAMETER)) Then
            diameter = CDbl(sourceValues(rowIndex, COL_WIRE_DIAMETER))
            diameterKey = NumericKey(diameter)
            If Not wireDiameters.Exists(diameterKey) Then
                wireDiameters.Add diameterKey, Array(diameter, 0#)
            End If
        End If
    Next rowIndex

    ' Second pass mirrors SUMIF(AD, diameter, AJ) from the original formula.
    For rowIndex = 1 To UBound(sourceValues, 1)
        AssertNotError sourceValues(rowIndex, COL_WIRE_DIAMETER)
        If IsUsableNumber(sourceValues(rowIndex, COL_WIRE_DIAMETER)) Then
            diameter = CDbl(sourceValues(rowIndex, COL_WIRE_DIAMETER))
            diameterKey = NumericKey(diameter)
            If wireDiameters.Exists(diameterKey) Then
                part = wireDiameters(diameterKey)
                part(1) = AddNumeric(part(1), sourceValues(rowIndex, COL_WIRE_LENGTH))
                wireDiameters(diameterKey) = part
            End If
        End If
    Next rowIndex

    If components.Count = 0 And wireDiameters.Count = 0 Then
        PurchaseSummary = "Нет выбранных компонентов"
        Exit Function
    End If

    Dim componentRows As Variant
    componentRows = BuildComponentRows(components)
    If components.Count > 1 Then SortRows componentRows

    Dim wireRows As Variant
    wireRows = BuildWireRows(wireDiameters)
    If wireDiameters.Count > 1 Then SortRows wireRows

    PurchaseSummary = CombineRows(componentRows, components.Count, _
                                  wireRows, wireDiameters.Count)
    Exit Function

CalculationError:
    ' Do not disguise damaged source data as an empty component selection.
    PurchaseSummary = CVErr(xlErrValue)
End Function

Private Function BuildComponentRows(ByVal components As Object) As Variant
    If components.Count = 0 Then
        BuildComponentRows = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To components.Count, 1 To 8)

    Dim keys As Variant
    keys = components.keys

    Dim index As Long
    Dim columnIndex As Long
    Dim part As Variant
    For index = 0 To components.Count - 1
        part = components(keys(index))
        For columnIndex = 0 To 7
            result(index + 1, columnIndex + 1) = part(columnIndex)
        Next columnIndex
    Next index

    BuildComponentRows = result
End Function

Private Function BuildWireRows(ByVal wireDiameters As Object) As Variant
    If wireDiameters.Count = 0 Then
        BuildWireRows = Empty
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To wireDiameters.Count, 1 To 8)

    Dim keys As Variant
    keys = wireDiameters.keys

    Dim index As Long
    Dim wireData As Variant
    For index = 0 To wireDiameters.Count - 1
        wireData = wireDiameters(keys(index))
        result(index + 1, 1) = "Провод эмалированный"
        result(index + 1, 2) = wireData(0)
        result(index + 1, 3) = "мм"
        result(index + 1, 4) = "Общий метраж, м"
        result(index + 1, 5) = Application.WorksheetFunction.Round(wireData(1), 1)
        result(index + 1, 6) = wireData(0)
        result(index + 1, 7) = vbNullString
        result(index + 1, 8) = vbNullString
    Next index

    BuildWireRows = result
End Function

Private Function CombineRows(ByVal componentRows As Variant, _
                             ByVal componentCount As Long, _
                             ByVal wireRows As Variant, _
                             ByVal wireCount As Long) As Variant
    Dim result() As Variant
    ReDim result(1 To componentCount + wireCount, 1 To 8)

    Dim outputRow As Long
    Dim sourceRow As Long
    Dim columnIndex As Long

    outputRow = 1
    For sourceRow = 1 To componentCount
        For columnIndex = 1 To 8
            result(outputRow, columnIndex) = componentRows(sourceRow, columnIndex)
        Next columnIndex
        outputRow = outputRow + 1
    Next sourceRow

    For sourceRow = 1 To wireCount
        For columnIndex = 1 To 8
            result(outputRow, columnIndex) = wireRows(sourceRow, columnIndex)
        Next columnIndex
        outputRow = outputRow + 1
    Next sourceRow

    CombineRows = result
End Function

' Stable insertion sort: component text first, numeric nominal second.
Private Sub SortRows(ByRef rows As Variant)
    Dim rowIndex As Long
    Dim compareIndex As Long

    For rowIndex = 2 To UBound(rows, 1)
        compareIndex = rowIndex
        Do While compareIndex > 1
            If CompareRows(rows, compareIndex - 1, compareIndex) <= 0 Then Exit Do
            SwapRows rows, compareIndex - 1, compareIndex
            compareIndex = compareIndex - 1
        Loop
    Next rowIndex
End Sub

Private Function CompareRows(ByRef rows As Variant, _
                             ByVal leftRow As Long, _
                             ByVal rightRow As Long) As Long
    CompareRows = StrComp(CStr(rows(leftRow, 1)), _
                          CStr(rows(rightRow, 1)), vbTextCompare)
    If CompareRows <> 0 Then Exit Function

    CompareRows = CompareValues(rows(leftRow, 2), rows(rightRow, 2))
End Function

Private Function CompareValues(ByVal leftValue As Variant, _
                               ByVal rightValue As Variant) As Long
    If IsUsableNumber(leftValue) And IsUsableNumber(rightValue) Then
        If CDbl(leftValue) < CDbl(rightValue) Then
            CompareValues = -1
        ElseIf CDbl(leftValue) > CDbl(rightValue) Then
            CompareValues = 1
        Else
            CompareValues = 0
        End If
    Else
        CompareValues = StrComp(CStr(leftValue), CStr(rightValue), vbTextCompare)
    End If
End Function

Private Sub SwapRows(ByRef rows As Variant, _
                     ByVal firstRow As Long, _
                     ByVal secondRow As Long)
    Dim columnIndex As Long
    Dim temporary As Variant

    For columnIndex = 1 To 8
        temporary = rows(firstRow, columnIndex)
        rows(firstRow, columnIndex) = rows(secondRow, columnIndex)
        rows(secondRow, columnIndex) = temporary
    Next columnIndex
End Sub

Private Function ZeroToBlank(ByVal value As Variant) As Variant
    AssertNotError value

    If IsEmpty(value) Then
        ZeroToBlank = vbNullString
    ElseIf VarType(value) = vbString And Len(value) = 0 Then
        ZeroToBlank = vbNullString
    ElseIf IsNumeric(value) Then
        If CDbl(value) = 0 Then
            ZeroToBlank = vbNullString
        Else
            ZeroToBlank = value
        End If
    Else
        ZeroToBlank = value
    End If
End Function

Private Function AddNumeric(ByVal currentValue As Variant, _
                            ByVal valueToAdd As Variant) As Double
    AssertNotError valueToAdd

    AddNumeric = CDbl(currentValue)
    If IsUsableNumber(valueToAdd) Then
        AddNumeric = AddNumeric + CDbl(valueToAdd)
    End If
End Function

Private Function IsEnabled(ByVal value As Variant) As Boolean
    If IsUsableNumber(value) Then IsEnabled = (CDbl(value) = 1#)
End Function

Private Function IsUsableNumber(ByVal value As Variant) As Boolean
    If IsError(value) Then Exit Function
    If IsEmpty(value) Then Exit Function
    If VarType(value) = vbString And Len(value) = 0 Then Exit Function
    IsUsableNumber = IsNumeric(value)
End Function

Private Function NumericKey(ByVal value As Double) As String
    NumericKey = CStr(value)
End Function

Private Sub AssertNotError(ByVal value As Variant)
    If IsError(value) Then
        Err.Raise vbObjectError + 2048, "PurchaseSummary", _
                  "Source range contains a calculation error."
    End If
End Sub


