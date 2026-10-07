Attribute VB_Name = "OptimalPairE24"
Option Explicit

Function GetPairedE24(Val1 As Double, Val2 As Double, ReturnIndex As Integer) As Double
    Dim E24 As Variant
    E24 = Array(1, 1.1, 1.2, 1.3, 1.5, 1.6, 1.8, 2, 2.2, 2.4, 2.7, 3, 3.3, 3.6, 3.9, 4.3, 4.7, 5.1, 5.6, 6.2, 6.8, 7.5, 8.2, 9.1, 10)

    If ReturnIndex < 1 Or ReturnIndex > 2 Then
        GetPairedE24 = 0
        Exit Function
    End If
    
    ' --- 1. БЛОК ЗАЩИТЫ ОТ НУЛЕЙ (Если компонента нет в схеме) ---
    
    ' Если вообще нет обоих компонентов
    If Val1 <= 0 And Val2 <= 0 Then
        GetPairedE24 = 0
        Exit Function
    End If
    
    ' Если нет второго компонента (С2 = 0, как в фильтрах 1, 2 и 3 порядков)
    If Val2 <= 0 Then
        If ReturnIndex = 2 Then
            GetPairedE24 = 0 ' Возвращаем 0 для пустой ячейки
            Exit Function
        Else
            ' Ищем обычное ближайшее значение только для Val1
            Dim V1_Low As Double, V1_High As Double
            V1_Low = GetNearestE24(Val1, -1, E24)
            V1_High = GetNearestE24(Val1, 1, E24)
            If (Val1 - V1_Low) <= (V1_High - Val1) Then
                GetPairedE24 = V1_Low
            Else
                GetPairedE24 = V1_High
            End If
            Exit Function
        End If
    End If
    
    ' Если нет первого компонента (для симметрии логики)
    If Val1 <= 0 Then
        If ReturnIndex = 1 Then
            GetPairedE24 = 0
            Exit Function
        Else
            Dim V2_Low As Double, V2_High As Double
            V2_Low = GetNearestE24(Val2, -1, E24)
            V2_High = GetNearestE24(Val2, 1, E24)
            If (Val2 - V2_Low) <= (V2_High - Val2) Then
                GetPairedE24 = V2_Low
            Else
                GetPairedE24 = V2_High
            End If
            Exit Function
        End If
    End If
    ' -------------------------------------------------------------
    
    ' --- 2. ОСНОВНОЙ АЛГОРИТМ ДЛЯ ПАРЫ (Если оба компонента > 0) ---
    Dim V1_L As Double, V1_H As Double
    Dim V2_L As Double, V2_H As Double
    
    V1_L = GetNearestE24(Val1, -1, E24)
    V1_H = GetNearestE24(Val1, 1, E24)
    V2_L = GetNearestE24(Val2, -1, E24)
    V2_H = GetNearestE24(Val2, 1, E24)
    
    ' Массив из 4 комбинаций (Низ/Низ, Низ/Верх, Верх/Низ, Верх/Верх)
    Dim Pairs(1 To 4, 1 To 2) As Double
    Pairs(1, 1) = V1_L: Pairs(1, 2) = V2_L
    Pairs(2, 1) = V1_L: Pairs(2, 2) = V2_H
    Pairs(3, 1) = V1_H: Pairs(3, 2) = V2_L
    Pairs(4, 1) = V1_H: Pairs(4, 2) = V2_H
    
    Dim MinSpread As Double
    Dim BestPair As Integer
    Dim CurrentSpread As Double
    Dim i As Integer
    
    MinSpread = 999999
    BestPair = 1
    
For i = 1 To 4
        Dim Ratio1 As Double, Ratio2 As Double
        Dim Spread As Double, AbsDev As Double
        
        Ratio1 = Pairs(i, 1) / Val1
        Ratio2 = Pairs(i, 2) / Val2
        
        ' 1. Дельта разброса (насколько синхронно они отклоняются - важно для формы фильтра)
        Spread = Abs(Ratio1 - Ratio2)
        
        ' 2. Сумма абсолютных отклонений (насколько далеко они улетели от оригинала - важно для частоты среза)
        AbsDev = Abs(1 - Ratio1) + Abs(1 - Ratio2)
        
        ' 3. Комплексный штрафной балл.
        ' Пропорция в 3 раза важнее, но абсолютное отклонение работает как балансир
        CurrentSpread = (Spread * 3) + AbsDev
        
        If CurrentSpread < MinSpread Then
            MinSpread = CurrentSpread
            BestPair = i
        End If
    Next i
    
    GetPairedE24 = Pairs(BestPair, ReturnIndex)
End Function

' Вспомогательная функция для поиска ближайшего E24 с заданным направлением (-1 = вниз, 1 = вверх)
Private Function GetNearestE24(Target As Double, Direction As Integer, E24Array As Variant) As Double
    Dim Multiplier As Double
    Dim NormTarget As Double
    Dim i As Integer
    Dim BestVal As Double
    
    If Target <= 0 Then
        GetNearestE24 = 0
        Exit Function
    End If
    
    Multiplier = 10 ^ Int(Application.WorksheetFunction.Log10(Target))
    NormTarget = Target / Multiplier
    
    If NormTarget < 1 Then
        NormTarget = NormTarget * 10
        Multiplier = Multiplier / 10
    End If
    
    ' Ищем значение по направлению
    If Direction = -1 Then
        BestVal = E24Array(0)
        For i = 0 To UBound(E24Array)
            If E24Array(i) <= NormTarget Then BestVal = E24Array(i)
        Next i
    Else
        BestVal = E24Array(UBound(E24Array))
        For i = UBound(E24Array) To 0 Step -1
            If E24Array(i) >= NormTarget Then BestVal = E24Array(i)
        Next i
    End If
    
    GetNearestE24 = BestVal * Multiplier
End Function

