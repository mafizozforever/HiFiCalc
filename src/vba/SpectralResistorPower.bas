Attribute VB_Name = "SpectralResistorPower"
Option Explicit


' Главная функция для вычисления мощности резистора (основной нагрев идет на последовательном R1 в L-pad)
Function CalcResistorPower(AmpPower As Double, FreqLow As Double, FreqHigh As Double, AttenuationDB As Double) As Variant
    
    ' Если ослабления нет, резистор не нужен
    If AttenuationDB <= 0 Then
        CalcResistorPower = 0
        Exit Function
    End If
    
    ' 1. Вычисляем долю мощности для заданной полосы частот (СЧ, ВЧ или НЧ)
    Dim K_low As Double, K_high As Double, K_band As Double
    K_low = GetSpectralCoeff(FreqLow)
    K_high = GetSpectralCoeff(FreqHigh)
    K_band = K_high - K_low
    If K_band < 0 Then K_band = 0
    
    ' 2. Вычисляем мощность, выделяемую на последовательном резисторе R1
    Dim K_vol As Double
    Dim P_in As Double
    Dim P_resistor As Double
    
    ' Коэффициент ослабления по напряжению
    K_vol = 10 ^ (AttenuationDB / 20)
    ' Мощность, пришедшая в полосу динамика
    P_in = AmpPower * K_band
    ' Тепловая мощность на резисторе R1
    P_resistor = P_in * ((K_vol - 1) / K_vol)
    
    ' 3. Округляем до ближайшего стандартного номинала в большую сторону
    Dim StdPower As Variant
    Dim i As Integer
    ' Массив стандартных мощностей резисторов (в Ваттах)
    StdPower = Array(5, 10, 15, 20, 25, 30, 50)
    
    For i = LBound(StdPower) To UBound(StdPower)
        If P_resistor <= StdPower(i) Then
            CalcResistorPower = StdPower(i)
            Exit Function
        End If
    Next i
    
    ' Если расчетная мощность превысила 50 Вт
    CalcResistorPower = "Перегрев: " & Round(P_resistor, 1) & " Вт"
    
End Function

' Вспомогательная функция: линейная интерполяция спектрального коэффициента
Private Function GetSpectralCoeff(Freq As Double) As Double
    Dim Freqs As Variant
    Dim Coeffs As Variant
    Dim i As Integer
    
    ' Опорные точки частот (в Гц) и их кумулятивные коэффициенты
    Freqs = Array(0, 100, 250, 400, 500, 800, 1000, 1500, 2000, 3000, 5000, 8000, 10000, 20000)
    Coeffs = Array(0, 0.2, 0.4, 0.5, 0.55, 0.65, 0.7, 0.75, 0.8, 0.85, 0.9, 0.95, 0.97, 1)
    
    ' Обработка крайних значений
    If Freq <= 0 Then
        GetSpectralCoeff = 0
        Exit Function
    ElseIf Freq >= 20000 Then
        GetSpectralCoeff = 1
        Exit Function
    End If
    
    ' Поиск диапазона и расчет пропорции
    For i = 0 To UBound(Freqs) - 1
        If Freq >= Freqs(i) And Freq <= Freqs(i + 1) Then
            GetSpectralCoeff = Coeffs(i) + (Freq - Freqs(i)) * (Coeffs(i + 1) - Coeffs(i)) / (Freqs(i + 1) - Freqs(i))
            Exit Function
        End If
    Next i
End Function

Function CalcRealTemp(FullTemp As Double, FreqLow As Double, FreqHigh As Double) As Double
    Dim K_band As Double
    
    ' Вычисляем долю спектральной энергии
    K_band = GetSpectralCoeff(FreqHigh) - GetSpectralCoeff(FreqLow)
    If K_band < 0 Then K_band = 0
    
    ' Считаем температуру
    CalcRealTemp = 25 + (FullTemp - 25) * K_band
    
    ' Устанавливаем базовый минимум в 25 градусов
    If CalcRealTemp < 25 Then CalcRealTemp = 25
    
End Function

Function CalcBscResistorPower(AmpPower As Double, FreqBSC As Double, FreqHigh As Double, ResistorValue As Double, SpeakerZ As Double) As Variant
    
    Dim K_bsc As Double, K_high As Double, K_band As Double
    Dim P_band As Double
    Dim P_resistor As Double
    
    ' Защита от деления на ноль и нулевых значений
    If SpeakerZ <= 0 Or ResistorValue <= 0 Then
        CalcBscResistorPower = 0
        Exit Function
    End If
    
    ' 1. Вычисляем долю мощности для частот ВЫШЕ бафл-степа (от F_bsc до F_high)
    K_bsc = GetSpectralCoeff(FreqBSC)
    K_high = GetSpectralCoeff(FreqHigh)
    K_band = K_high - K_bsc
    If K_band < 0 Then K_band = 0
    
    ' 2. Энергия, попадающая в этот частотный диапазон
    P_band = AmpPower * K_band
    
    ' 3. Вычисляем реальную тепловую мощность на резисторе BSC
    P_resistor = P_band * ((SpeakerZ * ResistorValue) / ((SpeakerZ + ResistorValue) ^ 2))
    
    ' 4. Округляем до ближайшего стандартного номинала в большую сторону
    Dim StdPower As Variant
    Dim i As Integer
    StdPower = Array(5, 10, 15, 20, 25, 30, 50)
    
    For i = LBound(StdPower) To UBound(StdPower)
        If P_resistor <= StdPower(i) Then
            CalcBscResistorPower = StdPower(i)
            Exit Function
        End If
    Next i
    
    ' Если расчетная мощность превысила 50 Вт
    CalcBscResistorPower = "Перегрев: " & Round(P_resistor, 1) & " Вт"
    
End Function

Function CalcParallelResistorPower(AmpPower As Double, FilterType As String, TargetFreq As Double, FreqLow As Double, FreqHigh As Double, Optional Bandwidth As Double = 200) As Variant
    
    Dim P_resistor As Double
    Dim K_low As Double, K_high As Double
    Dim C_low As Double, C_high As Double
    Dim I_low As Double, I_high As Double
    
    ' 1. Определяем рабочую полосу самого параллельного компонента (Цобеля или Режектора)
    Select Case UCase(Trim(FilterType))
        
        Case "ZOBEL", "ЦОБЕЛЬ"
            C_low = TargetFreq
            C_high = 20000 ' Верхняя граница звукового диапазона
            
        Case "NOTCH", "РЕЖЕКТОР"
            C_low = TargetFreq - (Bandwidth / 2)
            If C_low < 0 Then C_low = 0
            C_high = TargetFreq + (Bandwidth / 2)
            
        Case Else
            CalcParallelResistorPower = "Ошибка: Укажите 'Zobel' или 'Notch'"
            Exit Function
            
    End Select
    
    ' 2. Находим ПЕРЕСЕЧЕНИЕ диапазонов (Окно компонента + Окно динамика после кроссовера)
    ' Используем математическую логику: Max для нижних границ, Min для верхних
    If C_low > FreqLow Then I_low = C_low Else I_low = FreqLow
    If C_high < FreqHigh Then I_high = C_high Else I_high = FreqHigh
    
    ' 3. Расчет мощности только в зоне пересечения
    If I_low >= I_high Then
        ' Диапазоны не пересекаются! Сигнал сюда вообще не доходит (отрезан кроссовером)
        P_resistor = 0
    Else
        ' Берем долю энергии из матрицы только для зоны наложения
        K_low = GetSpectralCoeff(I_low)
        K_high = GetSpectralCoeff(I_high)
        
        If UCase(Trim(FilterType)) = "ZOBEL" Or UCase(Trim(FilterType)) = "ЦОБЕЛЬ" Then
            ' В Цобеле ток делится пополам с динамиком
            P_resistor = AmpPower * (K_high - K_low) * 0.5
        Else
            ' Для Режектора берем всю энергию пересекающейся полосы
            P_resistor = AmpPower * (K_high - K_low)
        End If
    End If
    
    ' 4. Выдаем стандартный номинал (минимум 5 Вт, даже если P_resistor = 0)
    Dim StdPower As Variant
    Dim i As Integer
    StdPower = Array(5, 10, 15, 20, 25, 30, 50)
    
    For i = LBound(StdPower) To UBound(StdPower)
        If P_resistor <= StdPower(i) Then
            CalcParallelResistorPower = StdPower(i)
            Exit Function
        End If
    Next i
    
    CalcParallelResistorPower = "Перегрев: " & Round(P_resistor, 1) & " Вт"
    
End Function

