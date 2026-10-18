Attribute VB_Name = "OptimalBobbin"
Option Explicit

Function OPTIMAL_BOBBIN(TargetL_mH As Variant, WireD_Cu As Variant, AmpPower As Variant, _
                        Impedance As Variant, FreqLow As Variant, FreqHigh As Variant) As String
    
    On Error GoTo ErrorHandler
    
    ' Конвертация в числа
    Dim L As Double: L = SafeToDouble(TargetL_mH)
    Dim d_cu As Double: d_cu = SafeToDouble(WireD_Cu)
    Dim P As Double: P = SafeToDouble(AmpPower)
    Dim Z As Double: Z = SafeToDouble(Impedance)
    Dim FL As Double: FL = SafeToDouble(FreqLow)
    Dim FH As Double: FH = SafeToDouble(FreqHigh)
    
    If L <= 0 Or d_cu <= 0 Then
        OPTIMAL_BOBBIN = ""
        Exit Function
    End If
    
    ' Настройки оптимизатора
    Const WINDING_SWELL As Double = 1.1 ' +10% на неплотность намотки (Оставляем твой запас)
    Const SAFETY_GAP As Double = 1#     ' Зазор до края щечки (мм) (Оставляем твой запас)
    Const MAX_TEMP As Double = 80#      ' Макс. температура (°C)
    
    Dim wsData As Worksheet
    On Error Resume Next
    Set wsData = Лист8
    On Error GoTo ErrorHandler
    If wsData Is Nothing Then
        OPTIMAL_BOBBIN = "ERR: Нет листа Data"
        Exit Function
    End If
    
    Dim LastWire As Long, LastBobbin As Long
    LastWire = wsData.Cells(wsData.rows.Count, 1).End(xlUp).Row
    LastBobbin = wsData.Cells(wsData.rows.Count, 8).End(xlUp).Row
    
    ' --- 1. ИЩЕМ ИЗОЛЯЦИЮ ДЛЯ ВЫБРАННОГО ПРОВОДА ---
    Dim d_ins As Double, FoundWire As Boolean
    Dim w As Long
    For w = 2 To LastWire
        If Abs(SafeToDouble(wsData.Cells(w, 1).value) - d_cu) < 0.001 Then
            d_ins = SafeToDouble(wsData.Cells(w, 2).value)
            FoundWire = True
            Exit For
        End If
    Next w
    
    If Not FoundWire Then
        OPTIMAL_BOBBIN = "Нет провода в базе"
        Exit Function
    End If

    ' --- 2. ПЕРЕБОР КАРКАСОВ ---
    Dim BestBobbin As String
    Dim MinScore As Double: MinScore = 1E+15
    Dim ValidFound As Boolean: ValidFound = False
    
    Dim b As Long
    For b = 2 To LastBobbin
        Dim b_name As String: b_name = CStr(wsData.Cells(b, 8).value)
        Dim b_din As Double: b_din = SafeToDouble(wsData.Cells(b, 9).value)
        Dim b_w As Double: b_w = SafeToDouble(wsData.Cells(b, 10).value)
        Dim b_maxD As Double: b_maxD = SafeToDouble(wsData.Cells(b, 11).value)
        
        If Len(b_name) > 0 And b_din > 0 And b_w > 0 And b_maxD > b_din Then
            
            Dim AvailableH As Double
            AvailableH = ((b_maxD - b_din) / 2) - SAFETY_GAP
            
            If AvailableH > 0 Then
                
                Dim Res As Variant
                Res = CoilCalc(L, d_cu, d_ins, b_din, b_w)
                Dim Calc_R As Double: Calc_R = Res(2)
                Dim Calc_Thick As Double: Calc_Thick = Res(3)
                Dim RealThick As Double: RealThick = Calc_Thick * WINDING_SWELL
                
                ' Влезает ли физически?
                If Calc_R > 0 And RealThick <= AvailableH Then
                    
                    ' ТЕПЛОВОЙ РАСЧЕТ
                    Dim HeatW As Double: HeatW = 0
                    If Z > 0 Then HeatW = (P / Z) * Calc_R
                    
                    Dim D_out As Double: D_out = b_din + 2 * Calc_Thick
                    Dim Area_Cyl As Double: Area_Cyl = 3.1416 * D_out * b_w
                    Dim Area_Sides As Double: Area_Sides = 2 * 3.1416 * ((D_out / 2) ^ 2 - (b_din / 2) ^ 2)
                    Dim TotalArea_m2 As Double: TotalArea_m2 = (Area_Cyl + Area_Sides) / 1000000
                    
                    Dim T_max As Double
                    If TotalArea_m2 > 0 Then T_max = 25 + (HeatW / (TotalArea_m2 * 10)) Else T_max = 25
                    
                    Dim T_spec As Double
                    T_spec = CalcRealTemp(T_max, FL, FH)
                    
                    ' Если не сгорит
                    If T_spec <= MAX_TEMP Then
                        ValidFound = True
                        
                        ' === УМНЫЙ КРИТЕРИЙ (ШТРАФ ЗА ФОРМУ БРУКСА) ===
                        ' Идеальная катушка имеет сечение квадрата (Толщина намотки = Ширине каркаса)
                        ' Чем сильнее пропорции отклоняются от 1:1, тем больше штрафной коэффициент
                        Dim ShapePenalty As Double
                        ShapePenalty = 1 + Abs(1 - (Calc_Thick / b_w))
                        
                        Dim Score As Double
                        ' Умножаем DCR на штраф. Это отсечет огромные полупустые каркасы!
                        Score = Calc_R * ShapePenalty
                        
                        If Score < MinScore Then
                            MinScore = Score
                            BestBobbin = b_name
                        End If
                        
                    End If
                End If
            End If
        End If
    Next b
    
    If ValidFound Then
        OPTIMAL_BOBBIN = BestBobbin
    Else
        ' Если провод слишком тонкий и греется, или слишком толстый и не влезает никуда
        OPTIMAL_BOBBIN = "Перегрев / Не влез"
    End If
    Exit Function

ErrorHandler:
    OPTIMAL_BOBBIN = "ERR: " & Err.Description
End Function

' ==========================================
' ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ (ЗАЩИТА ОТ ТЕКСТА/ЗАПЯТЫХ)
' ==========================================
Function SafeToDouble(val As Variant) As Double
    On Error Resume Next
    If IsNumeric(val) Then
        SafeToDouble = CDbl(val)
    Else
        Dim s As String
        s = CStr(val)
        s = Replace(s, ".", ",")
        If IsNumeric(s) Then
            SafeToDouble = CDbl(s)
        Else
            s = Replace(s, ",", ".")
            If IsNumeric(s) Then
                SafeToDouble = VBA.val(s)
            Else
                SafeToDouble = 0
            End If
        End If
    End If
    On Error GoTo 0
End Function

' ==========================================
' ФУНКЦИЯ РАСЧЕТА КАТУШКИ (ИЗНАЧАЛЬНАЯ)
' ==========================================
Function CoilCalc(TargetL_mH As Double, WireD_Cu As Double, WireD_Ins As Double, _
                  BobbinD_In As Double, BobbinH As Double) As Variant
    
    Dim N As Long, Layers As Long, TurnsPerLayer As Long
    Dim CurrentL_uH As Double, TargetL_uH As Double
    Dim R_avg_mm As Double, Thickness_mm As Double
    Dim R_avg_in As Double, H_in As Double, C_in As Double
    
    Const Pi As Double = 3.14159265359
    Const mmToInch As Double = 0.0393701
    
    If BobbinH <= 0 Or WireD_Ins <= 0 Or TargetL_mH <= 0 Then
        CoilCalc = Array(0, 0, 0, 0)
        Exit Function
    End If

    TargetL_uH = TargetL_mH * 1000
    TurnsPerLayer = Int(BobbinH / WireD_Ins)
    
    If TurnsPerLayer < 1 Then
        CoilCalc = Array(0, 0, 0, 0)
        Exit Function
    End If
    
    N = 0
    CurrentL_uH = 0
    Dim SafetyCounter As Long: SafetyCounter = 0
    
    Do While CurrentL_uH < TargetL_uH
        N = N + 1
        SafetyCounter = SafetyCounter + 1
        
        Layers = Int((N - 1) / TurnsPerLayer) + 1
        Thickness_mm = Layers * WireD_Ins
        R_avg_mm = (BobbinD_In / 2) + (Thickness_mm / 2)
        
        R_avg_in = R_avg_mm * mmToInch
        H_in = BobbinH * mmToInch
        C_in = Thickness_mm * mmToInch
        
        Dim Denominator As Double
        Denominator = (6 * R_avg_in + 9 * H_in + 10 * C_in)
        
        If Denominator > 0 Then
            CurrentL_uH = (0.8 * (R_avg_in ^ 2) * (N ^ 2)) / Denominator
        End If
        
        If SafetyCounter > 10000 Then
            CoilCalc = Array(0, 0, 0, 0)
            Exit Function
        End If
    Loop
    
    Dim TotalLength_m As Double, Resistance As Double, Section_mm2 As Double
    TotalLength_m = N * 2 * Pi * (R_avg_mm / 1000)
    Section_mm2 = Pi * ((WireD_Cu / 2) ^ 2)
    
    If Section_mm2 > 0 Then Resistance = (TotalLength_m * 0.01724) / Section_mm2 Else Resistance = 0
    
    CoilCalc = Array(N, TotalLength_m, Resistance, Thickness_mm)
End Function



