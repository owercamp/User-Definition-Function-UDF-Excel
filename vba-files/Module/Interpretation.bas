Attribute VB_Name = "Interpretation"
Option Explicit

Function INTERPRETACION(ByVal valorBuscado As Variant, ByVal valorRango As String, ByVal separador As String) As String

    Dim separateVal, Min, Max As Variant

    separateVal = VBA.Split(VBA.UCase(valorRango), VBA.UCase(separador))
    Min = Int(separateVal(0))
    Max = Int(separateVal(1))

    If valorBuscado >= Min And valorBuscado <= Max Then
        INTERPRETACION = "NORMAL"
    Else
        INTERPRETACION = "ANORMAL"
    End If

End Function

Function BUSCAROP(ByVal valor_buscado As Variant, ByRef rango_busqueda As Range, ByVal posicion As Variant) As Variant

    Dim Item As Variant

    For Each Item In rango_busqueda
        If VBA.Trim(Item) = VBA.Trim(valor_buscado) Then
            BUSCAROP = Item.Offset(0, posicion)
        End If
    Next Item

End Function

Function CONTARDATO(ByVal data As Object, ByVal text As String) As Integer

    Dim contador As Integer
    Dim List As Object
    Dim Item As Variant

    Set List = data
    contador = 0
    For Each Item In data
        If Item.Columns.Hidden = False Then
            If Trim(UCase(Item)) = Trim(UCase(text)) Then: contador = contador + 1
            End If
        Next Item

        CONTARDATO = contador

End Function

Function IMEDICALFACTURE(ByVal identity As Variant, ByRef rng_identity As Range, ByVal cups As Variant, ByRef rng_cups As Range) As LongLong

    Dim item As Variant
    Dim rowU, columnU As LongLong

    For Each item In rng_identity
        If Trim(item) = Trim(identity) Then: rowU = item.Row
        Next item

        For Each item In rng_cups
            If Trim(item) = Trim(cups) Then: columnU = item.Column
            Next item

            IMEDICALFACTURE = rng_identity.Parent.Cells(rowU, columnU)

End Function

Function FRAMINGHAM(ByVal Age As Integer, ByVal Cholesterol As Integer, ByVal Hdl As Integer, ByVal Ts_tbs As String, ByVal Smoking As String, ByVal Diabetes As String, ByVal Sex As String) As String

    Dim Ts_tb() As String
    Dim Ts As Integer
    Dim Logarithm, finalAge, finalCholesterol, finalHdl, finalTs, finalSmoking, finalDiabetes, summation, totalValue As Double
    Dim logOfAge, logOfCT, logOfHDL, logOfTS, logOfSmoke, logOfDiabetes, defaultValues, result, total As Variant

    '' los valores de la posicion 0 son para el Sex femenino y posicion 1 para el masculino ''
    logOfAge = Array(2.32888, 3.06117)
    logOfCT = Array(1.20904, 1.1237)
    logOfHDL = Array(-0.70833, -0.93263)
    logOfTS = Array(2.76157, 1.93303)
    logOfSmoke = Array(0.52873, 0.65451)
    logOfDiabetes = Array(0.69154, 0.57367)
    defaultValues = Array(26.1931, 23.9802)

    Ts_tb = VBA.Split(Ts_tbs, "/")
    Ts = CInt(Ts_tb(0))

    Select Case Trim(UCase(Sex))
     Case "FEMENINO"
        finalAge = WorksheetFunction.Ln(Age) * logOfAge(0)
        finalCholesterol = WorksheetFunction.Ln(Cholesterol) * logOfCT(0)
        finalHdl = WorksheetFunction.Ln(Hdl) * logOfHDL(0)
        finalTs = WorksheetFunction.Ln(Ts) * logOfTS(0)
        finalSmoking = 0
        finalDiabetes = 0
        If Trim(UCase(Smoking)) = "FUMA" Then: finalSmoking = logOfSmoke(0)
            If Trim(UCase(Diabetes)) = "SI" Then: finalDiabetes = logOfDiabetes(0)

                summation = finalAge + finalCholesterol + finalHdl + finalTs + finalSmoking + finalDiabetes
                totalValue = VBA.Exp(summation - defaultValues(0))
                result = 1 - (WorksheetFunction.Power(0.95012, totalValue))
                total = Round(result, 3) * 100
             Case "MASCULINO"
                finalAge = Round((WorksheetFunction.Ln(Age) * logOfAge(1)), 8)
                finalCholesterol = Round((WorksheetFunction.Ln(Cholesterol) * logOfCT(1)), 9)
                finalHdl = Round((WorksheetFunction.Ln(Hdl) * logOfHDL(1)), 9)
                finalTs = Round((WorksheetFunction.Ln(Ts) * logOfTS(1)), 9)
                finalSmoking = 0
                finalDiabetes = 0
                If Trim(UCase(Smoking)) = "FUMA" Or Trim(UCase(Smoking)) = "SI" Then: finalSmoking = logOfSmoke(1)
                    If Trim(UCase(Diabetes)) = "SI" Then: finalDiabetes = logOfDiabetes(1)

                        summation = Round((finalAge + finalCholesterol + finalHdl + finalTs + finalSmoking + finalDiabetes), 7)
                        totalValue = Round(Exp((summation - defaultValues(1))), 9)
                        result = 1 - (WorksheetFunction.Power(0.88936, totalValue))
                        total = Round(result, 3) * 100
                    End Select

                    If total < 10 Then
                        FRAMINGHAM = CStr(total) & "% - BAJO"
                    ElseIf total >= 10 And total <= 20 Then
                        FRAMINGHAM = CStr(total) & "% - MODERADO"
                    ElseIf total > 20 And total <= 30 Then
                        FRAMINGHAM = CStr(total) & "% - ALTO"
                    ElseIf total > 30 Then
                        FRAMINGHAM = CStr(total) & "% - MUY ALTO"
                    End If

End Function

