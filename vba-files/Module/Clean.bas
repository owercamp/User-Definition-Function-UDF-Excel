Attribute VB_Name = "Clean"

Sub limpieza()

    Dim num, pos, contador, total As Integer
    Dim rn As Range
    Dim collect, Categories As Collection
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    porciento = 0
    total = Range("A2", Range("A2").End(xlDown)).count
    valorPorcentaje = 1 / total
    Application.StatusBar = "limpiando " & porciento & "%"
    
    '' UBICACIÖN INICIAL ''
    rngPpal = ActiveCell.Address
    contador = 1
    nums = InputBox("Ingrese el número de Columnas a verificar", "Verificación")
    pos = 0
    
    '' SE RECORREN TODOS LOS DATOS PARA REALIZAR LIMPIEZA DE DIAGNOSTICOS DUPLICADOS O QUE CORRESPONDAN A LA MISMA CATEGORIZACIÓN ''
    Do While Not IsEmpty(ActiveCell)
        rngPpal = ActiveCell.Address
        num = nums
        For Item = 1 To num
            contador = 1
            pos = 0
            ActiveCell = VBA.Replace(ActiveCell, ".", "")
            Rng = ActiveCell.Address
            Do While contador < CInt(num)
            
                '' SE CREA UNA NUEVA COLECCION PARA SER ALMACENADA CON LOS CODIGOS REFERENCES AL MISMO TIPO ''
                Set collect = New Collection
                Set Categories = New Collection
                
                '' SE RECORRE EL LISTA DE LOS CODIGOS PARA SACAR LA CATEGORIZACIÓN DEL DIAGNOSTICO ''
                Set rn = Worksheets("Hoja2").Range("B3", Worksheets("Hoja2").Range("B3").End(xlDown))
                For Each ItemRN In rn
                    If Trim(UCase(ItemRN)) = Trim(UCase(ActiveCell)) Then
                        Categories.Add ItemRN.Offset(0, 1)
                    End If
                    DoEvents
                Next ItemRN
    
                '' SE AGREGAN TODOS LOS DIAGNOSTICOS QUE CUENTEN CON LA MISMA CATEGORIZACION AL CODIGO ACTIVO EN LA COLECCION''
                For Each code In rn
                    For Each Category In Categories
                        If code.Offset(0, 1) = Category And code <> ActiveCell And ActiveCell <> 0 Then
                            collect.Add code
                        End If
                        DoEvents
                    Next Category
                    DoEvents
                Next code
            
            
                pos = pos + 2
                ActiveCell.Offset(0, pos) = VBA.Replace(ActiveCell.Offset(0, pos), ".", "")
                If Trim(UCase(ActiveCell)) = Trim(UCase(ActiveCell.Offset(0, pos))) Then
                    '' CODIGO ''
                    ActiveCell.Offset(0, pos) = "0"
                    '' DESCRIPCION ''
                    ActiveCell.Offset(0, (pos + 1)) = "0"
                ElseIf Trim(UCase(ActiveCell)) <> Trim(UCase(ActiveCell.Offset(0, pos))) Then
                    For Each itemCollect In collect
                        If Trim(UCase(itemCollect)) = Trim(UCase(ActiveCell.Offset(0, pos))) Then
                            '' CODIGO ''
                            ActiveCell.Offset(0, pos) = "0"
                            '' DESCRIPCION ''
                            ActiveCell.Offset(0, (pos + 1)) = "0"
                        End If
                        DoEvents
                    Next
                End If
                contador = contador + 1
                DoEvents
            Loop
            Range(Rng).Offset(0, 2).Select
            num = num - 1
            DoEvents
        Next
        Range(rngPpal).Select
        ActiveCell.Offset(1, 0).Select
        porciento = porciento + (VBA.Round(valorPorcentaje * 100, 2))
        If porciento < 100 Then
            Application.StatusBar = "Limpiando " & VBA.Round(porciento, 2) & "%"
        Else
            Application.StatusBar = "100% Limpiado"
        End If
        DoEvents
    Loop
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub
