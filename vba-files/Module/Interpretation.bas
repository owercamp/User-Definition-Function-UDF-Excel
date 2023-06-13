Attribute VB_Name = "Interpretation"
Option Explicit

Public Function INTERPRETACION(ByVal valorBuscado As Variant, ByVal valorRango As String, ByVal separador As String) As String
  'TODO: Funcion que devuelve una interpretacion (NORMAL o ANORMAL) de acuerdo a si el valor buscado está dentro del rango definido por valorRango y separador
  '
  '? Parametros:
  '? @param valorBuscado: El valor que se desea buscar dentro del rango
  '? @param valorRango: El rango de valores separados por el separador definido. Ejemplo: "5@param10"
  '? @param separador: El separador usado para separar los valores del rango. Ejemplo: "-"
  '
  '? Devuelve:
  '? @return "NORMAL" si el valor buscado esta dentro del rango
  '? @return "ANORMAL" si el valor buscado esta fuera del rango
  '? @return "ERROR: El valor minimo no es un numero valido" si el valor minimo del rango no es un numero valido
  '? @return "ERROR: El valor maximo no es un numero valido" si el valor maximo del rango no es un numero valido

  '* Se divide el rango en dos valores, minimo y maximo
  Dim separateVal As Variant
  separateVal = VBA.Split(VBA.UCase(valorRango), VBA.UCase(separador))

  '* Se intenta parsear cada valor del rango como un numero entero
  Dim Min As Long
  Dim Max As Long
  If Not TryParseInt(separateVal(0), Min) Then
    INTERPRETACION = "ERROR: El valor m"&Chr(237)&"nimo no es un n"&Chr(250)&"mero v"&Chr(225)&"lido"
    Exit Function
  End If
  If Not TryParseInt(separateVal(1), Max) Then
    INTERPRETACION = "ERROR: El valor m"&Chr(225)&"ximo no es un n"&Chr(250)&"mero v"&Chr(225)&"lido"
    Exit Function
  End If

  '* Se compara el valor buscado con el rango
  If valorBuscado >= Min And valorBuscado <= Max Then
    INTERPRETACION = "NORMAL"
  Else
    INTERPRETACION = "ANORMAL"
  End If
End Function

Private Function TryParseInt(ByVal value As String, ByRef result As Long) As Boolean

  ''' <summary>
  ''' Toma un valor de cadena y trata de convertirlo en un numero entero largo. Devuelve verdadero si la conversion fue exitosa, falso de lo contrario.
  ''' </summary>
  ''' <param name="value">La cadena que se intentara convertir a un numero entero largo.</param>
  ''' <param name="result">El resultado de la conversion se almacenara en esta variable por referencia.</param>
  ''' <returns>Verdadero si la conversion fue exitosa, falso de lo contrario.</returns>

  On Error Resume Next
  result = CLng(value)
  TryParseInt = (Err.Number = 0)
  On Error GoTo 0
End Function

Public Function BUSCAROP(ByVal valor_buscado As Variant, ByRef rango_busqueda As Range, ByVal posicion As Variant) As Variant
  'TODO: Busca un valor en un rango y devuelve el valor de la celda correspondiente en una columna determinada.
  '
  '? Argumentos:
  '?  @param valor_buscado: El valor que se esta buscando en el rango.
  '?  @param rango_busqueda: El rango de celdas donde buscar.
  '?  @param posicion: El desplazamiento de la columna desde la celda encontrada hasta la celda que se debe devolver.
  '
  '? Retorna:
  '? @return El valor de la celda correspondiente en la columna especificada por posicion, si el valor buscado se encuentra en el rango. Si no se encuentra ninguna coincidencia en el rango de búsqueda, devuelve un error #N/A.

  Dim Item As Variant

  For Each Item In rango_busqueda
    If VBA.Trim(Item) = VBA.Trim(valor_buscado) Then
      BUSCAROP = Item.Offset(0, posicion)
      Exit Function
    End If
  Next Item

  BUSCAROP = CVErr(2042)
End Function

Public Function CONTARDATO(ByVal data As Object, ByVal text As String) As Integer
  'TODO: Esta funcion cuenta el numero de veces que aparece una cadena en una coleccion.
  '? Parámetros:
  '?  @param data: un objeto de coleccion
  '?  @param text: una cadena para buscar en la coleccion
  '? Devoluciones:
  '?  @return Un entero que representa el numero de veces que aparece la cadena en la coleccion.

  Dim contador As Integer
  Dim List As Object
  Dim Item As Variant

  Set List = data
  contador = 0
  For Each Item In List
    If Not Item.Columns.Hidden And Trim(UCase(Item)) = Trim(UCase(text)) Then
      contador = contador + 1
    End If
  Next Item

  CONTARDATO = contador
End Function

Public Function IMEDICALFACTURE(ByVal identity As Variant, ByRef rng_identity As Range, ByVal cups As Variant, ByRef rng_cups As Range) As LongPtr
  'TODO: Devuelve el valor ubicado en la intersección de los índices de fila y columna correspondientes a los valores de identity y cups dentro de los rangos especificados.
  '
  '? Argumentos:
  '?  @param identity: El valor de identity que se desea buscar dentro de los rangos.
  '?  @param rng_identity: El rango de celdas donde buscar.
  '?  @param cups: El valor de cups que se desea buscar dentro de los rangos.
  '?  @param rng_cups: El rango de celdas donde buscar.
  '
  '? Devoluciones:
  '? @return El valor ubicado en la intersección de los índices de fila y columna correspondientes a los valores de identity y cups dentro de los rangos especificados.

  Dim item As Variant
  Dim rowU, columnU As LongPtr

  '* Buscar el índice de fila correspondiente al valor de identity dado
  For Each item In rng_identity
    If Trim(item) = Trim(identity) Then
      rowU = item.Row
      Exit For
    End If
  Next item

  '* Buscar el índice de columna correspondiente al valor de cups dado
  For Each item In rng_cups
    If Trim(item) = Trim(cups) Then
      columnU = item.Column
      Exit For
    End If
  Next item

  '* Devolver el valor ubicado en la intersección de los índices de fila y columna
  IMEDICALFACTURE = rng_identity.Parent.Cells(rowU, columnU)
End Function

Public Function FRAMINGHAM(ByVal Age As Integer, ByVal Cholesterol As Integer, ByVal Hdl As Integer, ByVal Ts_tbs As String, ByVal Smoking As String, ByVal Diabetes As String, ByVal Sex As String) As String
  'TODO: Esta función utiliza el modelo de Framingham para estimar el riesgo de enfermedad cardiovascular en función de varios factores de riesgo.
  '
  '? Args:
  '? @param Age (int): edad de la persona (en años).
  '? @param Cholesterol (int): colesterol total de la persona (en mg/dL).
  '? @param Hdl (int): lipoproteína de alta densidad (HDL) de la persona (en mg/dL).
  '? @param Ts_tbs (str): relación entre el colesterol total y el HDL (en formato "X/Y").
  '? @param Smoking (str): indica si la persona fuma ("Fuma" si es fumador, de lo contrario "").
  '? @param Diabetes (str): indica si la persona tiene diabetes ("Si" si tiene diabetes, de lo contrario "").
  '? @param Sex (str): género de la persona ("Femenino" o "Masculino").
  '
  '? Returns:
  '? @return str: cadena que indica el nivel de riesgo cardiovascular, expresado como un porcentaje y una categoría del nivel de riesgo ("BAJO", "MODERADO", "ALTO" o "MUY ALTO").

  Dim Ts_tb() As String
  Dim Ts As Integer
  Dim Logarithm As Double, finalAge As Double, finalCholesterol As Double, finalHdl As Double, finalTs As Double, finalSmoking As Double, finalDiabetes As Double, summation As Double, totalValue As Double
  Dim logOfAge As Variant, logOfCT As Variant, logOfHDL As Variant, logOfTS As Variant, logOfSmoke As Variant, logOfDiabetes As Variant, defaultValues As Variant, result As Variant, total As Variant

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
    If Trim(UCase(Smoking)) = "FUMA" Then
      finalSmoking = logOfSmoke(0)
    End If
    If Trim(UCase(Diabetes)) = "SI" Then
      finalDiabetes = logOfDiabetes(0)
    End If

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
    If Trim(UCase(Smoking)) = "FUMA" Or Trim(UCase(Smoking)) = "SI" Then
      finalSmoking = logOfSmoke(1)
    End If
    If Trim(UCase(Diabetes)) = "SI" Then
      finalDiabetes = logOfDiabetes(1)
    End If

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
