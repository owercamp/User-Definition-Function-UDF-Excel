Attribute VB_Name = "Interpretation"
Option Explicit

Public Function INTERPRETACION(Byval valorBuscado As Variant, Byval valorRango As String, Byval separador As String) As String
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

Private Function TryParseInt(Byval value As String, Byref result As Long) As Boolean

  ''' <summary>
  ''' Toma un valor de cadena y trata de convertirlo en un numero entero largo. Devuelve verdadero si la conversion fue exitosa, falso de lo contrario.
  ''' </summary>
  ''' <param name="value">La cadena que se intentara convertir a un numero entero largo.</param>
  ''' <param name="result">El resultado de la conversion se almacenara en esta variable por referencia.</param>
  ''' <returns>Verdadero si la conversion fue exitosa, falso de lo contrario.</returns>

  On Error Resume Next
  result = CLng(value)
  TryParseInt = (Err.Number = 0)
  On Error Goto 0
End Function

Public Function BUSCAROP(Byval valor_buscado As Variant, Byref rango_busqueda As Range, Byval posicion As Variant) As Variant
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

Public Function CONTARDATO(Byval data As Object, Byval text As String) As Integer
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

Public Function IMEDICALFACTURE(Byval identity As Variant, Byref rng_identity As Range, Byval cups As Variant, Byref rng_cups As Range) As LongPtr
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

Public Function FRAMINGHAM(Byval Age As Integer, Byval Cholesterol As Integer, Byval Hdl As Integer, Byval Ts_tbs As String, Byval Smoking As String, Byval Diabetes As String, Byval Sex As String) As String

  Dim valueAge As Integer, valueDiabetes As Integer, valueSmoking As Integer, valueCholesterol As Integer
  Dim valueHdl As Integer, valueSystolic As Integer, valueDiastolic As Integer, valuebloodPressure As Integer

  ' systolic - diastolic blood pressure separation
  Dim tmpSystolic As Integer
  Dim tmpDiastolic As Integer
  valueSystolic = VBA.Split(Ts_tbs, "/")(0)
  valueDiastolic = VBA.Split(Ts_tbs, "/")(1)

  Select Case VBA.UCase(Sex)
   Case "MASCULINO"
    ' age validation by gender
    Select Case VBA.CInt(Age)
     Case 30 To 34
      valueAge = -1
     Case 35 To 39
      valueAge = 0
     Case 40 To 44
      valueAge = 1
     Case 45 To 49
      valueAge = 2
     Case 50 To 54
      valueAge = 3
     Case 55 To 59
      valueAge = 4
     Case 60 To 64
      valueAge = 5
     Case 65 To 69
      valueAge = 6
     Case 70 To 74
      valueAge = 7
     Case Is < 30
      valueAge = -1
     Case Else
      valueAge = "Valor no permitido"
    End Select
    ' validation of diabetes by gender
    Select Case VBA.UCase(Diabetes)
     Case "SI", 1
      valueDiabetes = 2
     Case Else
      valueDiabetes = 0
    End Select
    ' validation of smoking by gender
    Select Case VBA.UCase(Smoking)
     Case "SI", 1, "FUMA"
      valueSmoking = 2
     Case Else
      valueSmoking = 0
    End Select
    ' validation of total cholesterol by gender
    Select Case VBA.CInt(Cholesterol)
     Case Is < 160
      valueCholesterol = -3
     Case 160 To 199
      valueCholesterol = 0
     Case 200 To 239
      valueCholesterol = 1
     Case 240 To 279
      valueCholesterol = 2
     Case Is >= 280
      valueCholesterol = 3
    End Select
    ' validation of total cholesterol hdl by gender
    Select Case VBA.CInt(Hdl)
     Case Is < 35
      valueHdl = 2
     Case 35 To 44
      valueHdl = 1
     Case 45 To 59
      valueHdl = 0
     Case Is >= 60
      valueHdl = -2
    End Select
    ' blood pressure validation by gender
    Select Case VBA.CInt(valueSystolic)
     Case Is <= 129
      tmpSystolic = 0
     Case 130 To 139
      tmpSystolic = 1
     Case 140 To 159
      tmpSystolic = 2
     Case Is >= 160
      tmpSystolic = 3
    End Select
    Select Case VBA.CInt(valueDiastolic)
     Case Is <= 84
      tmpDiastolic = 0
     Case 85 To 89
      tmpDiastolic = 1
     Case 90 To 99
      tmpDiastolic = 2
     Case Is >= 100
      tmpDiastolic = 3
    End Select
    ' validation of blood pressure
    If (tmpSystolic >= tmpDiastolic) Then
      valuebloodPressure = tmpSystolic
    Else
      valuebloodPressure = tmpDiastolic
    End If
   Case "FEMENINO"
    ' age validation by gender
    Select Case VBA.CInt(Age)
     Case 30 To 34
      valueAge = -9
     Case 35 To 39
      valueAge = -4
     Case 40 To 44
      valueAge = 0
     Case 45 To 49
      valueAge = 3
     Case 50 To 54
      valueAge = 6
     Case 55 To 59
      valueAge = 7
     Case 60 To 74
      valueAge = 8
     Case Is < 30
      valueAge = -1
     Case Else
      valueAge = "Valor no permitido"
    End Select
    ' validation of diabetes by gender
    Select Case VBA.UCase(Diabetes)
     Case "SI", 1
      valueDiabetes = 4
     Case Else
      valueDiabetes = 0
    End Select
    ' validation of smoking by gender
    Select Case VBA.UCase(Smoking)
     Case "SI", 1, "FUMA"
      valueSmoking = 2
     Case Else
      valueSmoking = 0
    End Select
    ' validation of total cholesterol by gender
    Select Case VBA.CInt(Cholesterol)
     Case Is < 160
      valueCholesterol = -2
     Case 160 To 199
      valueCholesterol = 0
     Case 200 To 239
      valueCholesterol = 1
     Case 240 To 279
      valueCholesterol = 1
     Case Is >= 280
      valueCholesterol = 3
    End Select
    ' validation of total cholesterol hdl by gender
    Select Case VBA.CInt(Hdl)
     Case Is < 35
      valueHdl = 5
     Case 35 To 44
      valueHdl = 2
     Case 45 To 49
      valueHdl = 1
     Case 50 To 59
      valueHdl = 0
     Case Is >= 60
      valueHdl = -2
    End Select
    ' blood pressure validation by gender
    Select Case VBA.CInt(valueSystolic)
     Case Is < 120
      tmpSystolic = -3
     Case 120 To 139
      tmpSystolic = 0
     Case 140 To 159
      tmpSystolic = 2
     Case Is >= 160
      tmpSystolic = 3
    End Select
    Select Case VBA.CInt(valueDiastolic)
     Case Is < 80
      tmpDiastolic = -3
     Case 80 To 89
      tmpDiastolic = 0
     Case 90 To 99
      tmpDiastolic = 2
     Case Is >= 100
      tmpDiastolic = 3
    End Select
    ' validation of blood pressure
    If (tmpSystolic >= tmpDiastolic) Then
      valuebloodPressure = tmpSystolic
    Else
      valuebloodPressure = tmpDiastolic
    End If
  End Select

  Dim total As Integer
  total = valueAge + valueDiabetes + valueSmoking + valueCholesterol + valueHdl + valuebloodPressure

  Select Case VBA.Ucase(Sex)
    Case "MASCULINO"
      Select Case VBA.CInt(total)
        Case Is <= -1
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 2% Bajo"
        Case -1 To 1
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 3% Bajo"
        Case 2
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 4% Bajo"
        Case 3
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 5% Bajo"
        Case 4
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 7% Bajo"
        Case 5
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 8% Bajo"
        Case 6
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 10% Moderado"
        Case 7
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 13% Moderado"
        Case 8
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 16% Moderado"
        Case 9
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 20% Moderado"
        Case 10
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 25% Alto"
        Case 11
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 31% Alto"
        Case 12
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 37% Alto"
        Case 13
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 45% Alto"
        Case Is >= 14
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 53% Alto"
      End Select
    Case "FEMENINO"
      Select Case VBA.CInt(total)
        Case -2
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 1% Bajo"
        Case -1 To 1
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 2% Bajo"
        Case 2 To 3
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 3% Bajo"
        Case 4 To 5
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 4% Bajo"
        Case 6
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 5% Bajo"
        Case 7
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 6% Bajo"
        Case 8
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 7% Bajo"
        Case 9
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 8% Bajo"
        Case 10
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 10% Moderado"
        Case 11
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 11% Moderado"
        Case 12
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 13% Moderado"
        Case 13
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 15% Moderado"
        Case 14
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 18% Moderado"
        Case 15
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 20% Moderado"
        Case 16
          FRAMINGHAM = "Riesgo de EVC(10 Años) - 24% Alto"
        Case Is >= 17
          FRAMINGHAM = "Riesgo de EVC(10 Años) - >27% Alto"
      End Select
  End Select
End Function
