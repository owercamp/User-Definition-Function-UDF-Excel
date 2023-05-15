VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Esta subrutina se llama al abrir el libro
Private Sub Workbook_Open()
    infoInterpretacion
    infoBuscarop
    DescribeFunctionContarDato
    infoIMedicalFacture
    DescribeFunctionFramingham
End Sub

'''
' Function Name:
' IMEDICALFACTURE

'  Description:
' Trae el precios de facturación en Avancys según corresponde número de identificación y código CUPS

'  Arguments:
' 1. Numero de identificación a buscar: Número de identificación a buscar.
' 2. Matriz en la cual se consulta el número de identificación.
' 3. Número de identificación CUPS a buscar: Número de identificación CUPS a buscar.
' 4. Matriz en la cual se consulta el número de identificación CUPS
'''
Private Sub infoIMedicalFacture()

    Dim NombreFunc As String
    Dim DescFunc As String
    Dim Categoria As String
    Dim DescArg(1 To 4) As String

    NombreFunc = "IMEDICALFACTURE"
    DescFunc = "Trae el precios de facturaci"&Chr(243)&"n en Avancys segun corresponde N"&Chr(186)&" de indetificaci"&Chr(243)&"n y codigo CUPS"
    Categoria = "SOA"
    DescArg(1) = "N"&Chr(250)&"mero de identificaci"&Chr(243)&"n a buscar"
    DescArg(2) = "Matriz en la cual se consulta el N"&Chr(250)&"mero de identificaci"&Chr(243)&"n"
    DescArg(3) = "N"&Chr(250)&"mero de identificaci"&Chr(243)&"n CUPS a buscar"
    DescArg(4) = "Matriz en la cual se consulta el N"&Chr(250)&"mero de identificaci"&Chr(243)&"n CUPS"

    On Error Resume Next
    Application.MacroOptions _
    Macro:=NombreFunc, _
    Description:=DescFunc, _
    Category:=Categoria, _
    ArgumentDescriptions:=DescArg

End Sub

'''
' Function Name:
' BUSCAROP

'  Description:
' Trae la información hacia la izquierda o derecha desde un punto de partida (punto de búsqueda)

'  Arguments:
' 1. Valor a buscar dentro de una matriz: Valor a buscar dentro de una matriz.
' 2. Matriz en la cual se consulta el valor_buscado: Matriz en la cual se consulta el valor_buscado.
' 3. Posición hacia la derecha o izquierda de la columna requerida según plano cartesiano número positivo o negativo: Posición hacia la derecha o izquierda de la columna requerida según plano cartesiano número positivo o negativo.
'''
Private Sub infoBuscarop()

    Dim NombreFunc As String
    Dim DescFunc As String
    Dim Categoria As String
    Dim DescArg(1 To 3) As String

    NombreFunc = "BUSCAROP"
    DescFunc = "Trae la informaci"&Chr(243)&"n hacia la izquierda o derecha desde un punto de partida" & vbNewLine & "(punto de busqueda)"
    Categoria = "SOA"
    DescArg(1) = "Valor a buscar dentro de una matriz"
    DescArg(2) = "Matriz en la cual se consulta el valor_buscado"
    DescArg(3) = "Posici"&Chr(243)&"n hacia la derecha o izquierda de la columna requerida seg"&Chr(250)&"n plano cartesiano n"&Chr(250)&"mero positivo o negativo"

    On Error Resume Next
    Application.MacroOptions _
    Macro:=NombreFunc, _
    Description:=DescFunc, _
    Category:=Categoria, _
    ArgumentDescriptions:=DescArg

End Sub

'''
' Function Name:
' INTERPRETACION

'  Description:
' Clasificación como NORMAL o ANORMAL según índice

'  Arguments:
' 1. Valor o celda a consultar: Valor o celda a consultar.
' 2. Valor de rango o celda con rangos para la validación ejm: 1000 a 5000.
' 3. Valor por el cual va ha hacer dividido valorRango ejm: 'a'.
'''
Private Sub infoInterpretacion()

    Dim NombreFunc As String
    Dim DescFunc As String
    Dim Categoria As String
    Dim DescArg(1 To 3) As String

    NombreFunc = "INTERPRETACION"
    DescFunc = "clasificaci"&Chr(243)&"n como NORMAL o ANORMAL segun indice"
    Categoria = "SOA"
    DescArg(1) = "valor o celda a consultar"
    DescArg(2) = "Valor de rango o celda con rangos para la validaci"&Chr(243)&"n ejm: 1000 a 5000"
    DescArg(3) = "Valor por el cual va ha hacer dividido valorRango ejm: 'a'"

    On Error Resume Next
    Application.MacroOptions _
    Macro:=NombreFunc, _
    Description:=DescFunc, _
    Category:=Categoria, _
    ArgumentDescriptions:=DescArg

End Sub

'''
' Function Name:
' CONTARDATO

'  Description:
' Cuenta el carácter enviado solo en las celdas visibles.

'  Arguments:
' 1. Es el rango de celdas para validar: Es el rango de celdas para validar.
' 2. Es el valor a evaluar: Es el valor a evaluar.
'''
Private Sub DescribeFunctionContarDato()
    Dim NombreFunc As String
    Dim DescFunc As String
    Dim Categoria As String
    Dim DescArg(1 To 2) As String
    '
    NombreFunc = "CONTARDATO"
    DescFunc = "Cuenta el caracter enviado solo en las celdas visibles."
    Categoria = "SOA"
    DescArg(1) = "Es el rango de celdas para validar"
    DescArg(2) = "Es el valor a evaluar"
    '
    On Error Resume Next
    Application.MacroOptions _
    Macro:=NombreFunc, _
    Description:=DescFunc, _
    Category:=Categoria, _
    ArgumentDescriptions:=DescArg

End Sub

'''
' Function Name:
' FRAMINGHAM

'  Description:
' Realiza la evaluación para determinar el riesgo CARDIOVASCULAR según los datos proporcionados.

'  Arguments:
' 1. Es la celda en donde se encuentra el valor de la EDAD, también se puede proporcionar directamente su valor.
' 2. Es la celda en donde se encuentra el valor de COLESTEROL TOTAL, también se puede proporcionar directamente su valor.
' 3. Es la celda en donde se encuentra el valor de COLESTEROL DE ALTA DENSIDAD [HDL] INMUNOLÓGICO DIRECTO, también se puede proporcionar directamente su valor.
' 4. Es la celda en donde se encuentra el valor de TENSIÓN ARTERIAL, también se puede proporcionar directamente su valor, en este caso solo debe ser puesta la TENSIÓN SISTÓLICA.
' 5. Es la celda en donde se encuentra el valor de FUMA, también se puede proporcionar directamente su valor, en este caso solo debe ser puesto SI o NO.
' 6. Valor que indica si sufre de DIABETES, los valores permitidos son SI o NO.
' 7. Es la celda en donde se encuentra el valor del SEXO, también se puede proporcionar directamente su valor, en este caso solo debe ser puesto MASCULINO o FEMENINO.
'''
Private Sub DescribeFunctionFramingham()
    Dim NombreFunc As String
    Dim DescFunc As String
    Dim Categoria As String
    Dim DescArg(1 To 7) As String
    '
    NombreFunc = "FRAMINGHAM"
    DescFunc = "Realiza la evaluaci"&Chr(243)&"n para determinar el riesgo CARDIOVASCULAR seg"&Chr(250)&"n los datos proporcionados"
    Categoria = "SOA"
    DescArg(1) = "Es la celda en donde se encuentra el valor de la EDAD, tambien se puede proporcionar directamente su valor"
    DescArg(2) = "Es la celda en donde se encuentra el valor de COLESTEROL TOTAL, tambien se puede proporcionar directamente su valor"
    DescArg(3) = "Es la celda en donde se encuentra el valor de COLESTEROL DE ALTA DENSIDAD [HDL] INMUNOLOGICO DIRECTO, tambien se puede proporcionar directamente su valor"
    DescArg(4) = "Es la celda en donde se encuentra el valor de TENSION ARTERIAL, tambien se puede proporcionar directamente su valor, en este caso solo debe ser puesta la TENSION SISTOLICA"
    DescArg(5) = "Es la celda en donde se encuentra el valor de FUMA, tambien se puede proporcionar directamente su valor, en este caso solo debe ser puesto SI o NO"
    DescArg(6) = "Valor que indica si sufre de DIABETES, los valores permitidos son SI o NO"
    DescArg(7) = "Es la celda en donde se encuentra el valor del SEXO, tambien se puede proporcionar directamente su valor, en este caso solo debe ser puesto MASCULINO o FEMENINO"
    '
    On Error Resume Next
    Application.MacroOptions _
    Macro:=NombreFunc, _
    Description:=DescFunc, _
    Category:=Categoria, _
    ArgumentDescriptions:=DescArg

End Sub
