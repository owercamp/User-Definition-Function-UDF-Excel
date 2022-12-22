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

Private Sub Workbook_Open()
    Call infoInterpretacion
    Call infoBuscarop
    Call DescribeFunctionContarDato
    Call infoIMedicalFacture
    Call DescribeFunctionFramingham
End Sub

Private  Sub infoIMedicalFacture()

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
