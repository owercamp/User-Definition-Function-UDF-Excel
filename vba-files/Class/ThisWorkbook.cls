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
