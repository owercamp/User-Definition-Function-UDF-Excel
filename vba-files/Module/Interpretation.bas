Attribute VB_Name = "Interpretation"
Option Explicit

Function INTERPRETACION(ByVal valorBuscado As Variant, ByVal valorRango As String, ByVal separador As String) As String
Attribute INTERPRETACION.VB_Description = "clasificaci"&Chr(243)&"n como NORMAL o ANORMAL segun indice"
Attribute INTERPRETACION.VB_ProcData.VB_Invoke_Func = " \n19"

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
Attribute BUSCAROP.VB_Description = "Trae la informaci"&Chr(243)&"n hacia la izquierda o derecha desde un punto de partida\r\n(punto de busqueda)"
Attribute BUSCAROP.VB_ProcData.VB_Invoke_Func = " \n19"

    Dim Item As Variant
    
    For Each Item In rango_busqueda
        If VBA.Trim(Item) = VBA.Trim(valor_buscado) Then
            BUSCAROP = Item.Offset(0, posicion)
        End If
    Next Item
    
End Function

Function CONTARDATO(ByVal data As Object, ByVal text As String) As Integer
Attribute CONTARDATO.VB_Description = "Cuenta el caracter enviado solo en las celdas visibles."
Attribute CONTARDATO.VB_ProcData.VB_Invoke_Func = " \n19"
    
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


