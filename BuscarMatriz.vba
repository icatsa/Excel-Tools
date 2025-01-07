Option Explicit

' Formula para buscar datos en excel y devolver una matriz con todas las coincidencias.
' valor_buscado: valor que se busca en la matriz
' Matriz_tabla: rango de celdas donde se busca el valor
' Indicador_columnas: columna de la matriz donde se encuentra el valor que se quiere devolver
' fecha: valor de la fecha que se quiere buscar
' Columna_fecha: columna de la matriz donde se encuentra la fecha
' Opcion_busqueda: "v" para devolver la matriz en vertical, "h" para devolver la matriz en horizontal

Function BUSCAR_MATRIZ(valor_buscado As Range, Matriz_tabla As Range, Indicador_columnas As Integer, fecha as Range, Columna_fecha as Integer, Optional Opcion_busqueda As String = "v")
Dim r As Single, Lrow, Lcol As Single, temp() As Variant
Dim matchDate As Boolean


ReDim temp(0)

For r = 1 To Matriz_tabla.Rows.Count
    matchDate = False
        If Not IsEmpty(fecha) Then
            if(fecha = Matriz_tabla.Cells(r, Columna_fecha)) Then
                matchDate = True
            End If
        End If

    If Valor_buscado = Matriz_tabla.Cells(r, 1) and matchDate Then
        temp(UBound(temp)) = Matriz_tabla.Cells(r, Indicador_columnas)
        ReDim Preserve temp(UBound(temp) + 1)
    End If
Next r

If Opcion_busqueda = "h" Then
    Lcol = Range(Application.Caller.Address).Columns.Count
    For r = UBound(temp) To Lcol
        temp(UBound(temp)) = ""
        ReDim Preserve temp(UBound(temp) + 1)
    Next r
    ReDim Preserve temp(UBound(temp) - 1)
    BUSCAR_MATRIZ = temp
Else
    Lrow = Range(Application.Caller.Address).Rows.Count
    For r = UBound(temp) To Lrow
        temp(UBound(temp)) = ""
        ReDim Preserve temp(UBound(temp) + 1)
    Next r
    ReDim Preserve temp(UBound(temp) - 1)
    BUSCAR_MATRIZ = Application.Transpose(temp)
End If

End Function


