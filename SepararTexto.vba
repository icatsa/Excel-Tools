Public Function SEPARARTEXTO(RangoEntrada As Range, posicion As Integer, separador As String)
  Dim vSeparar As Variant
  Application.Volatile
  vSeparar = Split(RangoEntrada.Value, separador)
  SEPARARTEXTO = Trim(vSeparar(posicion - 1))
End Function
