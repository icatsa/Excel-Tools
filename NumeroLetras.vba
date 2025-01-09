Function NUMERO_LETRAS(Numero as Double) As String
Dim Letras As String
Dim Decimales As Double
Dim Numero As Double

Decimales = Numero - Int(Numero)
Numero = Int(Numero)

Dim Numeros(90) As String
Numeros(0) = "CERO"
Numeros(1) = "UNO"
Numeros(2) = "DOS"
Numeros(3) = "TRES"
Numeros(4) = "CUATRO"
Numeros(5) = "CINCO"
Numeros(6) = "SEIS"
Numeros(7) = "SIETE"
Numeros(8) = "OCHO"
Numeros(9) = "NUEVE"
Numeros(10) = "DIEZ"
Numeros(11) = "ONCE"
Numeros(12) = "DOCE"
Numeros(13) = "TRECE"
Numeros(14) = "CATORCE"
Numeros(15) = "QUINCE"
Numeros(16) = "DIECISEIS"
Numeros(17) = "DIECISIETE"
Numeros(18) = "DIECIOCHO"
Numeros(19) = "DIECINUEVE"
Numeros(20) = "VEINTE"
Numeros(21) = "VEINTIUNO"
Numeros(22) = "VEINTIDOS"
Numeros(23) = "VEINTITRES"
Numeros(24) = "VEINTICUATRO"
Numeros(25) = "VEINTICINCO"
Numeros(26) = "VEINTISEIS"
Numeros(27) = "VEINTISIETE"
Numeros(28) = "VEINTIOCHO"
Numeros(29) = "VEINTINUEVE"
Numeros(30) = "TREINTA"
Numeros(40) = "CUARENTA"
Numeros(50) = "CINCUENTA"
Numeros(60) = "SESENTA"
Numeros(70) = "SETENTA"
Numeros(80) = "OCHENTA"
Numeros(90) = "NOVENTA"
Do
    Select Case True
        Case Numero >= 100000000
                Letras = Letras & HandleHundreds(Int(Numero / 1000000), Numeros) & " MILLONES "
                Numero = Numero Mod 1000000
        Case Numero >= 1000000
                If Numero >= 2000000 Then
                    Letras = Letras & Numeros(Int(Numero / 1000000)) & " MILLONES "
                Else
                    Letras = Letras & "UN MILLÓN "
                End If
                Numero = Numero Mod 1000000

         Case Numero >= 100000
                Letras = Letras & HandleHundreds(Int(Numero / 1000), Numeros)
                Numero = Numero Mod 100000

        Case Numero >= 1000
                If Int(Numero / 1000) = 1 Then
                    Letras = Letras & "MIL "
                Else
                    Letras = Letras & Numeros(Int(Numero / 1000)) & " MIL "
                End If
                Numero = Numero Mod 1000
        
        Case Numero >= 100
                Letras = Letras & HandleHundreds(Numero, Numeros)
                Numero = Numero Mod 100
        
        Case Numero >= 10
                If Numero <= 30 Then
                    Letras = Letras & Numeros(Numero)
                    Numero = 0
                Else
                    Letras = Letras & Numeros(Int(Numero / 10) * 10)
                    Numero = Numero Mod 10
                    If Numero > 0 Then Letras = Letras & " Y "
                End If

        Case Numero > 0
                Letras = Letras & Numeros(Numero)
                Numero = 0
    End Select
Loop Until (Numero = 0)
Letras = Letras & " " & Format(Decimales * 100, "00") & "/100"
NUMERO_LETRAS = Letras
End Function

Function HandleHundreds(Numero As Double, Numeros() As String) As String
    Select Case Int(Numero / 100)
        Case 0: HandleHundreds = ""
        Case 1: HandleHundreds = IIf(Numero Mod 100 = 0, "CIEN ", "CIENTO ")
        Case 5: HandleHundreds = "QUINIENTOS "
        Case 7: HandleHundreds = "SETECIENTOS "
        Case 9: HandleHundreds = "NOVECIENTOS "
        Case Else: HandleHundreds = Numeros(Int(Numero / 100)) & "CIENTOS "
    End Select
End Function


Function DATE_STRING(theDate As Double)
DATE_STRING = day(theDate) & " DE " & UCase(Format(theDate, "mmmm")) & " DE " & year(theDate)
End Function
