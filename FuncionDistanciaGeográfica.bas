Attribute VB_Name = "Módulo1"
Function DistanciaGeografica(LatitudTienda, LongitudTienda, LatitudDestino, LongitudDestino, RadioTierra As Double) As Double
Dim distancia As Double
Dim DiferenciaLatitud As Double
Dim DiferenciaLongitud As Double
Dim a As Double
Dim b As Double


constante = WorksheetFunction.Pi / 180
DiferenciaLatitud = LatitudTienda - LatitudDestino
DiferenciaLongitud = LongitudTienda - LongitudDestino

a = Sin(DiferenciaLatitud * constante / 2) ^ 2 + Cos(LatitudTienda * constante) * Cos(LatitudDestino * constante) * Sin(DiferenciaLongitud * constante / 2) ^ 2

a = 2 * WorksheetFunction.Asin(Sqr(a))

DistanciaGeografica = a * RadioTierra * 1000
'Hecho por Jonathan Macalupu

End Function
