Attribute VB_Name = "Módulo11"
'Funcion que te permite calcular la distancia entre dos coordenadas geográficas utilizando la fórmula de Haversine
Function DistanciaGeografica(LatitudInicio, LongitudInicio, LatitudFin, LongitudFin, RadioTierra As Double) As Double
Dim distancia As Double
Dim DiferenciaLatitud As Double
Dim DiferenciaLongitud As Double
Dim a As Double
Dim b As Double


constante = WorksheetFunction.Pi / 180
DiferenciaLatitud = LatitudInicio - LatitudFin
DiferenciaLongitud = LongitudInicio - LongitudFin

a = Sin(DiferenciaLatitud * constante / 2) ^ 2 + Cos(LatitudInicio * constante) * Cos(LatitudFin * constante) * Sin(DiferenciaLongitud * constante / 2) ^ 2

a = 2 * WorksheetFunction.Asin(Sqr(a))

DistanciaGeografica = a * RadioTierra * 1000
'Hecho por Jonathan Macalupu

End Function
