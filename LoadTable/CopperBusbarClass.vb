Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class CopperBusbarClass

    Public busbarValue As String
    Public busbarPos As Point3d
    Public totalBreaker As Double

    Public Sub New(totalBreakerAux As Double, busbarPosAux As Point3d)

        Integer.TryParse(totalBreakerAux, totalBreaker)
        busbarPos = busbarPosAux

        If totalBreaker <> Nothing Then

            Select Case (totalBreaker * 1.3)
                Case < 140
                    busbarValue = "15x2"
                    Exit Select
                Case < 170
                    busbarValue = "15x3"
                    Exit Select
                Case < 185
                    busbarValue = "20x2"
                    Exit Select
                Case < 220
                    busbarValue = "20x3"
                    Exit Select
                Case < 270
                    busbarValue = "25x3"
                    Exit Select
                Case < 295
                    busbarValue = "20x5"
                    Exit Select
                Case < 315
                    busbarValue = "30x3"
                    Exit Select
                Case < 350
                    busbarValue = "25x5"
                    Exit Select
                Case < 400
                    busbarValue = "30x5"
                    Exit Select
                Case < 420
                    busbarValue = "40x3"
                    Exit Select
                Case < 520
                    busbarValue = "40x5"
                    Exit Select
                Case < 630
                    busbarValue = "50x5"
                    Exit Select
                Case < 760
                    busbarValue = "40x10"
                    Exit Select
                Case < 820
                    busbarValue = "50x10"
                    Exit Select
                Case < 970
                    busbarValue = "80x5"
                    Exit Select
                Case < 1060
                    busbarValue = "60x10"
                    Exit Select
                Case < 1200
                    busbarValue = "100x5"
                    Exit Select
                Case < 1380
                    busbarValue = "80x10"
                    Exit Select
                Case < 1700
                    busbarValue = "100x10"
                    Exit Select
                Case < 2000
                    busbarValue = "120x10"
                    Exit Select
                Case < 2500
                    busbarValue = "160x10"
                    Exit Select
                Case < 3000
                    busbarValue = "200x10"
                    Exit Select

            End Select

        End If

    End Sub

    ' Draws the busbar (dimensions for the phase/neutral/ground bars)
    Public Sub DrawBusbar()



    End Sub

End Class
