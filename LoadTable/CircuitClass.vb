Public Class CircuitClass

    Public circ As String       ' Circuit number
    Public desc As String       ' Description
    Public load As Double       ' Circuit load
    Public r As Double          ' Phase R
    Public s As Double          ' Phase S
    Public t As Double          ' Phase T
    Public phases As String     ' Phases used by the circuit (e.g R+S)
    Public wire As Double       ' Wire section
    Public breaker As String    ' Breaker value
    Public dr As String         ' DR value

    Public connection As String  ' Connection Type (Single/Two/Three-Phase)

    'Constructor for CSV file
    Public Sub New(circAux As String, descAux As String, loadAux As Integer, phasesAux As String, rAux As Integer, sAux As Integer, tAux As Integer, wireAux As Double, breakerAux As String, drAux As String)

        circ = circAux
        desc = descAux
        load = loadAux
        r = rAux
        s = sAux
        t = tAux
        phases = phasesAux
        wire = wireAux
        breaker = breakerAux
        dr = drAux

    End Sub

    'Constructor for when the attributes come from the project itself
    Public Sub New(circAux As String, loadAux As Integer, connAux As String, wireAux As Double)

        circ = circAux
        load = loadAux
        connection = connAux
        wire = wireAux

        desc = ""
        r = 0
        s = 0
        t = 0
        phases = "R"
        breaker = "-"
        breaker = GetBreaker()

    End Sub

    Public Sub New()

        circ = 0
        desc = ""
        load = 0
        r = 0
        s = 0
        t = 0
        phases = "R"
        wire = 0
        breaker = "-"
        breaker = GetBreaker()

        connection = ""

    End Sub

    Private Function GetBreaker()

        Select Case wire
            Case 2.5
                breaker = "16 A - C"
                Exit Select
            Case 4.0
                breaker = "20 A - C"
                Exit Select
            Case 6.0
                breaker = "32 A - C"
                Exit Select
            Case Else
                breaker = "16 A - C"
        End Select

        Return breaker

    End Function

    Private Function GetPhases()

        If r >= 0 Then

        End If

    End Function

End Class
