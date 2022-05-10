' Ralf Mouthaan
' University of Cambridge
' June 2019
'
' Measurement setup class for Jasper control

Option Explicit On
Option Strict On

Public Module Globals
    Public MeasurementSetup As New clsMeasurementSetup
End Module

Public Class clsMeasurementSetup

    Public SLM As clsSLM
    Public Tx1Holo As clsHologram
    Public Tx2Holo As clsHologram
    Public Tx1BlankHolo As clsBlankHologram
    Public Tx2BlankHolo As clsBlankHologram
    Public Tx1LoadedHolo As clsLoadedHologram
    Public Tx2LoadedHolo As clsLoadedHologram

    Public Sub New()

        SLM = New clsSLM

        Tx1BlankHolo = New clsBlankHologram
        Tx2BlankHolo = New clsBlankHologram
        Tx1BlankHolo.Width = 1250
        Tx2BlankHolo.Width = 1250
        Tx1Holo = Tx1BlankHolo
        Tx2Holo = Tx2BlankHolo

        Tx1LoadedHolo = New clsLoadedHologram
        Tx2LoadedHolo = New clsLoadedHologram

        SLM.lstHolograms.Add(Tx1Holo)
        SLM.lstHolograms.Add(Tx2Holo)
        SLM.SetLinearCalibrationFactor(4.2 * 2 / Math.PI)

        'Additionally, elsewhere:
        '   * Holograms need to be loaded
        '   * Screen numbers for SLMs need to be defined

    End Sub

    Public Property arrZernikes(ByVal index As Integer) As Double
        Get
            If index = 0 Then Return Tx1Holo.intOffsetx
            If index = 1 Then Return Tx1Holo.intOffsety
            If index = 2 Then Return Tx1Holo.dblTiltx
            If index = 3 Then Return Tx1Holo.dblTilty
            If index = 4 Then Return Tx1Holo.dblFocus
            If index = 5 Then Return Tx1Holo.dblAstigmatismx
            If index = 6 Then Return Tx1Holo.dblAstigmatismy
            If index = 7 Then Return Tx2Holo.intOffsetx
            If index = 8 Then Return Tx2Holo.intOffsety
            If index = 9 Then Return Tx2Holo.dblTiltx
            If index = 10 Then Return Tx2Holo.dblTilty
            If index = 11 Then Return Tx2Holo.dblFocus
            If index = 12 Then Return Tx2Holo.dblAstigmatismx
            If index = 13 Then Return Tx2Holo.dblAstigmatismy
            Return Nothing
        End Get
        Set(value As Double)
            If index = 0 Then Tx1Holo.intOffsetx = CInt(value)
            If index = 1 Then Tx1Holo.intOffsety = CInt(value)
            If index = 2 Then Tx1Holo.dblTiltx = value
            If index = 3 Then Tx1Holo.dblTilty = value
            If index = 4 Then Tx1Holo.dblFocus = value
            If index = 5 Then Tx1Holo.dblAstigmatismx = value
            If index = 6 Then Tx1Holo.dblAstigmatismy = value
            If index = 7 Then Tx2Holo.intOffsetx = CInt(value)
            If index = 8 Then Tx2Holo.intOffsety = CInt(value)
            If index = 9 Then Tx2Holo.dblTiltx = value
            If index = 10 Then Tx2Holo.dblTilty = value
            If index = 11 Then Tx2Holo.dblFocus = value
            If index = 12 Then Tx2Holo.dblAstigmatismx = value
            If index = 13 Then Tx2Holo.dblAstigmatismy = value
        End Set
    End Property

    Public Sub StartUp()
        SLM.StartUp()
    End Sub
    Public Sub ShutDown()
        SLM.ShutDown()
    End Sub

End Class
