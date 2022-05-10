' Ralf Mouthaan
' University of Cambridge
' June 2019
'
' Form class for Jasper control

Option Explicit On
Option Strict On

Public Class frmMain
    Private Sub frmMain_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

        Me.Enabled = False
        Application.DoEvents()

        MeasurementSetup = New clsMeasurementSetup
        MeasurementSetup.SLM.intScreenNo = 1 'Does not always match Windows screen number - might need to experiment

        'Set offset range from SLM resolution
        MeasurementSetup.SLM.StartUp()
        MeasurementSetup.SLM.ShutDown()
        nudTx1Offsetx.Maximum = MeasurementSetup.SLM.intResolutionx
        nudTx1Offsety.Maximum = MeasurementSetup.SLM.intResolutiony
        nudTx2Offsetx.Maximum = MeasurementSetup.SLM.intResolutionx
        nudTx2Offsety.Maximum = MeasurementSetup.SLM.intResolutiony

        'Guestimate offsets for holograms
        MeasurementSetup.Tx1Holo.intOffsetx = CInt(MeasurementSetup.SLM.intResolutionx * 1 / 4)
        MeasurementSetup.Tx1Holo.intOffsety = CInt(MeasurementSetup.SLM.intResolutiony / 2)
        MeasurementSetup.Tx2Holo.intOffsetx = CInt(MeasurementSetup.SLM.intResolutionx * 3 / 4)
        MeasurementSetup.Tx2Holo.intOffsety = CInt(MeasurementSetup.SLM.intResolutiony / 2)
        MeasurementSetup.Tx1Holo.dblScale = 1
        MeasurementSetup.Tx2Holo.dblScale = 1
        SetFormParametersFromMeasurementSetup()

        Me.Enabled = True
        Me.Activate()

    End Sub

    Private WriteOnly Property AllowUserInteraction As Boolean
        Set(value As Boolean)

            If value = False Then
                grpbxTx1.Enabled = False
                grpbxTx2.Enabled = False
                cmdUpdateSLM.Enabled = False
                cmdLoadZernikes.Enabled = False
                cmdSaveZernikes.Enabled = False
                cmdLoadHologramTx1.Enabled = False
                cmdLoadHologramTx2.Enabled = False
            ElseIf value = True Then
                grpbxTx1.Enabled = True
                grpbxTx2.Enabled = True
                cmdUpdateSLM.Enabled = True
                cmdLoadZernikes.Enabled = True
                cmdSaveZernikes.Enabled = True
                cmdLoadHologramTx1.Enabled = True
                cmdLoadHologramTx2.Enabled = True
            End If

            Application.DoEvents()

        End Set
    End Property


#Region "Data ferrying routines"
    Private Sub SetMeasurementSetupParametersFromForm()

        With MeasurementSetup

            'Set appropriate Tx1 hologram.
            'If loaded hologram is selected but nothing is loaded, switch back to blank holgoram
            If radTx1ShowLoaded.Checked = True Then
                If .Tx1LoadedHolo.bolLoaded = False Then
                    radTx1ShowBlank.Checked = True
                Else
                    .Tx1Holo = .Tx1LoadedHolo
                End If
            End If
            If radTx1ShowBlank.Checked = True Then
                .Tx1Holo = .Tx1BlankHolo
            End If

            'Set Tx1 hologram flags
            .Tx1Holo.bolVisible = chkTx1Visible.Checked
            .Tx1Holo.bolCircularAperture = chkTx1Circular.Checked
            .Tx1Holo.bolApplyZernikes = chkTx1Zernikes.Checked

            'Set Tx1 zernikes
            With .Tx1Holo
                .dblPiston = nudTx1Piston.Value
                .intOffsetx = CInt(nudTx1Offsetx.Value)
                .intOffsety = CInt(nudTx1Offsety.Value)
                .dblTiltx = nudTx1Tiltx.Value
                .dblTilty = nudTx1Tilty.Value
                .dblFocus = nudTx1Focus.Value
                .dblAstigmatismx = nudTx1Astigmatismx.Value
                .dblAstigmatismy = nudTx1Astigmatismy.Value
                .dblScale = nudTx1Scale.Value
                .dblRotation = nudTx1Rotation.Value
            End With

            'Set appropriate Tx2 hologram.
            'If loaded hologram is selected but nothing is loaded, switch back to blank holgoram
            If radTx2ShowLoaded.Checked = True Then
                If .Tx2LoadedHolo.bolLoaded = False Then
                    radTx2ShowBlank.Checked = True
                Else
                    .Tx2Holo = .Tx2LoadedHolo
                End If
            End If
            If radTx2ShowBlank.Checked = True Then
                .Tx2Holo = .Tx2BlankHolo
            End If

            'Set Tx1 hologram flags
            .Tx2Holo.bolVisible = chkTx2Visible.Checked
            .Tx2Holo.bolCircularAperture = chkTx2Circular.Checked
            .Tx2Holo.bolApplyZernikes = chkTx2Zernikes.Checked

            'Set Tx2 zernikes
            With .Tx2Holo
                .dblPiston = nudTx2Piston.Value
                .intOffsetx = CInt(nudTx2Offsetx.Value)
                .intOffsety = CInt(nudTx2Offsety.Value)
                .dblTiltx = nudTx2Tiltx.Value
                .dblTilty = nudTx2Tilty.Value
                .dblFocus = nudTx2Focus.Value
                .dblAstigmatismx = nudTx2Astigmatismx.Value
                .dblAstigmatismy = nudTx2Astigmatismy.Value
                .dblScale = nudTx2Scale.Value
                .dblRotation = nudTx2Rotation.Value
            End With

        End With

        MeasurementSetup.SLM.lstHolograms.Clear()
        MeasurementSetup.SLM.lstHolograms.Add(MeasurementSetup.Tx1Holo)
        MeasurementSetup.SLM.lstHolograms.Add(MeasurementSetup.Tx2Holo)

        'Note: Still need to refresh SLMs at this point!

    End Sub
    Private Sub SetFormParametersFromMeasurementSetup()

        With MeasurementSetup

            If TypeOf (.Tx1Holo) Is clsBlankHologram Then
                radTx1ShowBlank.Checked = True
            ElseIf TypeOf (.Tx1Holo) Is clsLoadedHologram Then
                radTx1ShowLoaded.Checked = True
            End If
            chkTx1Visible.Checked = .Tx1Holo.bolVisible
            chkTx1Circular.Checked = .Tx1Holo.bolCircularAperture
            chkTx1Zernikes.Checked = .Tx1Holo.bolApplyZernikes
            With .Tx1Holo
                nudTx1Offsetx.Value = .intOffsetx
                nudTx1Offsety.Value = .intOffsety
                nudTx1Tiltx.Value = CDec(.dblTiltx)
                nudTx1Tilty.Value = CDec(.dblTilty)
                nudTx1Focus.Value = CDec(.dblFocus)
                nudTx1Astigmatismx.Value = CDec(.dblAstigmatismx)
                nudTx1Astigmatismy.Value = CDec(.dblAstigmatismy)
                nudTx1Scale.Value = CDec(.dblScale)
                nudTx1Rotation.Value = CDec(.dblRotation)
            End With

            If TypeOf (.Tx2Holo) Is clsBlankHologram Then
                radTx2ShowBlank.Checked = True
            ElseIf TypeOf (.Tx2Holo) Is clsLoadedHologram Then
                radTx2ShowLoaded.Checked = True
            End If
            chkTx2Visible.Checked = .Tx2Holo.bolVisible
            chkTx2Circular.Checked = .Tx2Holo.bolCircularAperture
            chkTx2Zernikes.Checked = .Tx2Holo.bolApplyZernikes
            With .Tx2Holo
                nudTx2Offsetx.Value = .intOffsetx
                nudTx2Offsety.Value = .intOffsety
                nudTx2Tiltx.Value = CDec(.dblTiltx)
                nudTx2Tilty.Value = CDec(.dblTilty)
                nudTx2Focus.Value = CDec(.dblFocus)
                nudTx2Astigmatismx.Value = CDec(.dblAstigmatismx)
                nudTx2Astigmatismy.Value = CDec(.dblAstigmatismy)
                nudTx2Scale.Value = CDec(.dblScale)
                nudTx2Rotation.Value = CDec(.dblRotation)
            End With

        End With

    End Sub
#End Region

    Private Sub cmdUpdateSLM_Click(sender As Object, e As EventArgs) Handles cmdUpdateSLM.Click

        AllowUserInteraction = False
        MeasurementSetup.StartUp()
        SetMeasurementSetupParametersFromForm()
        MeasurementSetup.SLM.Refresh()
        AllowUserInteraction = True

    End Sub

    Public Sub SaveImgToFile(ByVal Img(,) As Double, ByVal Filename As String)

        Dim writer As New System.IO.StreamWriter(Filename)

        For i = 0 To UBound(Img, 1)
            For j = 0 To UBound(Img, 2)
                writer.Write(Img(i, j).ToString("F3"))
                If j < UBound(Img, 2) Then writer.Write(vbTab)
            Next
            If i < UBound(Img, 1) Then writer.WriteLine()
        Next

        writer.Close()

    End Sub

    Private Sub cmdSaveZernikes_Click(sender As Object, e As EventArgs) Handles cmdSaveZernikes.Click
        SetMeasurementSetupParametersFromForm()
        MeasurementSetup.Tx1Holo.SaveZernikes("D:\RPM Data Files\Tx1 Zernikes.txt")
        MeasurementSetup.Tx2Holo.SaveZernikes("D:\RPM Data Files\Tx2 Zernikes.txt")
    End Sub

    Private Sub cmdLoadZernikes_Click(sender As Object, e As EventArgs) Handles cmdLoadZernikes.Click
        MeasurementSetup.Tx1Holo.LoadZernikes("D:\RPM Data Files\Tx1 Zernikes.txt")
        MeasurementSetup.Tx2Holo.LoadZernikes("D:\RPM Data Files\Tx2 Zernikes.txt")
        SetFormParametersFromMeasurementSetup()
    End Sub

    Private Sub cmdLoadHologramTx1_Click(sender As Object, e As EventArgs) Handles cmdLoadHologramTx1.Click

        Dim dlg As New System.Windows.Forms.OpenFileDialog

        dlg.CheckFileExists = True
        dlg.InitialDirectory = "D:/"

        dlg.ShowDialog()

        If dlg.FileName = "" Then Exit Sub

        MeasurementSetup.Tx1LoadedHolo.LoadFile(dlg.FileName)

    End Sub

    Private Sub cmdLoadHologramTx2_Click(sender As Object, e As EventArgs) Handles cmdLoadHologramTx2.Click

        Dim dlg As New System.Windows.Forms.OpenFileDialog

        dlg.CheckFileExists = True
        dlg.InitialDirectory = "D:/"

        dlg.ShowDialog()

        If dlg.FileName = "" Then Exit Sub

        MeasurementSetup.Tx2LoadedHolo.LoadFile(dlg.FileName)

    End Sub
End Class
