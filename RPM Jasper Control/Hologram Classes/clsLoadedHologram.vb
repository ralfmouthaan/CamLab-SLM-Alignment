' Ralf Mouthaan
' University of Cambridge
' June 2019
'
' Class to display loaded hologram
' Currently assumes a square hologram
' Hologram is loaded from a text file, where it is saved as a tab-separated file of doubles corresponding to phase delays of the individual pixels

Option Explicit On
Option Strict On

Public Class clsLoadedHologram
    Inherits clsHologram

    Private _arrHologram(,) As Double
    Public bolLoaded As Boolean

    Public Sub New()
        Call MyBase.New
        bolLoaded = False
    End Sub

    Public Overrides Property arrRawHologram As Double(,)
        Get
            Return _arrHologram
        End Get
        Set(value As Double(,))
            Call Err.Raise(-1, "Loaded Hologram Class", "Cannot edit loaded hologram")
        End Set
    End Property
    Public Overrides Property arrRawHologram(i As Integer, j As Integer) As Double
        Get

            If i > _arrHologram.GetLength(0) - 1 Then Call Err.Raise(-1, "Loaded Hologram", "index beyond array dimensions")
            If j > _arrHologram.GetLength(1) - 1 Then Call Err.Raise(-1, "Loaded Hologram", "index beyond array dimensions")

            Return _arrHologram(i, j)

        End Get
        Set(value As Double)
            Call Err.Raise(-1, "Loaded Hologram Class", "Cannot set pixel values")
        End Set
    End Property

    Public Overrides Property rawWidth As Integer
        Get
            Return _arrHologram.GetLength(0)
        End Get
        Set(value As Integer)
            Call Err.Raise(-1, "Loaded Hologram Class", "Cannot set hologram width")
        End Set
    End Property

    Public Overrides Property rawHeight As Integer
        Get
            Return _arrHologram.GetLength(1)
        End Get
        Set(value As Integer)
            Call Err.Raise(-1, "Loaded Hologram Class", "Cannot set hologram height")
        End Set
    End Property

    Public Sub LoadFile(ByVal Filename As String)

        If IO.File.Exists(Filename) = False Then Call Err.Raise(-1, "Loaded Hologram Class", "File not found")

        If Filename.ToLower.EndsWith(".txt") Then
            Call LoadTXT(Filename)
        ElseIf Filename.EndsWith(".csv") Then
            Call LoadBinCSV(Filename)
        ElseIf Filename.ToLower.EndsWith(".bmp") Then
            Call LoadBmp(Filename)
        End If

    End Sub

    Private Sub LoadTXT(ByVal Filename As String)

        If Filename.EndsWith(".txt") = False Then Call Err.Raise(-1, "Loaded Hologram Class", "File must be a text file")

        Dim reader As IO.StreamReader
        Dim splitstr() As String

        reader = New IO.StreamReader(Filename)

        'First line
        splitstr = Split(reader.ReadLine, vbTab)
        ReDim _arrHologram(splitstr.Length - 1, splitstr.Length - 1)
        For i = 0 To splitstr.Length - 1
            _arrHologram(i, 0) = CDbl(splitstr(i))
        Next


        For j = 1 To _arrHologram.GetLength(1) - 1
            If reader.EndOfStream = True Then Call Err.Raise(-1, "Loaded Hologram Class", "Hologram does not seem to be square")
            splitstr = Split(reader.ReadLine, vbTab)
            For i = 0 To _arrHologram.GetLength(0) - 1
                _arrHologram(i, j) = CDbl(splitstr(i))
            Next
        Next

        reader.Close()

        Call RedefineLUTs()

        bolLoaded = True

    End Sub
    Private Sub LoadBinCSV(ByVal Filename As String)

        'Assumes hologram is binary with data stored as +1s and -1s

        If Filename.EndsWith(".csv") = False Then Call Err.Raise(-1, "Loaded Hologram Class", "File must be a csv file")

        Dim reader As IO.StreamReader
        Dim splitstr() As String

        reader = New IO.StreamReader(Filename)

        'First line
        splitstr = Split(reader.ReadLine, ",")
        ReDim _arrHologram(splitstr.Length - 1, splitstr.Length - 1)
        For i = 0 To splitstr.Length - 1
            _arrHologram(i, 0) = CDbl(splitstr(i)) / 2 * Math.PI
        Next


        For j = 1 To _arrHologram.GetLength(1) - 1
            If reader.EndOfStream = True Then Call Err.Raise(-1, "Loaded Hologram Class", "Hologram does not seem to be square")
            splitstr = Split(reader.ReadLine, ",")
            For i = 0 To _arrHologram.GetLength(0) - 1
                _arrHologram(i, j) = CDbl(splitstr(i)) / 2 * Math.PI
            Next
        Next

        reader.Close()

        Call RedefineLUTs()

        bolLoaded = True

    End Sub
    Private Sub LoadPhaseCSV(ByVal Filename As String)

        ' Assumes hologram is multiphase with data stored as an angle in radians

        If Filename.EndsWith(".csv") = False Then Call Err.Raise(-1, "Loaded Hologram Class", "File must be a csv file")

        Dim reader As IO.StreamReader
        Dim splitstr() As String

        reader = New IO.StreamReader(Filename)

        'First line
        splitstr = Split(reader.ReadLine, ",")
        ReDim _arrHologram(splitstr.Length - 1, splitstr.Length - 1)
        For i = 0 To splitstr.Length - 1
            _arrHologram(i, 0) = CDbl(splitstr(i))
        Next


        For j = 1 To _arrHologram.GetLength(1) - 1
            If reader.EndOfStream = True Then Call Err.Raise(-1, "Loaded Hologram Class", "Hologram does not seem to be square")
            splitstr = Split(reader.ReadLine, ",")
            For i = 0 To _arrHologram.GetLength(0) - 1
                _arrHologram(i, j) = CDbl(splitstr(i))
            Next
        Next

        reader.Close()

        Call RedefineLUTs()

        bolLoaded = True

    End Sub
    Private Sub LoadBMP(ByVal Filename As String)

        Dim bm As New Bitmap(Filename)
        ReDim _arrHologram(bm.Width - 1, bm.Height - 1)

        For i = 0 To bm.Height - 1
            For j = 0 To bm.Width - 1
                _arrHologram(i, j) = bm.GetPixel(i, j).R / 255 * 2 * Math.PI
            Next
        Next

        bolLoaded = True

        Call RedefineLUTs()

        'Dim bData As Imaging.BitmapData = bm.LockBits(New Rectangle(0, 0, bm.Width, bm.Height), Imaging.ImageLockMode.ReadOnly, bm.PixelFormat)

    End Sub

End Class
