Imports System
Imports System.IO

Module Main
    Public sErrorMsgs(4) As String

    Public iSettings(8) As Integer

    Public sPaths(6) As String
    '0 = Islands (Saved)
    '1 = Game location (Saved)
    '2 = Yuzu (Saved)
    '3 = Ryujinx (Saved)
    '4 = Save Location Yuzu (Runtime)
    '5 = Save Location Ryujinx (Runtime)






    Public Sub GetIslands()
        If sPaths(0) = "" Then
            MsgBox("No directory specified for Animal Crossing islands", MsgBoxStyle.Information, frmMain.Text)
            Exit Sub
        End If

        Dim iList As Integer = 0
        Dim iDir As Integer = Len(sPaths(0)) + 1

        For Each Dir As String In Directory.GetDirectories(sPaths(0))
            If My.Computer.FileSystem.FileExists(Dir & "\main.dat") Then
                If My.Computer.FileSystem.FileExists(Dir & "\Villager0\profile.dat") Then
                    If My.Computer.FileSystem.FileExists(Dir & "\landname.dat") Then

                        frmMain.IslandListView.Items.Add(GetIslandName(Dir & "\landname.dat"), IIf(iSettings(5) = 1, IIf(GetImage(Dir) = 1, frmMain.ImageList256.Images.Count - 1, 0), 0))
                        frmMain.IslandListView.Items(iList).SubItems.Add(Dir.Substring(iDir, Len(Dir) - iDir))

                        If My.Computer.FileSystem.FileExists(Dir & "\description.txt") = True Then
                            frmMain.IslandListView.Items(iList).SubItems.Add(System.IO.File.ReadAllText(Dir & "\description.txt"))
                        Else
                            frmMain.IslandListView.Items(iList).SubItems.Add("")
                        End If

                        frmMain.IslandListView.Items(iList).SubItems.Add(Format(IO.File.GetLastWriteTime(Dir & "\main.dat"), "yyyy-MM-dd HH:mm:ss"))
                        frmMain.IslandListView.Items(iList).Tag = Dir
                        iList = iList + 1

                    End If
                End If
            End If
        Next

        Dim imgList As New System.Windows.Forms.ImageList
        Select Case iSettings(0)
            Case 0
                imgList = frmMain.ImageList032
            Case 1
                imgList = frmMain.ImageList064
            Case 2
                imgList = frmMain.ImageList128
            Case 3
                imgList = frmMain.ImageList256
        End Select


        frmMain.IslandListView.LargeImageList = imgList
        frmMain.IslandListView.SmallImageList = imgList


    End Sub


    Public Function GetIslandName(ByVal sFileName As String)
        Dim i As Integer
        Dim sIsland As String = ""
        Dim iByte As Byte

        Dim bStream As New System.IO.FileStream(sFileName, IO.FileMode.Open)
        Dim bReader As New System.IO.StreamReader(bStream)

        For i = 0 To 21
            bStream.Position = i
            iByte = bStream.ReadByte

            Select Case iByte
                Case 0
                Case 2
                Case Else
                    sIsland = sIsland & Chr(iByte)
            End Select

            'End If
        Next i


        bStream.Close()
        bReader.Close()

        Return sIsland
    End Function

    Public Function GetImage(ByVal sDirectory As String) As Integer
        If My.Computer.FileSystem.FileExists(sDirectory & "\preview.png") Then
            frmMain.ImageList256.Images.Add(ResizeImageSub(Bitmap.FromFile(sDirectory & "\preview.png"), 256, 256))
            frmMain.ImageList128.Images.Add(ResizeImageSub(Bitmap.FromFile(sDirectory & "\preview.png"), 128, 128))
            frmMain.ImageList064.Images.Add(ResizeImageSub(Bitmap.FromFile(sDirectory & "\preview.png"), 64, 64))
            frmMain.ImageList032.Images.Add(ResizeImageSub(Bitmap.FromFile(sDirectory & "\preview.png"), 32, 32))
            GetImage = 1
        ElseIf My.Computer.FileSystem.FileExists(sDirectory & "\preview.jpg") Then
            frmMain.ImageList256.Images.Add(ResizeImageSub(Bitmap.FromFile(sDirectory & "\preview.jpg"), 256, 256))
            frmMain.ImageList128.Images.Add(ResizeImageSub(Bitmap.FromFile(sDirectory & "\preview.jpg"), 128, 128))
            frmMain.ImageList064.Images.Add(ResizeImageSub(Bitmap.FromFile(sDirectory & "\preview.jpg"), 64, 64))
            frmMain.ImageList032.Images.Add(ResizeImageSub(Bitmap.FromFile(sDirectory & "\preview.jpg"), 32, 32))
            GetImage = 1
        ElseIf My.Computer.FileSystem.FileExists(sDirectory & "\preview.bmp") Then
            frmMain.ImageList256.Images.Add(ResizeImageSub(Bitmap.FromFile(sDirectory & "\preview.bmp"), 256, 256))
            frmMain.ImageList128.Images.Add(ResizeImageSub(Bitmap.FromFile(sDirectory & "\preview.bmp"), 128, 128))
            frmMain.ImageList064.Images.Add(ResizeImageSub(Bitmap.FromFile(sDirectory & "\preview.bmp"), 64, 64))
            frmMain.ImageList032.Images.Add(ResizeImageSub(Bitmap.FromFile(sDirectory & "\preview.bmp"), 32, 32))
            GetImage = 1
        Else
            GetImage = 0
        End If

    End Function
    Public Function GetPreviewImage(ByVal sDirectory As String) As String

        If My.Computer.FileSystem.FileExists(sDirectory & "\preview.png") Then
            GetPreviewImage = sDirectory & "\preview.png"
        ElseIf My.Computer.FileSystem.FileExists(sDirectory & "\preview.jpg") Then
            GetPreviewImage = sDirectory & "\preview.jpg"
        ElseIf My.Computer.FileSystem.FileExists(sDirectory & "\preview.bmp") Then
            GetPreviewImage = sDirectory & "\preview.bmp"
        Else
            GetPreviewImage = ""
        End If

    End Function
    Public Function BGImage(ByVal sBitmap As Bitmap, ByVal iSize As Integer) As Bitmap

        Dim thumb As New Bitmap(iSize, iSize)
        Dim g As Graphics = Graphics.FromImage(thumb)
        Dim cPen As New SolidBrush(frmMain.IslandListView.BackColor)
        g.FillRectangle(cPen, 0, 0, iSize, iSize)
        g.DrawImage(sBitmap, New Rectangle(0, 0, iSize, iSize))
        g.Dispose()

        BGImage = New Bitmap(thumb)


    End Function

    Public Function ResizeImageSub(ByVal ssBitmap As Bitmap, ByVal iNewWidth As Integer, ByVal iNewHeight As Integer) As Bitmap
        Dim iThumbWidth As Integer
        Dim iThumbHeight As Integer
        Dim iTemp As Integer
        Dim iPosW As Integer
        Dim iPosH As Integer
        Dim iCalc As Integer

        If ssBitmap.Width > ssBitmap.Height Then
            iThumbWidth = iNewWidth
            iTemp = ssBitmap.Width / iNewWidth
            iThumbHeight = Int(ssBitmap.Height / iTemp)
            iCalc = iNewHeight - iThumbHeight
            iPosW = 0
            iPosH = Int(iCalc / 2)
        ElseIf ssBitmap.Height > ssBitmap.Width Then
            iThumbHeight = iNewHeight
            iTemp = ssBitmap.Height / iNewHeight
            iThumbWidth = Int(ssBitmap.Width / iTemp)
            iCalc = iNewWidth - iThumbWidth
            iPosW = Int(iCalc / 2)
            iPosH = 0
        Else
            iThumbWidth = iNewWidth
            iThumbHeight = iNewHeight
            iPosW = 0
            iPosH = 0
        End If


        Dim thumb As New Bitmap(iNewWidth, iNewHeight)
        Dim g As Graphics = Graphics.FromImage(thumb)
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.DrawImage(ssBitmap, New Rectangle(iPosW, iPosH, iThumbWidth, iThumbHeight))
        g.Dispose()
        ssBitmap.Dispose()
        ResizeImageSub = thumb

    End Function

 

End Module
