Imports System
Imports System.IO

Public Class frmMain





    Private Sub frmMain_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Settings_Save()
    End Sub



    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        MenuStrip1.Visible = False
        sEmuNames(0) = "Yuzu"
        sEmuNames(1) = "Ryujinx"

        sErrorMsgs(0) = "Island folder not found. Please specify a valid directory in settings."
        sErrorMsgs(1) = "Animal Crossing game not found. Please specify a valid file in settings"
        sErrorMsgs(2) = "Yuzu not found. Please specify a valid path in settings."
        sErrorMsgs(3) = "Ryujinx not found. Please specify a valid path in settings."

        ImageList032.Images.Add(BGImage(My.Resources.img032, 32))
        ImageList064.Images.Add(BGImage(My.Resources.img064, 64))
        ImageList128.Images.Add(BGImage(My.Resources.img128, 128))
        ImageList256.Images.Add(BGImage(My.Resources.img256, 256))


        Dim str As String
        Dim strArr() As String
        Dim i As Integer

        sPaths(4) = Environment.GetEnvironmentVariable("UserProfile") & "\AppData\Roaming\yuzu\nand\user\save\0000000000000000\00000000000000000000000000000000\01006F8002326000" 'link "01006F8002326000" dir (More info needed?)
        sPaths(5) = Environment.GetEnvironmentVariable("UserProfile") & "\AppData\Roaming\Ryujinx\bis\user\save\0000000000000001\0" 'link "0" dir (More info needed?)



        If File.Exists(Application.StartupPath & "\settings.ini") = False Then


            iSettings(0) = 1
            iSettings(1) = 530
            iSettings(2) = 380
            iSettings(3) = 460
            iSettings(4) = 150
            iSettings(5) = 0
            iSettings(6) = 1
            iSettings(7) = 0

        Else

            Dim fStream As New System.IO.FileStream(Application.StartupPath & "\settings.ini", IO.FileMode.Open)
            Dim sReader As New System.IO.StreamReader(fStream)
            sReader.ReadLine()

         For i = 0 To 3
                str = sReader.ReadLine
                strArr = str.Split("=")
                sPaths(i) = strArr(1)
            Next i

            sReader.ReadLine()
            sReader.ReadLine()

            For i = 0 To 7
                str = sReader.ReadLine
                strArr = str.Split("=")
                iSettings(i) = strArr(1)
            Next i

            fStream.Close()
            sReader.Close()

        End If


        ColumnHeader1.Width = iSettings(1)
        ColumnHeader2.Width = iSettings(2)
        ColumnHeader3.Width = iSettings(3)
        ColumnHeader4.Width = iSettings(4)


        If sPaths(0) = "" Then
        Else
            GetIslands()
        End If


        'IslandListView.Sorting = SortOrder.Ascending (Not implemented)
        IslandListView.Sort()


        ToolStripStatusLabel2.Text = ""
        CType(FileToolStripMenuItem.DropDown, ToolStripDropDownMenu).ShowImageMargin = False
        ToolStripButton5.ToolTipText = "Open " & sEmuNames(iSettings(6) - 1) & " Animal Crossing Save Folder"


    End Sub





    Public Sub RunCMD(ByVal command As String, Optional ByVal ShowWindow As Boolean = False, Optional ByVal WaitForProcessComplete As Boolean = False, Optional ByVal permanent As Boolean = False)
        Dim p As Process = New Process()
        Dim pi As ProcessStartInfo = New ProcessStartInfo()
        pi.Arguments = " " + If(ShowWindow AndAlso permanent, "/K", "/C") + " " + command
        pi.FileName = "cmd.exe"
        pi.CreateNoWindow = Not ShowWindow
        If ShowWindow Then
            pi.WindowStyle = ProcessWindowStyle.Normal
        Else
            pi.WindowStyle = ProcessWindowStyle.Hidden
        End If
        p.StartInfo = pi
        p.Start()
        If WaitForProcessComplete Then Do Until p.HasExited : Loop

    End Sub




    Private Sub frmMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        IslandListView.Width = Me.Width - 5
        IslandListView.Height = Me.Height - 75
    End Sub
    Private Sub IslandListView_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles IslandListView.DoubleClick
        OpenIsland(iSettings(6))
    End Sub

    Private Sub IslandListView_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles IslandListView.MouseUp
        If e.Button = Windows.Forms.MouseButtons.Right Then

            If My.Computer.FileSystem.FileExists(sPaths(2)) Then
                YuzuToolStripMenuItem.Enabled = True
            Else
                YuzuToolStripMenuItem.Enabled = False
            End If

            If My.Computer.FileSystem.FileExists(sPaths(3)) Then
                RyujinxToolStripMenuItem.Enabled = True
            Else
                RyujinxToolStripMenuItem.Enabled = False
            End If

            'OpenFolderToolStripMenuItem.Text = "Open " & """" & IslandListView.FocusedItem.Text & """" & " Folder" ?

            ContextMenuStrip1.Show(IslandListView, New Point(e.X, e.Y))
        End If
    End Sub

    Private Sub IslandListView_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles IslandListView.SelectedIndexChanged
        ToolStripStatusLabel2.Text = IslandListView.FocusedItem.Tag
    End Sub



    Public Sub Settings_Save()
        Dim fileReader As System.IO.StreamWriter
        fileReader = System.IO.File.CreateText(Application.StartupPath & "\settings.ini")

        fileReader.WriteLine("[paths]")
        fileReader.WriteLine("islands=" & sPaths(0))
        fileReader.WriteLine("animalcrossing=" & sPaths(1))
        fileReader.WriteLine("yuzu=" & sPaths(2))
        fileReader.WriteLine("ryujinx=" & sPaths(3))
        fileReader.WriteLine("")
        fileReader.WriteLine("[settings]")
        fileReader.WriteLine("iconsize=" & iSettings(0))
        fileReader.WriteLine("columnwidth1=" & ColumnHeader1.Width)
        fileReader.WriteLine("columnwidth2=" & ColumnHeader2.Width)
        fileReader.WriteLine("columnwidth3=" & ColumnHeader3.Width)
        fileReader.WriteLine("columnwidth4=" & ColumnHeader4.Width)
        fileReader.WriteLine("previews=" & iSettings(5))
        fileReader.WriteLine("emulatordefault=" & iSettings(6))
        fileReader.WriteLine("fullscreen=" & iSettings(7))

        fileReader.Close()
    End Sub


    Private Sub SettingsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SettingsToolStripMenuItem.Click
        frmSettings.ShowDialog()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    Public Sub OpenIsland(ByVal iEmu As Integer)

        Dim sFull As String = ""
        Dim p As Process
        Dim acEmuGameActive As Byte = 0


        For Each p In Process.GetProcessesByName(sEmuNames(iEmu - 1))

            If p.MainWindowTitle.Substring(0, sEmuNames(iEmu - 1).Length).ToLower = sEmuNames(iEmu - 1).ToLower Then
                acEmuGameActive = 1
            End If


            If InStr(1, p.MainWindowTitle, "Animal Crossing: New Horizons") > 0 Then
                acEmuGameActive = 2
            End If

        Next


        If My.Computer.FileSystem.FileExists(sPaths(1)) = False Then
            MsgBox(sErrorMsgs(1), MsgBoxStyle.Information, Me.Text)
            Exit Sub
        End If


        If My.Computer.FileSystem.FileExists(sPaths(iEmu + 1)) = False Then
            MsgBox(sErrorMsgs(iEmu + 1), MsgBoxStyle.Information, Me.Text)
        Else

            Select Case acEmuGameActive
                Case 0
                    sFull = IIf(iEmu = 1, IIf(iSettings(7) = 1, " -f", ""), IIf(iSettings(7) = 1, " --fullscreen", ""))
                    RunCMD("rmdir " & """" & sPaths(iEmu + 3) & """")
                    RunCMD("mklink /J " & """" & sPaths(iEmu + 3) & """" & " " & """" & IslandListView.FocusedItem.Tag & """")
                    Shell("""" & sPaths(iEmu + 1) & """" & sFull & IIf(iEmu = 1, " -g ", " ") & """" & sPaths(1) & """", AppWinStyle.NormalFocus)
                Case 1
                    RunCMD("rmdir " & """" & sPaths(iEmu + 3) & """")
                    RunCMD("mklink /J " & """" & sPaths(iEmu + 3) & """" & " " & """" & IslandListView.FocusedItem.Tag & """")
                    MsgBox(sEmuNames(iEmu - 1) & " is already running." & " Island has been set to " & Me.IslandListView.FocusedItem.Text & ".", MsgBoxStyle.Information, Me.Text)
                Case 2
                    MsgBox("Animal Crossing: New Horzions is already running on " & sEmuNames(iEmu - 1) & ". Please close the game to continue.", MsgBoxStyle.Information, Me.Text)
            End Select

        End If
     

    End Sub
    Private Sub YuzuToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles YuzuToolStripMenuItem.Click
        OpenIsland(1)
    End Sub





    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        ImageList032.Images.Clear()
        ImageList064.Images.Clear()
        ImageList128.Images.Clear()
        ImageList256.Images.Clear()

        ImageList032.Images.Add(BGImage(My.Resources.img032, 32))
        ImageList064.Images.Add(BGImage(My.Resources.img064, 64))
        ImageList128.Images.Add(BGImage(My.Resources.img128, 128))
        ImageList256.Images.Add(BGImage(My.Resources.img256, 256))

        IslandListView.Items.Clear()
        GetIslands()

        ' MsgBox(ImageList256.Images.Count)
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        frmAboutBox.ShowDialog()

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        frmSettings.ShowDialog()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        ImageList032.Images.Clear()
        ImageList064.Images.Clear()
        ImageList128.Images.Clear()
        ImageList256.Images.Clear()

        ImageList032.Images.Add(BGImage(My.Resources.img032, 32))
        ImageList064.Images.Add(BGImage(My.Resources.img064, 64))
        ImageList128.Images.Add(BGImage(My.Resources.img128, 128))
        ImageList256.Images.Add(BGImage(My.Resources.img256, 256))

        IslandListView.Items.Clear()
        GetIslands()
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        frmAboutBox.ShowDialog()
    End Sub

    Private Sub RyujinxToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RyujinxToolStripMenuItem.Click
        OpenIsland(2)
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If IslandListView.Items.Count = 0 Then
            Exit Sub
        Else
            frmProperties.ShowDialog()
        End If
    End Sub


    Private Sub PropertiesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PropertiesToolStripMenuItem.Click
        frmProperties.ShowDialog()
    End Sub

    Private Sub OpenFolderToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenFolderToolStripMenuItem.Click
        Shell("explorer.exe " & IslandListView.FocusedItem.Tag, AppWinStyle.NormalFocus)
    End Sub



    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click

        If My.Computer.FileSystem.DirectoryExists(sPaths(3 + iSettings(6))) = True Then
            Shell("explorer.exe " & sPaths(3 + iSettings(6)), AppWinStyle.NormalFocus)
        Else
            MsgBox(sEmuNames(iSettings(6) - 1) & " save directory not found.", MsgBoxStyle.Information, Me.Text)
        End If

    End Sub
End Class
