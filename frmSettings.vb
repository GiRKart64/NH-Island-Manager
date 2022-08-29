

Public Class frmSettings

    Public iChanged As Integer = iSettings(5)
    Private Sub frmSettings_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.ComboBox1.SelectedItem = Me.ComboBox1.Items().Item(iSettings(0))

        TextBox1.Text = sPaths(0)
        TextBox2.Text = sPaths(1)
        TextBox3.Text = sPaths(2)
        TextBox4.Text = sPaths(3)

        CheckBox1.CheckState() = iSettings(5)
        CheckBox2.CheckState() = iSettings(7)


        Select Case iSettings(6)
            Case 1
                RadioButton1.Checked = True
            Case 2
                RadioButton2.Checked = True
        End Select

    End Sub

 

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim FolderBrowserDialog1 As New FolderBrowserDialog
        With FolderBrowserDialog1
            .RootFolder = Environment.SpecialFolder.Desktop
            .SelectedPath = IIf(sPaths(0) = "", "c:\windows", IIf(My.Computer.FileSystem.DirectoryExists(sPaths(0)) = False, "c:\windows", sPaths(0)))
            .Description = "Select the source directory"
            If .ShowDialog = DialogResult.OK Then
                TextBox1.Text = FolderBrowserDialog1.SelectedPath.ToString
                sPaths(0) = FolderBrowserDialog1.SelectedPath.ToString
            End If
        End With
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Paths(TextBox2, "Switch Games (*.xci *.nsp)|*.xci;*.nsp|All files (*.*)|*.*")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Paths(TextBox3, "Executables (*.exe)|*.exe|All files (*.*)|*.*")
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Paths(TextBox4, "Executables (*.exe)|*.exe|All files (*.*)|*.*")
    End Sub

    Private Sub Paths(ByVal txtBox As TextBox, ByVal sType As String)

        Dim od As New OpenFileDialog()

        With od
            .Filter = sType
        End With


        If od.ShowDialog = DialogResult.OK Then

            txtBox.Text = od.FileName.ToString

        End If


    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        iSettings(6) = IIf(RadioButton1.Checked = True, 1, 2)
        frmMain.ToolStripButton5.ToolTipText = IIf(iSettings(6) = 1, "Open Yuzu Animal Crossing Save Folder", "Open Ryujinx Animal Crossing Save Folder")

        If CheckBox1.CheckState = iSettings(5) Then
            iSettings(5) = CheckBox1.CheckState
        Else
            iSettings(5) = CheckBox1.CheckState
            frmMain.ImageList032.Images.Clear()
            frmMain.ImageList064.Images.Clear()
            frmMain.ImageList128.Images.Clear()
            frmMain.ImageList256.Images.Clear()


            frmMain.ImageList032.Images.Add(BGImage(My.Resources.img032, 32))
            frmMain.ImageList064.Images.Add(BGImage(My.Resources.img064, 64))
            frmMain.ImageList128.Images.Add(BGImage(My.Resources.img128, 128))
            frmMain.ImageList256.Images.Add(BGImage(My.Resources.img256, 256))
     
        End If

        frmMain.IslandListView.Items.Clear()
        GetIslands()

        iSettings(0) = ComboBox1.SelectedIndex
        sPaths(0) = TextBox1.Text
        sPaths(1) = TextBox2.Text
        sPaths(2) = TextBox3.Text
        sPaths(3) = TextBox4.Text

        iSettings(7) = CheckBox2.CheckState


        Dim imgList As New System.Windows.Forms.ImageList
        Select Case ComboBox1.SelectedIndex
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
        Me.Close()
    End Sub





    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.Close()
    End Sub


End Class