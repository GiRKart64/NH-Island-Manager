Public Class frmProperties
    Dim sPreviewImage As String
    Dim sNewImage As String
    Dim ImageChanged As Boolean




    Private Sub frmProperties_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ImageChanged = False
        'frmMain.ToolStripStatusLabel2.Text
        TabPage1.Text = frmMain.IslandListView.FocusedItem.Text
        PictureBox1.ImageLocation = GetPreviewImage(frmMain.IslandListView.FocusedItem.Tag)
        If PictureBox1.ImageLocation = "" Then
            PictureBox1.Image = My.Resources.img256
        End If


        Dim i As Integer = 0
        Dim iRes As Integer


        'MsgBox(frmMain.IslandListView.FocusedItem.Tag & "\Villager" & 1)
        For i = 0 To 7
            If My.Computer.FileSystem.DirectoryExists(frmMain.IslandListView.FocusedItem.Tag & "\Villager" & i) = True Then
                iRes = i + 1
            End If
        Next i


        TextBox1.Text = iRes
        TextBox3.Text = Format(IO.File.GetLastWriteTime(frmMain.IslandListView.FocusedItem.Tag & "\main.dat"), "yyyy-MM-dd HH:mm:ss")

        If My.Computer.FileSystem.FileExists(frmMain.IslandListView.FocusedItem.Tag & "\description.txt") = True Then
            TextBox2.Text = System.IO.File.ReadAllText(frmMain.IslandListView.FocusedItem.Tag & "\description.txt")
        End If

    End Sub

 
  

    Private Sub PictureBox1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.DoubleClick
        Dim sType As String = "Images (*.png *.jpg *.bmp)|*.png;*.jpg;*.bmp|All files (*.*)|*.*"
        Dim od As New OpenFileDialog()

        With od
            .Filter = sType
        End With


        If od.ShowDialog = DialogResult.OK Then
            sNewImage = od.FileName.ToString
            sPreviewImage = frmMain.IslandListView.FocusedItem.Tag & "\preview." & sNewImage.Substring(Len(sNewImage.ToString) - 3, 3)
            PictureBox1.ImageLocation = sNewImage
        End If


        ImageChanged = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ImageChanged = True Then

            If My.Computer.FileSystem.FileExists(frmMain.IslandListView.FocusedItem.Tag & "\preview.jpg") = True Then
                My.Computer.FileSystem.DeleteFile(frmMain.IslandListView.FocusedItem.Tag & "\preview.jpg")
            End If

            If My.Computer.FileSystem.FileExists(frmMain.IslandListView.FocusedItem.Tag & "\preview.png") = True Then
                My.Computer.FileSystem.DeleteFile(frmMain.IslandListView.FocusedItem.Tag & "\preview.png")
            End If

            If My.Computer.FileSystem.FileExists(frmMain.IslandListView.FocusedItem.Tag & "\preview.bmp") = True Then
                My.Computer.FileSystem.DeleteFile(frmMain.IslandListView.FocusedItem.Tag & "\preview.bmp")
            End If

            My.Computer.FileSystem.CopyFile(sNewImage, sPreviewImage, True)

            If GetImage(frmMain.IslandListView.FocusedItem.Tag) = 1 Then
                frmMain.IslandListView.Items(frmMain.IslandListView.FocusedItem.Index).ImageIndex = frmMain.ImageList256.Images.Count - 1
            End If
        End If


        If TextBox2.Text = "" Then
        Else
            frmMain.IslandListView.Items(frmMain.IslandListView.FocusedItem.Index).SubItems.Item(2).Text = TextBox2.Text
            My.Computer.FileSystem.WriteAllText(frmMain.IslandListView.FocusedItem.Tag & "\description.txt", TextBox2.Text, False)
        End If


        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
        Me.Dispose()
    End Sub
End Class