Imports System.IO
Imports System.Net.Http
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.Win32
Imports Newtonsoft.Json

Public Class Form1
    Private Pages_current_page As Integer
    Private Pages_last_page As Integer
    Private Pages_per_page As Integer
    Private Pages_total As Integer
    Private Descending_sorting_order As Boolean = True
    Private Exactly_Resolution As Boolean = True
    Private SettingsLoadLock As Boolean = True
    ' Declare the SystemParametersInfo function from user32.dll
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SystemParametersInfo(ByVal uAction As UInteger, ByVal uParam As UInteger, ByVal lpvParam As String, ByVal fuWinIni As UInteger) As Boolean
    End Function

    ' Constants for setting the desktop background
    Private Const SPI_SETDESKWALLPAPER As UInteger = 20
    Private Const SPIF_UPDATEINIFILE As UInteger = &H1
    Private Const SPIF_SENDCHANGE As UInteger = &H2

    ' Form1 - Load
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ContextMenuStrip_Colors.Renderer = New ToolStripProfessionalRenderer(New ColorTable())
        ContextMenuStrip_ODD.Renderer = New ToolStripProfessionalRenderer(New ColorTable())
        ContextMenuStrip_WallpaperStyles.Renderer = New ToolStripProfessionalRenderer(New ColorTable())

        If Directory.Exists(My.Settings.DownloadDir) Then
            FolderBrowserDialog_DownloadDirectory.SelectedPath = My.Settings.DownloadDir
        Else
            If Not Directory.Exists(Application.StartupPath + "\Backgrounds") Then
                Directory.CreateDirectory(Application.StartupPath + "\Backgrounds")
            End If
            FolderBrowserDialog_DownloadDirectory.SelectedPath = Application.StartupPath + "\Backgrounds"
            My.Settings.DownloadDir = Application.StartupPath + "\Backgrounds"
        End If

        If My.Settings.RatioIndex = 0 And My.Settings.ResolutionIndex >= 0 Then
            My.Settings.ResolutionIndex = -1
        End If

        ComboBox_Sorting.SelectedIndex = 1
        CheckBox_AIArt.Checked = My.Settings.AIArt
        CheckBox_Anime.Checked = My.Settings.Anime
        ComboBox_Ratio.SelectedIndex = My.Settings.RatioIndex
        ComboBox_Resolution.SelectedIndex = My.Settings.ResolutionIndex
        TextBox_CustomResolutionWidth.Text = My.Settings.CustomResolutionWidth
        TextBox_CustomResolutionHeight.Text = My.Settings.CustomResolutionHeight
        Exactly_Resolution = My.Settings.ExactlyResolution
        If Exactly_Resolution = True Then
            Label_Exactly.Text = "Exactly"
        Else
            Label_Exactly.Text = "At Least"
        End If

        For i As Integer = 1 To 24
            ' Panel - Panel_Wallpaper
            Dim Panel_Wallpaper = New System.Windows.Forms.Panel
            Panel_Wallpaper.Size = New Size(300, 224)
            Panel_Wallpaper.BackColor = Color.FromArgb(28, 30, 34)
            Panel_Wallpaper.Name = "Panel_Wallpaper" & i.ToString()
            Panel_Wallpaper.Margin = New Padding(6, 6, 6, 6)


            ' PictureBox - PictureBox_Downloaded
            Dim PictureBox_Downloaded = New MyPictureBox
            PictureBox_Downloaded.Width = 24
            PictureBox_Downloaded.Height = 24
            PictureBox_Downloaded.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox_Downloaded.Name = "PictureBox_Downloaded" & i.ToString()
            PictureBox_Downloaded.Image = My.Resources.SteveZondicons_Checkmark_Outline
            PictureBox_Downloaded.BackColor = Color.Orange
            Panel_Wallpaper.Controls.Add(PictureBox_Downloaded)


            ' PictureBox - PictureBox_Wallpaper
            Dim PictureBox_Wallpaper = New MyPictureBox
            PictureBox_Wallpaper.Dock = DockStyle.Fill
            PictureBox_Wallpaper.SizeMode = PictureBoxSizeMode.CenterImage
            PictureBox_Wallpaper.Name = "PictureBox_Wallpaper" & i.ToString()
            PictureBox_Wallpaper.Cursor = Cursors.Hand
            AddHandler PictureBox_Wallpaper.MouseClick, AddressOf PictureBox_Wallpaper_MouseClick
            Panel_Wallpaper.Controls.Add(PictureBox_Wallpaper)


            ' Panel - Panel_Bottom
            Dim Panel_Bottom = New System.Windows.Forms.Panel()
            Panel_Bottom.Height = 24
            Panel_Bottom.BackColor = Color.FromArgb(28, 30, 34)
            Panel_Bottom.Name = "Panel_Bottom" & i.ToString()
            Panel_Bottom.Dock = DockStyle.Bottom
            Panel_Wallpaper.Controls.Add(Panel_Bottom)


            ' Label - Label_Resolution
            Dim Label_Resolution = New System.Windows.Forms.Label
            Label_Resolution.Font = New Font("Microsoft Sans Serif", 10, FontStyle.Italic)
            Label_Resolution.ForeColor = Color.WhiteSmoke
            Label_Resolution.TextAlign = ContentAlignment.MiddleCenter
            Label_Resolution.Dock = DockStyle.Fill
            Label_Resolution.Name = "Label_Resolution" & i.ToString()
            Panel_Bottom.Controls.Add(Label_Resolution)


            ' Button - Button_Download
            Dim Button_Download = New MyButton
            Button_Download.Image = My.Resources.SteveZondicons_Arrow_Down_Thick
            Button_Download.BackColor = Color.MediumSeaGreen
            Button_Download.Dock = DockStyle.Left
            Button_Download.FlatStyle = FlatStyle.Flat
            Button_Download.FlatAppearance.BorderSize = 0
            Button_Download.Width = 24
            Button_Download.Name = "Button_Download" & i.ToString()
            AddHandler Button_Download.MouseClick, AddressOf PictureBox_Download_MouseClick
            Panel_Bottom.Controls.Add(Button_Download)

            ' Add BackgroundPanel to FlowLayoutPanel1
            FlowLayoutPanel1.Controls.Add(Panel_Wallpaper)
        Next

        ' Wallhaven API URL with a search query
        GetWallpapersAsync(BuildApiUrl(TextBox_SearchQuery.Text))

        SettingsLoadLock = False
    End Sub

    ' GetWallpapersAsync - Updated to hide unused Panels
    Private Async Sub GetWallpapersAsync(apiUrl As String)
        ' Create an HttpClient instance
        Using client As New HttpClient()
            Dim retries As Integer = 0
            Dim maxRetries As Integer = 3
            Dim retryDelay As Integer = 60000 ' 60 seconds delay

            ' Get all Panel controls within FlowLayoutPanel1
            Dim panels As New List(Of Panel)

            For Each panel As Panel In FlowLayoutPanel1.Controls.OfType(Of Panel)()
                panels.Add(panel)
                panel.Controls(2).Controls(0).Text = ""
                panel.Controls(0).Visible = False
            Next

            ' Set loading image to PictureBoxes inside each panel
            For Each panel As Panel In panels
                Dim pictureBox As MyPictureBox = panel.Controls.OfType(Of MyPictureBox).ElementAt(1)
                If pictureBox IsNot Nothing Then
                    pictureBox.Image = My.Resources.Loading
                End If
            Next

            While retries < maxRetries
                Try
                    ' Send the HTTP GET request
                    Dim response As HttpResponseMessage = Await client.GetAsync(apiUrl)

                    ' Handle 429 Too Many Requests error
                    If response.StatusCode = 429 Then
                        Console.WriteLine("Rate limit hit: Too many requests. Waiting 60 seconds before retrying...")
                        Await Task.Delay(retryDelay)
                        retries += 1
                        Continue While
                    End If

                    ' Ensure success status code
                    response.EnsureSuccessStatusCode()

                    ' Read the response content as a string
                    Dim jsonResponse As String = Await response.Content.ReadAsStringAsync()

                    ' Parse the JSON response
                    Dim wallpaperResponse As WallhavenApiResponse = JsonConvert.DeserializeObject(Of WallhavenApiResponse)(jsonResponse)

                    ' Update pagination labels
                    Pages_current_page = wallpaperResponse.meta.current_page
                    Pages_last_page = wallpaperResponse.meta.last_page
                    Pages_total = wallpaperResponse.meta.total

                    Label_Pages.Text = Pages_current_page & "/" & Pages_last_page
                    NumericUpDown_Page.Value = Pages_current_page
                    NumericUpDown_Page.Maximum = Pages_last_page


                    ' Hide the Panels that are not used
                    For i As Integer = wallpaperResponse.data.Count To panels.Count - 1
                        panels(i).Visible = False ' Hide unused Panels
                    Next


                    If Not Directory.Exists(FolderBrowserDialog_DownloadDirectory.SelectedPath) Then
                        If Not Directory.Exists(Application.StartupPath + "\Backgrounds") Then
                            Directory.CreateDirectory(Application.StartupPath + "\Backgrounds")
                        End If
                        FolderBrowserDialog_DownloadDirectory.SelectedPath = Application.StartupPath + "\Backgrounds"
                        My.Settings.DownloadDir = Application.StartupPath + "\Backgrounds"
                    End If

                    ' Dictionary to store file names and their full paths
                    Dim FilesToCheck As New Dictionary(Of String, String)

                    ' Get all matching file paths
                    Dim filePaths = Directory.GetFiles(FolderBrowserDialog_DownloadDirectory.SelectedPath, "*.png", SearchOption.AllDirectories).ToList()
                    filePaths.AddRange(Directory.GetFiles(FolderBrowserDialog_DownloadDirectory.SelectedPath, "*.jpg", SearchOption.AllDirectories))
                    filePaths.AddRange(Directory.GetFiles(FolderBrowserDialog_DownloadDirectory.SelectedPath, "*.jpeg", SearchOption.AllDirectories))

                    ' Store file names as keys and full paths as values
                    For Each filePath In filePaths
                        Dim fileName As String = Path.GetFileName(filePath)
                        If Not FilesToCheck.ContainsKey(fileName) Then
                            FilesToCheck.Add(fileName, filePath) ' Store full path for later retrieval
                        End If
                    Next


                    ' Process and display the wallpapers
                    For i As Integer = 0 To Math.Min(wallpaperResponse.data.Count - 1, panels.Count - 1)
                        Dim wallpaper = wallpaperResponse.data(i)
                        Dim Panel_Wallpaper = panels(i)
                        Dim PictureBox_Wallpaper As MyPictureBox = Panel_Wallpaper.Controls.OfType(Of MyPictureBox).ElementAt(1)

                        Panel_Wallpaper.Controls(2).Controls(0).Text = wallpaper.resolution

                        ' Download the image from the path (wallpaper.thumbs.small)
                        Dim imageBytes As Byte() = Await client.GetByteArrayAsync(wallpaper.thumbs.small)

                        ' Load the image into a MemoryStream and assign it to the PictureBox
                        Using ms As New MemoryStream(imageBytes)
                            If PictureBox_Wallpaper IsNot Nothing Then
                                PictureBox_Wallpaper.Image = Image.FromStream(ms)
                            End If
                        End Using

                        Dim Button_Download As MyButton = TryCast(Panel_Wallpaper.Controls(2).Controls(1), MyButton)
                        Button_Download.wallhaven_path = wallpaper.path

                        If FilesToCheck.ContainsKey(Path.GetFileName(wallpaper.path)) Then
                            Button_Download.Visible = False
                            Panel_Wallpaper.Controls(0).Visible = True
                            PictureBox_Wallpaper.wallhaven_downloaded = True
                            ' Retrieve the full file path from the dictionary
                            PictureBox_Wallpaper.wallhaven_path = FilesToCheck(Path.GetFileName(wallpaper.path))
                        Else
                            Button_Download.Visible = True
                            Panel_Wallpaper.Controls(0).Visible = False
                            PictureBox_Wallpaper.wallhaven_downloaded = False
                            PictureBox_Wallpaper.wallhaven_path = wallpaper.path
                        End If

                        Panel_Wallpaper.Visible = True  ' Ensure the used Panel is visible
                    Next

                    Exit While

                Catch ex As Exception
                    Console.WriteLine("Error occurred: " & ex.Message)
                    retries += 1
                    If retries >= maxRetries Then Exit While
                End Try
            End While
        End Using
    End Sub

    ' PictureBox_Wallpaper - MouseClick
    Private Async Sub PictureBox_Wallpaper_MouseClick(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            Try

                If Not Directory.Exists(FolderBrowserDialog_DownloadDirectory.SelectedPath) Then
                    If Not Directory.Exists(Application.StartupPath + "\Backgrounds") Then
                        Directory.CreateDirectory(Application.StartupPath + "\Backgrounds")
                    End If
                    FolderBrowserDialog_DownloadDirectory.SelectedPath = Application.StartupPath + "\Backgrounds"
                    My.Settings.DownloadDir = Application.StartupPath + "\Backgrounds"
                End If

                ' Ensure the sender is a MyPictureBox
                Dim pb As MyPictureBox = TryCast(sender, MyPictureBox)
                If pb IsNot Nothing AndAlso Not String.IsNullOrEmpty(pb.wallhaven_path) Then
                    If pb.wallhaven_downloaded = True Then
                        Dim result As Boolean = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, pb.wallhaven_path, SPIF_UPDATEINIFILE Or SPIF_SENDCHANGE)

                        ' Check if the wallpaper was set successfully
                        If result Then
                            Console.WriteLine("Wallpaper set successfully from local!")
                        Else
                            Console.WriteLine("Failed to set wallpaper.")
                        End If
                    Else
                        ' Create an HttpClient to download the image from wallhaven_path
                        Using client As New HttpClient()
                            ' Download the image bytes
                            Dim imageBytes As Byte() = Await client.GetByteArrayAsync(pb.wallhaven_path)

                            ' Create a file path to save the image (e.g., in the user's temp directory)
                            Dim wallpaperPath As String = Path.Combine(Path.GetTempPath(), "wallpaper.bmp")

                            ' Save the downloaded image bytes as a BMP file (BMP is required for wallpapers on older systems)
                            Using ms As New MemoryStream(imageBytes)
                                Dim img As Image = Image.FromStream(ms)
                                img.Save(wallpaperPath) ', System.Drawing.Imaging.ImageFormat.Bmp)
                            End Using

                            ' Set the wallpaper using the Windows API
                            Dim result As Boolean = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, wallpaperPath, SPIF_UPDATEINIFILE Or SPIF_SENDCHANGE)

                            ' Check if the wallpaper was set successfully
                            If result Then
                                Console.WriteLine("Wallpaper set successfully!")
                            Else
                                Console.WriteLine("Failed to set wallpaper.")
                            End If
                        End Using
                    End If

                    If Me.WindowState = FormWindowState.Maximized Then
                        FlowLayoutPanel1.Visible = False
                        Timer_FadeOut.Start()
                    End If

                End If
            Catch ex As Exception
                MsgBox("Error: " & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub

    'PictureBox_Download - MouseClick
    Private Async Sub PictureBox_Download_MouseClick(sender As Object, e As MouseEventArgs)
        Dim control As Control = TryCast(sender, Control)
        If control IsNot Nothing AndAlso TypeOf control.Parent Is Panel Then

            If Not Directory.Exists(FolderBrowserDialog_DownloadDirectory.SelectedPath) Then
                If Not Directory.Exists(Application.StartupPath + "\Backgrounds") Then
                    Directory.CreateDirectory(Application.StartupPath + "\Backgrounds")
                End If
                FolderBrowserDialog_DownloadDirectory.SelectedPath = Application.StartupPath + "\Backgrounds"
                My.Settings.DownloadDir = Application.StartupPath + "\Backgrounds"
            End If

            Dim fileurl As String = DirectCast(sender, MyButton).wallhaven_path
            Dim filename As String = Path.GetFileName(fileurl)

            ' Retrieve the parent Panel
            Dim ownerPanelBottom As Panel = DirectCast(control.Parent, Panel)
            Dim ownerPanelMain As Panel = CType(ownerPanelBottom.Parent, Panel)

            Using client As New HttpClient()
                ' Download the image bytes
                Dim imageBytes As Byte() = Await client.GetByteArrayAsync(fileurl)

                ' Create a file path to save the image (e.g., in the user's temp directory)
                Dim wallpaperPath As String = Path.Combine(FolderBrowserDialog_DownloadDirectory.SelectedPath, filename)

                ' Save the downloaded image bytes
                Using ms As New MemoryStream(imageBytes)
                    Dim img As Image = Image.FromStream(ms)
                    img.Save(wallpaperPath)
                End Using
            End Using

            DirectCast(ownerPanelMain.Controls(1), MyPictureBox).wallhaven_path = Path.Combine(FolderBrowserDialog_DownloadDirectory.SelectedPath, filename)
            DirectCast(ownerPanelMain.Controls(1), MyPictureBox).wallhaven_downloaded = True
            DirectCast(ownerPanelMain.Controls(0), MyPictureBox).Visible = True
            DirectCast(sender, MyButton).Visible = False
        Else
            MessageBox.Show("Unable to find the Owner Panel.")
        End If
    End Sub

    'Form1 - Resize
    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Maximized Then
            FlowLayoutPanel1.AutoScroll = False
            FlowLayoutPanel1.PerformLayout()
            FlowLayoutPanel1.AutoScroll = True
        End If
    End Sub

    ' Form1 - ResizeEnd
    'Private Sub Form1_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
    'Console.WriteLine(Me.Size.ToString)
    'End Sub

    ' Button_DownloadDirectory - Click
    Private Sub Button_DownloadDirectory_Click(sender As Object, e As EventArgs) Handles Button_DownloadDirectory.Click
        If FolderBrowserDialog_DownloadDirectory.ShowDialog() = DialogResult.OK Then
            My.Settings.DownloadDir = FolderBrowserDialog_DownloadDirectory.SelectedPath
        End If
    End Sub

    ' Button_Prev - Click
    Private Sub Button_Prev_Click(sender As Object, e As EventArgs) Handles Button_Prev.Click
        If Not Pages_current_page = 1 Then
            GetWallpapersAsync(BuildApiUrl(TextBox_SearchQuery.Text, Pages_current_page - 1))
        End If
    End Sub

    ' Button_Next - Click
    Private Sub Button_Next_Click(sender As Object, e As EventArgs) Handles Button_Next.Click
        If Not Pages_current_page = Pages_last_page Then
            GetWallpapersAsync(BuildApiUrl(TextBox_SearchQuery.Text, Pages_current_page + 1))
        End If
    End Sub

    ' Button_Search - Click
    Private Sub Button_Search_Click(sender As Object, e As EventArgs) Handles Button_Search.Click
        If Not NumericUpDown_Page.Value = Pages_current_page Then
            GetWallpapersAsync(BuildApiUrl(TextBox_SearchQuery.Text, CType(NumericUpDown_Page.Value, Integer?)))
        Else
            GetWallpapersAsync(BuildApiUrl(TextBox_SearchQuery.Text))
        End If
    End Sub

    ' TextBox_SearchQuery - KeyDown
    Private Sub TextBox_SearchQuery_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_SearchQuery.KeyDown
        If e.KeyCode = Keys.Enter Then
            ' Prevent the "ding" sound on Enter key press
            e.SuppressKeyPress = True
            FlowLayoutPanel1.Select()

            ' Call GetWallpapersAsync with the search query
            Dim query As String = TextBox_SearchQuery.Text.Trim()
            If Not String.IsNullOrWhiteSpace(query) Then
                ' Build the API URL based on the search query
                Dim apiUrl As String = BuildApiUrl(query:=query)
                ' Call the asynchronous method
                GetWallpapersAsync(apiUrl)
            End If
        End If
    End Sub

    'NumericUpDown_Page - KeyDown
    Private Sub NumericUpDown_Page_KeyDown(sender As Object, e As KeyEventArgs) Handles NumericUpDown_Page.KeyDown
        If e.KeyCode = Keys.Enter Then
            ' Prevent the "ding" sound on Enter key press
            e.SuppressKeyPress = True
            FlowLayoutPanel1.Select()
            NumericUpDown_Page.Hide()
            GetWallpapersAsync(BuildApiUrl(TextBox_SearchQuery.Text, CType(NumericUpDown_Page.Value, Integer?)))
        End If
    End Sub

    ' BuildApiUrl
    Private Function BuildApiUrl(Optional query As String = "", Optional page As Integer? = Nothing) As String
        ' Base API URL with static parameters
        Dim apiUrl As String = "https://wallhaven.cc/api/v1/search?"

        ' Add query if provided
        If Not String.IsNullOrWhiteSpace(query) Then
            apiUrl &= $"q={Uri.EscapeDataString(query)}&"
        End If

        ' categories & purity
        If CheckBox_Anime.Checked Then
            apiUrl &= "categories=110&purity=100&"
        Else
            apiUrl &= "categories=100&purity=100&"
        End If

        ' Add resolutions if provided
        Dim resolution As String = ""
        If ComboBox_Ratio.SelectedIndex = 0 Then
            resolution = TextBox_CustomResolutionWidth.Text & "x" & TextBox_CustomResolutionHeight.Text
        Else
            resolution = ComboBox_Resolution.Items.Item(ComboBox_Resolution.SelectedIndex).ToString().Replace(" ", "")
        End If

        If Not String.IsNullOrWhiteSpace(resolution) Then
            If Exactly_Resolution Then
                apiUrl &= $"resolutions={resolution}&"
            Else
                apiUrl &= $"atleast={resolution}&"
            End If
        End If

        ' sorting
        Dim sorting As String = ComboBox_Sorting.SelectedItem?.ToString().ToLower() ' Get the selected sorting value
        If Not String.IsNullOrWhiteSpace(sorting) Then
            If sorting = "date added" Then
                sorting = "date_added"
            End If
            apiUrl &= $"sorting={Uri.EscapeDataString(sorting)}&"
        End If

        ' order
        If Descending_sorting_order = True Then
            apiUrl &= "order=desc&"
        Else
            apiUrl &= "order=asc&"
        End If

        ' Add colors if provided
        If Not Label_SelectedColor.Text = "No Color" Then
            apiUrl &= $"colors={Label_SelectedColor.Text.Substring(1)}&"
        End If

        ' ai_art_filter
        If CheckBox_AIArt.Checked Then
            apiUrl &= "ai_art_filter=0&"
        Else
            apiUrl &= "ai_art_filter=1&"
        End If

        ' Add page if provided
        If page.HasValue Then
            apiUrl &= $"page={page.Value}&"
        End If

        ' Remove the trailing '&' if any
        If apiUrl.EndsWith("&") Then
            apiUrl = apiUrl.TrimEnd("&"c)
        End If

        Console.WriteLine(apiUrl)
        Return apiUrl
    End Function

    ' Label_Color - MouseClick
    Private Sub Label_Color_MouseClick(sender As Object, e As MouseEventArgs) Handles Label_SelectedColor.MouseClick
        ContextMenuStrip_Colors.Show(Label_SelectedColor, e.Location)
    End Sub

    ' ComboBox_Ratio - SelectedIndexChanged
    Private Sub ComboBox_Ratio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Ratio.SelectedIndexChanged
        ComboBox_Resolution.SuspendLayout()
        ComboBox_Resolution.Items.Clear()

        If ComboBox_Ratio.SelectedIndex = 0 Then
            ComboBox_Resolution.Visible = False
            TextBox_CustomResolutionWidth.Visible = True
            Label_CustomResolutionX.Visible = True
            TextBox_CustomResolutionHeight.Visible = True
        ElseIf ComboBox_Ratio.SelectedIndex = 1 Then
            ComboBox_Resolution.Items.Add("2560 x 1080")
            ComboBox_Resolution.Items.Add("3440 x 1440")
            ComboBox_Resolution.Items.Add("3840 x 1600")
        ElseIf ComboBox_Ratio.SelectedIndex = 2 Then
            ComboBox_Resolution.Items.Add("1280 x 720")
            ComboBox_Resolution.Items.Add("1600 x 900")
            ComboBox_Resolution.Items.Add("1920 x 1080")
            ComboBox_Resolution.Items.Add("2560 x 1440")
            ComboBox_Resolution.Items.Add("3840 x 2160")
        ElseIf ComboBox_Ratio.SelectedIndex = 3 Then
            ComboBox_Resolution.Items.Add("1280x 800")
            ComboBox_Resolution.Items.Add("1600 x 1000")
            ComboBox_Resolution.Items.Add("1920 x 1200")
            ComboBox_Resolution.Items.Add("2560 x 1600")
            ComboBox_Resolution.Items.Add("3840 x 2400")
        ElseIf ComboBox_Ratio.SelectedIndex = 4 Then
            ComboBox_Resolution.Items.Add("1280 x 960")
            ComboBox_Resolution.Items.Add("1600 x 1200")
            ComboBox_Resolution.Items.Add("1920 x 1440")
            ComboBox_Resolution.Items.Add("2560 x 1920")
            ComboBox_Resolution.Items.Add("3840 x 2880")
        ElseIf ComboBox_Ratio.SelectedIndex = 5 Then
            ComboBox_Resolution.Items.Add("1280 x 1024")
            ComboBox_Resolution.Items.Add("1600 x 1280")
            ComboBox_Resolution.Items.Add("1920 x 1536")
            ComboBox_Resolution.Items.Add("2560 x 2048")
            ComboBox_Resolution.Items.Add("3840 x 3072")
        End If

        If SettingsLoadLock = False Then
            If Not ComboBox_Ratio.SelectedIndex = 0 Then
                ComboBox_Resolution.Visible = True
                TextBox_CustomResolutionWidth.Visible = False
                Label_CustomResolutionX.Visible = False
                TextBox_CustomResolutionHeight.Visible = False
                ComboBox_Resolution.SelectedIndex = 0
            End If
        End If

        ComboBox_Resolution.ResumeLayout()

        'If Not ComboBox_Ratio.SelectedIndex = -1 Then
        My.Settings.RatioIndex = ComboBox_Ratio.SelectedIndex
        'End If
    End Sub

    ' ComboBox_Resolution - SelectedIndexChanged
    Private Sub ComboBox_Resolution_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_Resolution.SelectedIndexChanged
        If Not ComboBox_Resolution.SelectedIndex = -1 Then
            My.Settings.ResolutionIndex = ComboBox_Resolution.SelectedIndex
        End If
    End Sub

    ' Button_SortingOrder - Click
    Private Sub Button_SortingOrder_Click(sender As Object, e As EventArgs) Handles Button_SortingOrder.Click
        If Descending_sorting_order = True Then
            Descending_sorting_order = False
            Button_SortingOrder.Image = My.Resources.SteveZondicons_Arrow_Up
        Else
            Descending_sorting_order = True
            Button_SortingOrder.Image = My.Resources.SteveZondicons_Arrow_Down
        End If
    End Sub

    ' Label_Exactly - Click
    Private Sub Label_Exactly_Click(sender As Object, e As EventArgs) Handles Label_Exactly.Click
        If Exactly_Resolution = True Then
            Exactly_Resolution = False
            Label_Exactly.Text = "At Least"
        Else
            Exactly_Resolution = True
            Label_Exactly.Text = "Exactly"
        End If
        My.Settings.ExactlyResolution = Exactly_Resolution
    End Sub

    ' Label_Pages - Click
    Private Sub Label_Pages_Click(sender As Object, e As EventArgs) Handles Label_Pages.Click
        If NumericUpDown_Page.Visible = True Then
            NumericUpDown_Page.Visible = False
        Else
            NumericUpDown_Page.Visible = True
        End If
    End Sub

    ' Form1 - KeyUp
    Private Sub Form1_KeyUp(sender As Object, e As KeyEventArgs) Handles Me.KeyUp
        If Not TextBox_SearchQuery.Focused And Not ComboBox_Sorting.Focused And Not ComboBox_Ratio.Focused And Not ComboBox_Resolution.Focused And Not TextBox_CustomResolutionWidth.Focused And Not TextBox_CustomResolutionHeight.Focused Then
            If e.KeyCode = Keys.PageUp Or e.KeyCode = Keys.Up Or e.KeyCode = Keys.Right Then
                Button_Next.PerformClick()
            ElseIf e.KeyCode = Keys.PageDown Or e.KeyCode = Keys.Down Or e.KeyCode = Keys.Left Then
                Button_Prev.PerformClick()
            End If
        End If
    End Sub

    ' ComboBox_Resolution - DropDownClosed
    Private Sub ComboBox_Resolution_DropDownClosed(sender As Object, e As EventArgs) Handles ComboBox_Resolution.DropDownClosed
        FlowLayoutPanel1.Select()
    End Sub

    ' ComboBox_Sorting - DropDownClosed
    Private Sub ComboBox_Sorting_DropDownClosed(sender As Object, e As EventArgs) Handles ComboBox_Sorting.DropDownClosed
        FlowLayoutPanel1.Select()
    End Sub

    ' ComboBox_Ratio - DropDownClosed
    Private Sub ComboBox_Ratio_DropDownClosed(sender As Object, e As EventArgs) Handles ComboBox_Ratio.DropDownClosed
        FlowLayoutPanel1.Select()
    End Sub

    ' TextBox_CustomResolutionWidth - KeyPress
    Private Sub TextBox_CustomResolutionWidth_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_CustomResolutionWidth.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True ' Block invalid input
        End If
    End Sub

    ' TextBox_CustomResolutionHeight - KeyPress
    Private Sub TextBox_CustomResolutionHeight_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_CustomResolutionHeight.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True ' Block invalid input
        End If
    End Sub

    ' TextBox_CustomResolutionWidth - TextChanged
    Private Sub TextBox_CustomResolutionWidth_TextChanged(sender As Object, e As EventArgs) Handles TextBox_CustomResolutionWidth.TextChanged
        ' Remove any non-numeric characters (handles pasting)
        Dim cleanText As String = ""
        For Each c As Char In TextBox_CustomResolutionWidth.Text
            If Char.IsDigit(c) Then
                cleanText &= c
            End If
        Next

        ' Update the TextBox only if it changed
        If TextBox_CustomResolutionWidth.Text <> cleanText Then
            TextBox_CustomResolutionWidth.Text = cleanText
            TextBox_CustomResolutionWidth.SelectionStart = TextBox_CustomResolutionWidth.Text.Length ' Keep cursor at the end
            My.Settings.CustomResolutionWidth = cleanText
        End If

        My.Settings.CustomResolutionWidth = TextBox_CustomResolutionWidth.Text
    End Sub

    ' TextBox_CustomResolutionHeight - TextChanged
    Private Sub TextBox_CustomResolutionHeight_TextChanged(sender As Object, e As EventArgs) Handles TextBox_CustomResolutionHeight.TextChanged
        ' Remove any non-numeric characters (handles pasting)
        Dim cleanText As String = ""
        For Each c As Char In TextBox_CustomResolutionHeight.Text
            If Char.IsDigit(c) Then
                cleanText &= c
            End If
        Next

        ' Update the TextBox only if it changed
        If TextBox_CustomResolutionHeight.Text <> cleanText Then
            TextBox_CustomResolutionHeight.Text = cleanText
            TextBox_CustomResolutionHeight.SelectionStart = TextBox_CustomResolutionHeight.Text.Length ' Keep cursor at the end
            My.Settings.CustomResolutionHeight = cleanText
        End If

        My.Settings.CustomResolutionHeight = TextBox_CustomResolutionHeight.Text
    End Sub

    ' CheckBox_AIArt - CheckedChanged
    Private Sub CheckBox_AIArt_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_AIArt.CheckedChanged
        My.Settings.AIArt = CheckBox_AIArt.Checked
    End Sub

    ' Timer_FadeOut - Tick
    Private Sub Timer_FadeOut_Tick(sender As Object, e As EventArgs) Handles Timer_FadeOut.Tick
        If Me.Opacity > 0.05 Then
            Me.Opacity -= 0.05 ' Decrease opacity gradually
        Else
            Timer_FadeOut.Stop() ' Stop when fully invisible
            Threading.Thread.Sleep(800)
            Timer_FadeIn.Start()
        End If
    End Sub

    ' Timer_FadeIn - Tick
    Private Sub Timer_FadeIn_Tick(sender As Object, e As EventArgs) Handles Timer_FadeIn.Tick
        If Me.Opacity < 0.95 Then
            Me.Opacity += 0.05 ' Increase opacity gradually
        Else
            FlowLayoutPanel1.Visible = True
            Timer_FadeIn.Stop() ' Stop when fully visible
        End If
    End Sub

    ' ContextMenuStrip_ODD - Opening
    Private Sub ContextMenuStrip_ODD_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip_ODD.Opening
        If Directory.Exists(FolderBrowserDialog_DownloadDirectory.SelectedPath) Then
            ToolStripMenuItem_OpenDownloadDirectory.Enabled = True
        Else
            ToolStripMenuItem_OpenDownloadDirectory.Enabled = False
        End If
    End Sub

    ' ToolStripMenuItem_OpenDownloadDirectory - Click
    Private Sub ToolStripMenuItem_OpenDownloadDirectory_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_OpenDownloadDirectory.Click
        Process.Start(FolderBrowserDialog_DownloadDirectory.SelectedPath)
    End Sub

    ' Button_DownloadDirectory - MouseEnter
    Private Sub Button_DownloadDirectory_MouseEnter(sender As Object, e As EventArgs) Handles Button_DownloadDirectory.MouseEnter
        If Directory.Exists(FolderBrowserDialog_DownloadDirectory.SelectedPath) Then
            ToolTip1.SetToolTip(Button_DownloadDirectory, "Download Directory: " + FolderBrowserDialog_DownloadDirectory.SelectedPath)
        Else
            ToolTip1.SetToolTip(Button_DownloadDirectory, "Download Directory: Not vaild")
        End If
    End Sub

    ' ToolTip1 - Draw
    Private Sub ToolTip1_Draw(sender As Object, e As DrawToolTipEventArgs) Handles ToolTip1.Draw
        ' Create a brush with the custom background color
        Dim bgBrush As New SolidBrush(Color.FromArgb(28, 30, 34))
        Dim textBrush As New SolidBrush(Color.DodgerBlue)
        Dim font As New Font("Arial", 10, FontStyle.Italic)

        ' Fill the tooltip background
        e.Graphics.FillRectangle(bgBrush, e.Bounds)

        ' Draw the tooltip text
        e.Graphics.DrawString(e.ToolTipText, font, textBrush, e.Bounds.X + 4, e.Bounds.Y + 2)

        ' Dispose of brushes
        bgBrush.Dispose()
        textBrush.Dispose()
    End Sub

    ' ToolTip1 - Popup
    Private Sub ToolTip1_Popup(sender As Object, e As PopupEventArgs) Handles ToolTip1.Popup
        ' Adjust tooltip size if needed
        e.ToolTipSize = TextRenderer.MeasureText(ToolTip1.GetToolTip(e.AssociatedControl), New Font("Arial", 10, FontStyle.Italic))
        e.ToolTipSize = New Size(e.ToolTipSize.Width, e.ToolTipSize.Height + 8)
    End Sub

    ' CheckBox_Anime - CheckedChanged
    Private Sub CheckBox_Anime_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_Anime.CheckedChanged
        My.Settings.Anime = CheckBox_Anime.Checked
    End Sub

    ' Button_WallpaperStyles - MouseClick
    Private Sub Button_WallpaperStyles_MouseClick(sender As Object, e As MouseEventArgs) Handles Button_WallpaperStyles.MouseClick
        If e.Button = MouseButtons.Left Then
            ContextMenuStrip_WallpaperStyles.Show(Button_WallpaperStyles, e.Location)
        End If
    End Sub

    ' SetWallpaperStyle(style)
    Sub SetWallpaperStyle(style As String)
        Try
            ' Open Registry Key
            Dim regKey As RegistryKey = Registry.CurrentUser.OpenSubKey("Control Panel\Desktop", True)

            ' Set style based on input
            Select Case style.ToLower()
                Case "tiled"
                    regKey.SetValue("WallpaperStyle", "0")
                    regKey.SetValue("TileWallpaper", "1")
                Case "centered"
                    regKey.SetValue("WallpaperStyle", "0")
                    regKey.SetValue("TileWallpaper", "0")
                Case "stretched"
                    regKey.SetValue("WallpaperStyle", "2")
                    regKey.SetValue("TileWallpaper", "0")
                Case "fit"
                    regKey.SetValue("WallpaperStyle", "6")
                    regKey.SetValue("TileWallpaper", "0")
                Case "fill"
                    regKey.SetValue("WallpaperStyle", "10")
                    regKey.SetValue("TileWallpaper", "0")
                Case Else
                    Throw New ArgumentException("Invalid wallpaper style. Use: tiled, centered, stretched, fit, fill.")
            End Select

            regKey.Close()

            ' Refresh desktop by reapplying the current wallpaper
            Dim currentWallpaper As String = Registry.GetValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Wallpaper", "").ToString()
            SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, currentWallpaper, SPIF_UPDATEINIFILE Or SPIF_SENDCHANGE)

            Console.WriteLine("Wallpaper style updated successfully!")
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        End Try
    End Sub

    ' ToolStripMenuItem_StyleTiled_Click
    Private Sub ToolStripMenuItem_StyleTiled_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_StyleTiled.Click
        SetWallpaperStyle("tiled")
    End Sub

    ' ToolStripMenuItem_StyleCentered - Click
    Private Sub ToolStripMenuItem_StyleCentered_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_StyleCentered.Click
        SetWallpaperStyle("centered")
    End Sub

    ' ToolStripMenuItem_StyleStretched - Click
    Private Sub ToolStripMenuItem_StyleStretched_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_StyleStretched.Click
        SetWallpaperStyle("stretched")
    End Sub

    ' ToolStripMenuItem_StyleFit - Click
    Private Sub ToolStripMenuItem_StyleFit_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_StyleFit.Click
        SetWallpaperStyle("fit")
    End Sub

    ' ToolStripMenuItem_StyleFill - Click
    Private Sub ToolStripMenuItem_StyleFill_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_StyleFill.Click
        SetWallpaperStyle("fill")
    End Sub

    ' ToolStripMenuItem_NoColor - Click
    Private Sub ToolStripMenuItem_NoColor_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_NoColor.Click
        Label_SelectedColor.Text = "No Color"
        Label_SelectedColor.ForeColor = Color.WhiteSmoke
    End Sub

    ' ToolStripMenuItem_660000 - Click
    Private Sub ToolStripMenuItem_660000_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_660000.Click
        Label_SelectedColor.Text = "#660000"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_660000.BackColor
    End Sub

    ' ToolStripMenuItem_990000 - Click
    Private Sub ToolStripMenuItem_990000_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_990000.Click
        Label_SelectedColor.Text = "#990000"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_990000.BackColor
    End Sub

    ' ToolStripMenuItem_cc0000 - Click
    Private Sub ToolStripMenuItem_cc0000_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_cc0000.Click
        Label_SelectedColor.Text = "#cc0000"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_cc0000.BackColor
    End Sub

    ' ToolStripMenuItem_cc3333 - Click
    Private Sub ToolStripMenuItem_cc3333_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_cc3333.Click
        Label_SelectedColor.Text = "#cc3333"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_cc3333.BackColor
    End Sub

    ' ToolStripMenuItem_ea4c88 - Click
    Private Sub ToolStripMenuItem_ea4c88_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_ea4c88.Click
        Label_SelectedColor.Text = "#ea4c88"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_ea4c88.BackColor
    End Sub

    ' ToolStripMenuItem_993399 - Click
    Private Sub ToolStripMenuItem_993399_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_993399.Click
        Label_SelectedColor.Text = "#993399"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_993399.BackColor
    End Sub

    ' ToolStripMenuItem_663399 - Click
    Private Sub ToolStripMenuItem_663399_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_663399.Click
        Label_SelectedColor.Text = "#663399"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_663399.BackColor
    End Sub

    ' ToolStripMenuItem_333399 - Click
    Private Sub ToolStripMenuItem_333399_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_333399.Click
        Label_SelectedColor.Text = "#333399"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_333399.BackColor
    End Sub

    ' ToolStripMenuItem_0066cc - Click
    Private Sub ToolStripMenuItem_0066cc_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_0066cc.Click
        Label_SelectedColor.Text = "#0066cc"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_0066cc.BackColor
    End Sub

    ' ToolStripMenuItem_0099cc - Click
    Private Sub ToolStripMenuItem_0099cc_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_0099cc.Click
        Label_SelectedColor.Text = "#0099cc"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_0099cc.BackColor
    End Sub

    ' ToolStripMenuItem_66cccc - Click
    Private Sub ToolStripMenuItem_66cccc_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_66cccc.Click
        Label_SelectedColor.Text = "#66cccc"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_66cccc.BackColor
    End Sub

    ' ToolStripMenuItem_77cc33 - Click
    Private Sub ToolStripMenuItem_77cc33_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_77cc33.Click
        Label_SelectedColor.Text = "#77cc33"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_77cc33.BackColor
    End Sub

    ' ToolStripMenuItem_669900 - Click
    Private Sub ToolStripMenuItem_669900_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_669900.Click
        Label_SelectedColor.Text = "#669900"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_669900.BackColor
    End Sub

    ' ToolStripMenuItem_336600 - Click
    Private Sub ToolStripMenuItem_336600_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_336600.Click
        Label_SelectedColor.Text = "#336600"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_336600.BackColor
    End Sub

    ' ToolStripMenuItem_666600 - Click
    Private Sub ToolStripMenuItem_666600_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_666600.Click
        Label_SelectedColor.Text = "#666600"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_666600.BackColor
    End Sub

    ' ToolStripMenuItem_999900 - Click
    Private Sub ToolStripMenuItem_999900_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_999900.Click
        Label_SelectedColor.Text = "#999900"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_999900.BackColor
    End Sub

    ' ToolStripMenuItem_cccc33 - Click
    Private Sub ToolStripMenuItem_cccc33_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_cccc33.Click
        Label_SelectedColor.Text = "#cccc33"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_cccc33.BackColor
    End Sub

    ' ToolStripMenuItem_ffff00 - Click
    Private Sub ToolStripMenuItem_ffff00_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_ffff00.Click
        Label_SelectedColor.Text = "#ffff00"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_ffff00.BackColor
    End Sub

    ' ToolStripMenuItem_ffcc33 - Click
    Private Sub ToolStripMenuItem_ffcc33_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_ffcc33.Click
        Label_SelectedColor.Text = "#ffcc33"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_ffcc33.BackColor
    End Sub

    ' ToolStripMenuItem_ff9900 - Click
    Private Sub ToolStripMenuItem_ff9900_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_ff9900.Click
        Label_SelectedColor.Text = "#ff9900"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_ff9900.BackColor
    End Sub

    ' ToolStripMenuItem_ff6600 - Click
    Private Sub ToolStripMenuItem_ff6600_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_ff6600.Click
        Label_SelectedColor.Text = "#ff6600"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_ff6600.BackColor
    End Sub

    ' ToolStripMenuItem_cc6633 - Click
    Private Sub ToolStripMenuItem_cc6633_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_cc6633.Click
        Label_SelectedColor.Text = "#cc6633"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_cc6633.BackColor
    End Sub

    ' ToolStripMenuItem_996633 - Click
    Private Sub ToolStripMenuItem_996633_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_996633.Click
        Label_SelectedColor.Text = "#996633"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_996633.BackColor
    End Sub

    ' ToolStripMenuItem_663300 - Click
    Private Sub ToolStripMenuItem_663300_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_663300.Click
        Label_SelectedColor.Text = "#663300"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_663300.BackColor
    End Sub

    ' ToolStripMenuItem_000000 - Click
    Private Sub ToolStripMenuItem_000000_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_000000.Click
        Label_SelectedColor.Text = "#000000"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_000000.BackColor
    End Sub

    ' ToolStripMenuItem_999999 - Click
    Private Sub ToolStripMenuItem_999999_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_999999.Click
        Label_SelectedColor.Text = "#999999"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_999999.BackColor
    End Sub

    ' ToolStripMenuItem_cccccc - Click
    Private Sub ToolStripMenuItem_cccccc_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_cccccc.Click
        Label_SelectedColor.Text = "#cccccc"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_cccccc.BackColor
    End Sub

    ' ToolStripMenuItem_ffffff - Click
    Private Sub ToolStripMenuItem_ffffff_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_ffffff.Click
        Label_SelectedColor.Text = "#ffffff"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_ffffff.BackColor
    End Sub

    ' ToolStripMenuItem_424153 - Click
    Private Sub ToolStripMenuItem_424153_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem_424153.Click
        Label_SelectedColor.Text = "#424153"
        Label_SelectedColor.ForeColor = ToolStripMenuItem_424153.BackColor
    End Sub
End Class

' Define the response model to match the Wallhaven API response structure
Public Class WallhavenApiResponse
    Public Property data As List(Of Wallpaper)
    Public Property meta As MetaInfo ' Meta information
End Class

' Define the Wallpaper class
Public Class Wallpaper
    Public Property id As String
    Public Property path As String
    Public Property resolution As String
    Public Property colors As List(Of String) ' Colors as a list of strings
    Public Property thumbs As Thumbnails ' Thumbnails object containing different sizes
End Class

' Define the Thumbnails class
Public Class Thumbnails
    Public Property large As String
    Public Property original As String
    Public Property small As String
End Class

' Define the MetaInfo class to match the "meta" field in the response
Public Class MetaInfo
    Public Property current_page As Integer
    Public Property last_page As Integer
    Public Property per_page As Integer
    Public Property total As Integer
    Public Property query As String
    Public Property seed As Object ' seed can be null, so we use Object
End Class

' MyPictureBox
Public Class MyPictureBox
    Inherits PictureBox
    Public Property wallhaven_path As String
    Public Property wallhaven_downloaded As Boolean
End Class

' MyButton
Public Class MyButton
    Inherits System.Windows.Forms.Button
    Public Property wallhaven_path As String
End Class