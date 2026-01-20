Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports NTwain
Imports NTwain.Data
Imports System.Drawing
Imports System.Collections.Generic
Imports System.IO
Imports System.Net.WebSockets
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports Image = System.Drawing.Image

Public Class Form1
    Private scannedImages As New List(Of System.Drawing.Image) ' Store multiple scanned images
    Private currentImageIndex As Integer = -1 ' Track current image in PictureBox
    Private WithEvents btnSend As New Button ' Button for saving to PDF and sending to WebSocket
    'Private WithEvents btnScan As New Button ' Button for scanning
    Private WithEvents pictureBox As New PictureBox ' PictureBox for image preview
    Private previewCard As Panel
    Private WithEvents btnZoomIn As New Button ' Button for zooming in
    Private WithEvents btnZoomOut As New Button ' Button for zooming out
    Private WithEvents btnNext As New Button ' Button for next image
    Private WithEvents btnPrev As New Button ' Button for previous image
    Private WithEvents progressBar As New ProgressBar ' Loading indicator
    Private WithEvents txtSavePath As New TextBox ' TextBox for custom save path
    Private WithEvents btnSetPath As New Button ' Button to apply custom path
    Private WithEvents lblImageInfo As New Label ' Image index display
    Private zoomLevel As Single = 1.0F ' Current zoom level for PictureBox
    Private savePath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) ' Default to Downloads
    Private loaNo As String = ""
    Private transactionNo As String = ""
    Private batchNo As String = ""
    Private userToken As String = ""

    ' Helper function to resize icons
    Private Function ResizeImage(sourceImage As Image, targetSize As Size) As Image
        Dim resizedImage As New Bitmap(targetSize.Width, targetSize.Height)
        Using g As Graphics = Graphics.FromImage(resizedImage)
            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            g.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
            g.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
            g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
            g.DrawImage(sourceImage, 0, 0, targetSize.Width, targetSize.Height)
        End Using
        Return resizedImage
    End Function

    Private Async Function SendToWebSocket(message As String) As Task
        Try
            Dim baseUri As String = "wss://yourwebsocketserver.com/scan_endpoint"
            If String.IsNullOrEmpty(userToken) Then
                baseUri = baseUri
            Else
                baseUri = baseUri & "?token=" & Uri.EscapeDataString(userToken)
            End If
            Dim wsUri As New Uri(baseUri)
            Using ws As New ClientWebSocket()
                Await ws.ConnectAsync(wsUri, CancellationToken.None)
                Dim buffer As Byte() = Encoding.UTF8.GetBytes(message)
                Await ws.SendAsync(New ArraySegment(Of Byte)(buffer), WebSocketMessageType.Text, True, CancellationToken.None)
                Await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "Done", CancellationToken.None)
            End Using
            Me.Invoke(Sub() MessageBox.Show("Websocket IN!"))
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error sa WebSocket: " & ex.Message))
        End Try
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        EnsureCustomProtocolRegistered()
        ShowLaunchParameters()
        ' Set form size
        Me.Size = New Size(900, 650)
        Me.MinimumSize = New Size(700, 500) ' Ensure form is resizable but not too small
        Me.BackColor = Color.FromArgb(243, 244, 246)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Regular)

        Dim mainLayout As New TableLayoutPanel()
        mainLayout.Dock = DockStyle.Fill
        mainLayout.Padding = New Padding(16)
        mainLayout.ColumnCount = 1
        mainLayout.RowCount = 3
        mainLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 70.0F))
        mainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 30.0F))
        Me.Controls.Add(mainLayout)

        Dim headerPanel As New FlowLayoutPanel()
        headerPanel.FlowDirection = FlowDirection.TopDown
        headerPanel.WrapContents = False
        headerPanel.AutoSize = True
        headerPanel.Dock = DockStyle.Top

        Dim lblTitle As New Label()
        lblTitle.Text = "Custom Scan App"
        lblTitle.Font = New System.Drawing.Font("Segoe UI Semibold", 16.0F, System.Drawing.FontStyle.Bold)
        lblTitle.ForeColor = Color.FromArgb(17, 24, 39)
        lblTitle.AutoSize = True

        Dim lblSubtitle As New Label()
        lblSubtitle.Text = "Scan documents and send as PDF"
        lblSubtitle.Font = New System.Drawing.Font("Segoe UI", 10.0F, System.Drawing.FontStyle.Regular)
        lblSubtitle.ForeColor = Color.FromArgb(107, 114, 128)
        lblSubtitle.AutoSize = True

        headerPanel.Controls.Add(lblTitle)
        headerPanel.Controls.Add(lblSubtitle)
        mainLayout.Controls.Add(headerPanel, 0, 0)

        previewCard = New Panel()
        previewCard.Dock = DockStyle.Fill
        previewCard.BackColor = Color.White
        previewCard.Padding = New Padding(12)
        previewCard.BorderStyle = BorderStyle.FixedSingle
        previewCard.AutoScroll = True

        pictureBox.Dock = DockStyle.None
        pictureBox.SizeMode = PictureBoxSizeMode.Zoom
        pictureBox.BorderStyle = BorderStyle.None
        pictureBox.BackColor = Color.FromArgb(249, 250, 251)
        pictureBox.Location = New Point(0, 0)
        pictureBox.Size = New Size(600, 400)
        previewCard.Controls.Add(pictureBox)
        AddHandler previewCard.SizeChanged, Sub()
                                                UpdatePreviewSize()
                                            End Sub
        mainLayout.Controls.Add(previewCard, 0, 1)

        ' Setup Scan button
        btnScan.Text = "Scan"
        ApplyPrimaryButtonStyle(btnScan, Color.FromArgb(79, 70, 229), 120)
        Try
            Dim icon As Image = Image.FromFile(Path.Combine(Application.StartupPath, "Icons", "scan.png"))
            btnScan.Image = ResizeImage(icon, New Size(20, 20))
            icon.Dispose()
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error loading scan icon: " & ex.Message))
        End Try

        ' Setup Send button
        btnSend.Text = "Save as PDF and Send"
        ApplyPrimaryButtonStyle(btnSend, Color.FromArgb(16, 185, 129), 190)
        Try
            Dim icon As Image = Image.FromFile(Path.Combine(Application.StartupPath, "Icons", "save.png"))
            btnSend.Image = ResizeImage(icon, New Size(20, 20))
            icon.Dispose()
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error loading save icon: " & ex.Message))
        End Try

        ' Setup Zoom In button
        btnZoomIn.Text = "Zoom In"
        ApplySecondaryButtonStyle(btnZoomIn, 120)
        Try
            Dim icon As Image = Image.FromFile(Path.Combine(Application.StartupPath, "Icons", "zoom_in.png"))
            btnZoomIn.Image = ResizeImage(icon, New Size(18, 18))
            icon.Dispose()
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error loading zoom in icon: " & ex.Message))
        End Try

        ' Setup Zoom Out button
        btnZoomOut.Text = "Zoom Out"
        ApplySecondaryButtonStyle(btnZoomOut, 120)
        Try
            Dim icon As Image = Image.FromFile(Path.Combine(Application.StartupPath, "Icons", "zoom_out.png"))
            btnZoomOut.Image = ResizeImage(icon, New Size(18, 18))
            icon.Dispose()
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error loading zoom out icon: " & ex.Message))
        End Try

        ' Setup Previous button
        btnPrev.Text = "Previous"
        ApplySecondaryButtonStyle(btnPrev, 120)
        Try
            Dim icon As Image = Image.FromFile(Path.Combine(Application.StartupPath, "Icons", "previous.png"))
            btnPrev.Image = ResizeImage(icon, New Size(18, 18))
            icon.Dispose()
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error loading previous icon: " & ex.Message))
        End Try

        ' Setup Next button
        btnNext.Text = "Next"
        ApplySecondaryButtonStyle(btnNext, 120)
        Try
            Dim icon As Image = Image.FromFile(Path.Combine(Application.StartupPath, "Icons", "next.png"))
            btnNext.Image = ResizeImage(icon, New Size(18, 18))
            icon.Dispose()
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error loading next icon: " & ex.Message))
        End Try

        ' Setup Save Path TextBox
        txtSavePath.Size = New Size(220, 30)
        txtSavePath.Text = savePath ' Default to Downloads

        ' Setup Set Path button
        btnSetPath.Text = "Set Path"
        ApplySecondaryButtonStyle(btnSetPath, 120)
        Try
            Dim icon As Image = Image.FromFile(Path.Combine(Application.StartupPath, "Icons", "folder.png"))
            btnSetPath.Image = ResizeImage(icon, New Size(18, 18))
            icon.Dispose()
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error loading folder icon: " & ex.Message))
        End Try

        ' Setup ProgressBar
        progressBar.Size = New Size(220, 10)
        progressBar.Style = ProgressBarStyle.Marquee
        progressBar.Visible = False ' Hidden by default

        lblImageInfo.Text = "No scanned images"
        lblImageInfo.AutoSize = True
        lblImageInfo.ForeColor = Color.FromArgb(107, 114, 128)

        Dim controlsCard As New Panel()
        controlsCard.Dock = DockStyle.Fill
        controlsCard.BackColor = Color.White
        controlsCard.Padding = New Padding(12)
        controlsCard.BorderStyle = BorderStyle.FixedSingle

        Dim controlsLayout As New TableLayoutPanel()
        controlsLayout.Dock = DockStyle.Fill
        controlsLayout.ColumnCount = 1
        controlsLayout.RowCount = 3
        controlsLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        controlsLayout.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        Dim actionRow As New FlowLayoutPanel()
        actionRow.Dock = DockStyle.Top
        actionRow.AutoSize = True
        actionRow.WrapContents = True
        actionRow.Controls.Add(btnScan)
        actionRow.Controls.Add(btnSend)
        actionRow.Controls.Add(btnZoomIn)
        actionRow.Controls.Add(btnZoomOut)
        actionRow.Controls.Add(btnPrev)
        actionRow.Controls.Add(btnNext)
        actionRow.Controls.Add(lblImageInfo)

        Dim pathRow As New FlowLayoutPanel()
        pathRow.Dock = DockStyle.Top
        pathRow.AutoSize = True
        pathRow.WrapContents = True
        Dim lblPath As New Label()
        lblPath.Text = "Save Path:"
        lblPath.AutoSize = True
        lblPath.ForeColor = Color.FromArgb(75, 85, 99)
        lblPath.Margin = New Padding(0, 8, 8, 0)
        pathRow.Controls.Add(lblPath)
        pathRow.Controls.Add(txtSavePath)
        pathRow.Controls.Add(btnSetPath)
        pathRow.Controls.Add(progressBar)

        controlsLayout.Controls.Add(actionRow, 0, 0)
        controlsLayout.Controls.Add(pathRow, 0, 1)
        controlsCard.Controls.Add(controlsLayout)
        mainLayout.Controls.Add(controlsCard, 0, 2)
        UpdatePreviewSize()
    End Sub

    Private Sub ApplyPrimaryButtonStyle(button As Button, backColor As Color, width As Integer)
        button.BackColor = backColor
        button.ForeColor = Color.White
        button.FlatStyle = FlatStyle.Flat
        button.FlatAppearance.BorderSize = 0
        button.Width = width
        button.Height = 36
        button.Font = New System.Drawing.Font("Segoe UI Semibold", 9.5F, System.Drawing.FontStyle.Bold)
        button.TextAlign = ContentAlignment.MiddleCenter
        button.ImageAlign = ContentAlignment.MiddleLeft
        button.Padding = New Padding(8, 0, 8, 0)
        button.Margin = New Padding(0, 0, 12, 12)
    End Sub

    Private Sub ApplySecondaryButtonStyle(button As Button, width As Integer)
        button.BackColor = Color.FromArgb(229, 231, 235)
        button.ForeColor = Color.FromArgb(31, 41, 55)
        button.FlatStyle = FlatStyle.Flat
        button.FlatAppearance.BorderSize = 0
        button.Width = width
        button.Height = 34
        button.Font = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular)
        button.TextAlign = ContentAlignment.MiddleCenter
        button.ImageAlign = ContentAlignment.MiddleLeft
        button.Padding = New Padding(8, 0, 8, 0)
        button.Margin = New Padding(0, 0, 12, 12)
    End Sub

    Private Sub UpdateImageInfo()
        If scannedImages.Count = 0 OrElse currentImageIndex < 0 Then
            lblImageInfo.Text = "No scanned images"
        Else
            lblImageInfo.Text = $"Image {currentImageIndex + 1} of {scannedImages.Count}"
        End If
    End Sub

    Private Sub UpdatePreviewSize()
        If previewCard Is Nothing Then
            Return
        End If
        Dim width As Integer = Math.Max(1, CInt(previewCard.ClientSize.Width * zoomLevel))
        Dim height As Integer = Math.Max(1, CInt(previewCard.ClientSize.Height * zoomLevel))
        pictureBox.Size = New Size(width, height)
    End Sub

    Private Sub ShowLaunchParameters()
        Try
            Dim args = Environment.GetCommandLineArgs()
            Dim urlArg As String = args.FirstOrDefault(Function(a) a.StartsWith("customscan://", StringComparison.OrdinalIgnoreCase))
            If String.IsNullOrEmpty(urlArg) Then
                Return
            End If

            Dim uri As New Uri(urlArg)
            Dim parameters = ParseQueryString(uri.Query)
            If parameters.Count = 0 Then
                Return
            End If

            loaNo = GetParam(parameters, "loa_no")
            transactionNo = GetParam(parameters, "transaction_no")
            batchNo = GetParam(parameters, "batch_no")
            userToken = GetParam(parameters, "token")

            Dim message As String =
                "loa_no: " & loaNo & Environment.NewLine &
                "transaction_no: " & transactionNo & Environment.NewLine &
                "batch_no: " & batchNo & Environment.NewLine &
                "token: " & userToken
            MessageBox.Show(message, "Custom Scan App Parameters")
        Catch ex As Exception
            MessageBox.Show("Failed to read launch parameters: " & ex.Message)
        End Try
    End Sub

    Private Function ParseQueryString(query As String) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        If String.IsNullOrWhiteSpace(query) Then
            Return result
        End If
        Dim trimmed = query.TrimStart("?"c)
        For Each pair In trimmed.Split("&"c)
            If String.IsNullOrWhiteSpace(pair) Then
                Continue For
            End If
            Dim parts = pair.Split("="c, 2)
            Dim key = Uri.UnescapeDataString(parts(0))
            Dim value As String = ""
            If parts.Length > 1 Then
                value = Uri.UnescapeDataString(parts(1))
            End If
            If Not result.ContainsKey(key) Then
                result(key) = value
            End If
        Next
        Return result
    End Function

    Private Function GetParam(parameters As Dictionary(Of String, String), key As String) As String
        Dim value As String = ""
        If parameters.TryGetValue(key, value) Then
            Return value
        End If
        Return ""
    End Function

    Private Sub EnsureCustomProtocolRegistered()
        Try
            Dim exePath As String = Application.ExecutablePath
            Dim commandValue As String = $"""{exePath}"" ""%1"""
            Using customKey As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\Classes\customscan")
                If customKey.GetValue("URL Protocol") Is Nothing Then
                    customKey.SetValue("", "URL:Custom Scan Protocol")
                    customKey.SetValue("URL Protocol", "")
                End If
            End Using
            Using commandKey As RegistryKey = Registry.CurrentUser.CreateSubKey("Software\Classes\customscan\shell\open\command")
                Dim existing As Object = commandKey.GetValue("")
                If existing Is Nothing OrElse Not String.Equals(existing.ToString(), commandValue, StringComparison.OrdinalIgnoreCase) Then
                    commandKey.SetValue("", commandValue)
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Failed to register customscan protocol: " & ex.Message)
        End Try
    End Sub

    Private Sub btnSetPath_Click(sender As Object, e As EventArgs) Handles btnSetPath.Click
        Try
            Dim newPath As String = txtSavePath.Text.Trim()
            If String.IsNullOrEmpty(newPath) Then
                savePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                MessageBox.Show("Save path reset to Downloads folder.")
            ElseIf Directory.Exists(newPath) Then
                savePath = newPath
                MessageBox.Show("Save path set to: " & savePath)
            Else
                MessageBox.Show("Invalid path. Please enter a valid directory.")
            End If
            txtSavePath.Text = savePath
        Catch ex As Exception
            MessageBox.Show("Error setting save path: " & ex.Message)
        End Try
    End Sub

    Private Sub btnScan_Click(sender As Object, e As EventArgs) Handles btnScan.Click
        Try
            ' Show loading indicator
            Me.Invoke(Sub()
                          progressBar.Visible = True
                          btnScan.Enabled = False ' Disable scan button during scanning
                      End Sub)

            ' Initialize TWAIN session
            Dim appId = TWIdentity.CreateFromAssembly(DataGroups.Image, GetType(Form1).Assembly)
            Dim twainApp As New TwainSession(appId)
            twainApp.SynchronizationContext = Nothing
            twainApp.Open()

            ' List available scanners for debugging
            Dim sources = twainApp.GetSources().ToList()
            If Not sources.Any() Then
                Me.Invoke(Sub() MessageBox.Show("No TWAIN-compatible scanners detected. Please ensure your scanner is connected and drivers are installed."))
                twainApp.Close()
                Me.Invoke(Sub()
                              progressBar.Visible = False
                              btnScan.Enabled = True
                          End Sub)
                Return
            End If
            Me.Invoke(Sub() MessageBox.Show("Detected scanners: " & String.Join(", ", sources.Select(Function(s) s.Name))))

            ' Use ShowSourceSelector as per original logic
            Dim dataSource = twainApp.ShowSourceSelector()
            If dataSource Is Nothing Then
                Me.Invoke(Sub() MessageBox.Show("Walang napiling scanner o kinansela ang pagpili."))
                twainApp.Close()
                Me.Invoke(Sub()
                              progressBar.Visible = False
                              btnScan.Enabled = True
                          End Sub)
                Return
            End If

            ' Open the selected data source
            dataSource.Open()

            ' Set scanning parameters
            If dataSource.Capabilities.CapXferCount.CanSet Then
                dataSource.Capabilities.CapXferCount.SetValue(1)
            End If
            If dataSource.Capabilities.ICapPixelType.CanSet AndAlso
               dataSource.Capabilities.ICapPixelType.GetValues().Contains(PixelType.RGB) Then
                dataSource.Capabilities.ICapPixelType.SetValue(PixelType.RGB)
            End If
            If dataSource.Capabilities.ICapUnits.CanSet AndAlso
               dataSource.Capabilities.ICapUnits.GetValues().Contains(Unit.Inches) Then
                dataSource.Capabilities.ICapUnits.SetValue(Unit.Inches)
            End If
            If dataSource.Capabilities.ICapXResolution.CanSet Then
                dataSource.Capabilities.ICapXResolution.SetValue(New TWFix32 With {.Whole = 300, .Fraction = 0})
            End If
            If dataSource.Capabilities.ICapYResolution.CanSet Then
                dataSource.Capabilities.ICapYResolution.SetValue(New TWFix32 With {.Whole = 300, .Fraction = 0})
            End If

            ' Handle DataTransferred event
            Dim scanCompleted As Boolean = False
            AddHandler twainApp.DataTransferred, Sub(s, ea)
                                                     Using stream As Stream = ea.GetNativeImageStream()
                                                         If stream IsNot Nothing AndAlso stream.Length > 0 Then
                                                             Try
                                                                 ' Create a new Bitmap to avoid GDI+ issues
                                                                 Dim tempImage As New Bitmap(stream)
                                                                 Dim newImage As New Bitmap(tempImage) ' Clone to prevent disposal issues
                                                                 scannedImages.Add(newImage) ' Add to list
                                                                 tempImage.Dispose()
                                                                 currentImageIndex = scannedImages.Count - 1 ' Show latest image
                                                                 If newImage IsNot Nothing Then
                                                                     Me.Invoke(Sub()
                                                                                   pictureBox.Image = scannedImages(currentImageIndex) ' Display in PictureBox
                                                                                   UpdateImageInfo()
                                                                                   MessageBox.Show("Image successfully captured and displayed.")
                                                                               End Sub)
                                                                 Else
                                                                     Me.Invoke(Sub() MessageBox.Show("Failed to create image from stream."))
                                                                 End If
                                                             Catch ex As Exception
                                                                 Me.Invoke(Sub() MessageBox.Show("Error processing image stream: " & ex.Message))
                                                             End Try
                                                         Else
                                                             Me.Invoke(Sub() MessageBox.Show("Walang na-scan na larawan o walang laman ang stream."))
                                                         End If
                                                     End Using
                                                 End Sub

            ' Handle SourceDisabled event
            AddHandler twainApp.SourceDisabled, Sub(s, ea)
                                                    scanCompleted = True
                                                    Me.Invoke(Sub()
                                                                  progressBar.Visible = False ' Hide loading indicator
                                                                  btnScan.Enabled = True ' Re-enable scan button
                                                              End Sub)
                                                End Sub

            ' Enable scanner
            dataSource.Enable(SourceEnableMode.ShowUI, True, Me.Handle)

            ' Wait for scan to complete
            While Not scanCompleted
                Thread.Sleep(100)
                Application.DoEvents()
            End While

            dataSource.Close()
            twainApp.Close()
        Catch ex As Exception
            Me.Invoke(Sub()
                          MessageBox.Show("Error sa pag-scan: " & ex.Message)
                          progressBar.Visible = False ' Hide loading indicator on error
                          btnScan.Enabled = True
                      End Sub)
        End Try
    End Sub

    Private Async Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
        If scannedImages.Count > 0 Then
            Try
                ' Show loading indicator
                Me.Invoke(Sub()
                              progressBar.Visible = True
                              btnSend.Enabled = False ' Disable send button during processing
                          End Sub)

                ' Save each image as a separate PDF
                Dim savedFiles As New List(Of String)()
                Dim filesPayload As New List(Of Dictionary(Of String, String))()
                For i As Integer = 0 To scannedImages.Count - 1
                    Dim currentIndex = i + 1
                    Dim img = scannedImages(i)
                    If img.Width > 0 AndAlso img.Height > 0 Then
                        Dim filePath As String = Path.Combine(savePath, $"scan_pdf{currentIndex}.pdf")
                        Using document As New Document(PageSize.A4, 10, 10, 10, 10)
                            Using stream As New FileStream(filePath, FileMode.Create)
                                PdfWriter.GetInstance(document, stream)
                                document.Open()
                                Dim pdfImg As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(img, System.Drawing.Imaging.ImageFormat.Png)
                                pdfImg.ScaleToFit(document.PageSize.Width - 20, document.PageSize.Height - 20)
                                document.Add(pdfImg)
                                document.Close()
                            End Using
                        End Using
                        ' Verify PDF file exists and has content
                        If File.Exists(filePath) AndAlso New FileInfo(filePath).Length > 0 Then
                            savedFiles.Add(filePath)
                            Dim pdfBytes As Byte() = File.ReadAllBytes(filePath)
                            Dim base64Pdf As String = Convert.ToBase64String(pdfBytes)
                            filesPayload.Add(New Dictionary(Of String, String) From {
                                {"file_name", Path.GetFileName(filePath)},
                                {"content_type", "application/pdf"},
                                {"data_base64", base64Pdf}
                            })
                        Else
                            Dim missingFilePath = filePath
                            Me.Invoke(Sub() MessageBox.Show($"PDF {missingFilePath} is empty or was not created correctly."))
                        End If
                    Else
                        Dim badIndex = currentIndex
                        Me.Invoke(Sub() MessageBox.Show($"Skipping invalid image {badIndex} with zero dimensions."))
                    End If
                Next

                If filesPayload.Count > 0 Then
                    Dim payload As New Dictionary(Of String, Object) From {
                        {"type", "scan_upload"},
                        {"loa_no", loaNo},
                        {"transaction_no", transactionNo},
                        {"batch_no", batchNo},
                        {"token", userToken},
                        {"files", filesPayload}
                    }
                    Dim jsonPayload As String = System.Text.Json.JsonSerializer.Serialize(payload)
                    Await SendToWebSocket(jsonPayload)
                End If

                ' Hide loading indicator and show result
                Me.Invoke(Sub()
                              progressBar.Visible = False
                              btnSend.Enabled = True
                              If savedFiles.Count > 0 Then
                                  MessageBox.Show("PDFs successfully saved at: " & String.Join(", ", savedFiles))
                              Else
                                  MessageBox.Show("No valid PDFs were saved.")
                              End If
                          End Sub)
            Catch ex As Exception
                Me.Invoke(Sub()
                              MessageBox.Show("Error sa pag-save o pag-send: " & ex.Message)
                              progressBar.Visible = False
                              btnSend.Enabled = True
                          End Sub)
            End Try
        Else
            Me.Invoke(Sub()
                          MessageBox.Show("Walang larawan na na-scan para i-save.")
                          progressBar.Visible = False
                          btnSend.Enabled = True
                      End Sub)
        End If
    End Sub

    Private Sub btnZoomIn_Click(sender As Object, e As EventArgs) Handles btnZoomIn.Click
        If pictureBox.Image IsNot Nothing Then
            zoomLevel = Math.Min(2.5F, zoomLevel + 0.1F)
            pictureBox.SizeMode = PictureBoxSizeMode.Zoom
            UpdatePreviewSize()
        End If
    End Sub

    Private Sub btnZoomOut_Click(sender As Object, e As EventArgs) Handles btnZoomOut.Click
        If pictureBox.Image IsNot Nothing AndAlso zoomLevel > 0.2F Then
            zoomLevel = Math.Max(0.4F, zoomLevel - 0.1F)
            pictureBox.SizeMode = PictureBoxSizeMode.Zoom
            UpdatePreviewSize()
        End If
    End Sub

    Private Sub btnPrev_Click(sender As Object, e As EventArgs) Handles btnPrev.Click
        If scannedImages.Count > 0 AndAlso currentImageIndex > 0 Then
            currentImageIndex -= 1
            pictureBox.Image = scannedImages(currentImageIndex)
            zoomLevel = 1.0F
            pictureBox.SizeMode = PictureBoxSizeMode.Zoom
            UpdatePreviewSize()
            UpdateImageInfo()
        End If
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        If scannedImages.Count > 0 AndAlso currentImageIndex < scannedImages.Count - 1 Then
            currentImageIndex += 1
            pictureBox.Image = scannedImages(currentImageIndex)
            zoomLevel = 1.0F
            pictureBox.SizeMode = PictureBoxSizeMode.Zoom
            UpdatePreviewSize()
            UpdateImageInfo()
        End If
    End Sub
End Class
