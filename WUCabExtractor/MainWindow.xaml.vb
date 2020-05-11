Imports System.IO
Imports System.Windows.Forms
Imports System.Windows.Window
Imports System.Xml
Class MainWindow
    Dim InputDirectory As String
    Dim OutputDiectory As String
    Dim ResourceDirectory As String
    Dim IsKeepStructure As Boolean
    Dim EmptyList As New List(Of String)
    Dim MessageList As New List(Of String)
    Sub RefreshMessageList()
        lstMessage.ItemsSource = EmptyList
        lstMessage.ItemsSource = MessageList
        DoEvents()
    End Sub
    Sub AddMessage(MessageText As String)
        MessageList.Add(MessageText)
        RefreshMessageList()
        lstMessage.SelectedIndex = lstMessage.Items.Count - 1
        lstMessage.ScrollIntoView(lstMessage.SelectedItem)
    End Sub
    Sub LockUI()
        txtInputDir.IsEnabled = False
        txtResourceDir.IsEnabled = False
        txtOutputDir.IsEnabled = False
        btnBrowseInput.IsEnabled = False
        btnBrowseResource.IsEnabled = False
        btnBrowseOutput.IsEnabled = False
        btnStart.IsEnabled = False
        chkKeepStructure.IsEnabled = False
    End Sub
    Sub UnlockUI()
        txtInputDir.IsEnabled = True
        txtResourceDir.IsEnabled = True
        txtOutputDir.IsEnabled = True
        btnBrowseInput.IsEnabled = True
        btnBrowseResource.IsEnabled = True
        btnBrowseOutput.IsEnabled = True
        btnStart.IsEnabled = True
        chkKeepStructure.IsEnabled = True
    End Sub
    Private Sub SetTaskbarProgess(MaxValue As Integer, MinValue As Integer, CurrentValue As Integer, Optional State As Shell.TaskbarItemProgressState = Shell.TaskbarItemProgressState.Normal)
        If MaxValue <= MinValue Or CurrentValue < MinValue Or CurrentValue > MaxValue Then
            Exit Sub
        End If
        TaskbarItem.ProgressValue = (CurrentValue - MinValue) / (MaxValue - MinValue)
        TaskbarItem.ProgressState = State
    End Sub
    Function GetPathFromFile(FilePath As String) As String
        If FilePath.Trim = "" Then
            Return ""
        End If
        If FilePath(FilePath.Length - 1) = "\" Then
            Return FilePath
        End If
        Try
            Return FilePath.Substring(0, FilePath.LastIndexOf("\"))
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Function GetNameFromFullPath(FullPath As String) As String
        If FullPath.Trim = "" Then
            Return ""
        End If
        If FullPath(FullPath.Length - 1) = "\" Then
            Return ""
        End If
        Try
            Return FullPath.Substring(FullPath.LastIndexOf("\") + 1, FullPath.LastIndexOf(".") - FullPath.LastIndexOf("\") - 1)
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Private Sub btnBrowseInput_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowseInput.Click
        Dim FolderBrowser As New FolderBrowserDialog
        With FolderBrowser
            .Description = "请指定 MUM 文件的位置，然后单击""确定""按钮。"
        End With
        If FolderBrowser.ShowDialog() = Forms.DialogResult.OK Then
            InputDirectory = FolderBrowser.SelectedPath
            If InputDirectory(InputDirectory.Length - 1) <> "\" Then
                InputDirectory = InputDirectory & "\"
            End If
            txtInputDir.Text = InputDirectory
        End If
    End Sub
    Private Sub btnBrowseResource_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowseResource.Click
        Dim FolderBrowser As New FolderBrowserDialog
        With FolderBrowser
            .Description = "请指定用于抽取文件的目录的位置，然后单击""确定""按钮。"
        End With
        If FolderBrowser.ShowDialog() = Forms.DialogResult.OK Then
            ResourceDirectory = FolderBrowser.SelectedPath
            If ResourceDirectory(ResourceDirectory.Length - 1) <> "\" Then
                ResourceDirectory = ResourceDirectory & "\"
            End If
            txtResourceDir.Text = ResourceDirectory
        End If
    End Sub

    Private Sub btnBrowseOutput_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowseOutput.Click
        Dim FolderBrowser As New FolderBrowserDialog
        With FolderBrowser
            .Description = "请指定重建完成的目录结构要输出的位置，然后单击""确定""按钮。"
        End With
        If FolderBrowser.ShowDialog() = Forms.DialogResult.OK Then
            OutputDiectory = FolderBrowser.SelectedPath
            If OutputDiectory(OutputDiectory.Length - 1) <> "\" Then
                OutputDiectory = OutputDiectory & "\"
            End If
            txtOutputDir.Text = OutputDiectory
        End If
    End Sub

    Private Sub btnStart_Click(sender As Object, e As RoutedEventArgs) Handles btnStart.Click
        LockUI()
        If txtInputDir.Text.Trim = "" Then
            MessageBox.Show("CAB 输入路径不能为空。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UnlockUI()
            Exit Sub
        End If
        If txtResourceDir.Text.Trim = "" Then
            MessageBox.Show("文件抽取源路径不能为空。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UnlockUI()
            Exit Sub
        End If
        If txtOutputDir.Text.Trim = "" Then
            MessageBox.Show("输出路径不能为空。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UnlockUI()
            Exit Sub
        End If
        If Not Directory.Exists(OutputDiectory) Then
            Try
                Directory.CreateDirectory(OutputDiectory)
            Catch ex As Exception
                MessageBox.Show("试图创建输出目录""" & OutputDiectory & """时发生错误: " & vbCrLf & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UnlockUI()
                Exit Sub
            End Try
        End If
        With prgProgress
            .Minimum = 0
            .Maximum = 100
            .Value = 0
        End With
        MessageList.Clear()
        RefreshMessageList()

        AddMessage("正在确定 CAB 文件总数。")
        Dim nMUMFileCount As Integer = Directory.GetFiles(InputDirectory, "*.mum", SearchOption.TopDirectoryOnly).Length
        If nMUMFileCount = 0 Then
            MessageBox.Show("输入目录""" & InputDirectory & """中不包含任何 MUM 文件。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            AddMessage("输入目录""" & InputDirectory & """中不包含任何 MUM 文件。")
            AddMessage("发生错误，取消操作。")
            UnlockUI()
            Exit Sub
        End If
        AddMessage("计算完毕，共有 " & nMUMFileCount.ToString & " 个 MUM 文件。")
        With prgProgress
            .Minimum = 0
            .Maximum = nMUMFileCount
            .Value = 0
        End With
        SetTaskbarProgess(prgProgress.Maximum, 0, prgProgress.Value)
        Dim nSuccess As UInteger = 0
        Dim nFail As UInteger = 0
        Dim nIgnored As UInteger = 0
        Dim IsErrorOccurred As Boolean = False

        For Each MUMFilePath In Directory.EnumerateFiles(InputDirectory, "*.mum", SearchOption.TopDirectoryOnly)
            Dim MUMFileName As String = GetNameFromFullPath(MUMFilePath)
            Dim UpdateInfoFile As New XmlDocument
            AddMessage("正在打开包描述文件""" & MUMFilePath & """。")
            RefreshMessageList()
            Try
                UpdateInfoFile.Load(MUMFilePath)
            Catch ex As Exception
                AddMessage("无法打开包描述文件""" & MUMFilePath & """，发生错误: " & ex.Message)
                nFail += 1
                prgProgress.Value += 1
                SetTaskbarProgess(prgProgress.Maximum, 0, prgProgress.Value)
                Continue For
            End Try
            AddMessage("成功打开包描述文件""" & MUMFilePath & """。")
            Dim nsMgr As New XmlNamespaceManager(UpdateInfoFile.NameTable)
            nsMgr.AddNamespace("ns", "urn:schemas-microsoft-com:asm.v3")
            Dim CustomInformationNode As XmlNode = UpdateInfoFile.SelectSingleNode("/ns:assembly/ns:package/ns:customInformation", nsMgr)
            AddMessage("正在定位 XML 节点""/assembly/package/customInformation""。")
            If IsNothing(CustomInformationNode) Then
                AddMessage("XML 节点定位失败。")
                nFail += 1
                prgProgress.Value += 1
                SetTaskbarProgess(prgProgress.Maximum, 0, prgProgress.Value)
                Continue For
            End If
            AddMessage("XML 节点""assembly/package/customInformation""定位成功，共有 " & CustomInformationNode.ChildNodes.Count & " 条记录。")
            Dim TempFileInfo As New WindowsUpdatePackageFileNodeProperties
            Dim FileList As XmlNodeList = CustomInformationNode.ChildNodes
            For Each FileNode As XmlNode In FileList
                Dim FileElement As XmlElement = FileNode
                If FileElement.Name <> "file" Then
                    AddMessage("已忽略一个节点，因为它的类型是""" & FileElement.Name & """而不是""file""。")
                    Continue For
                End If
                Try
                    With TempFileInfo
                        .Name = FileElement.GetAttribute("name").ToString
                        .CabPath = FileElement.GetAttribute("cabpath").ToString
                    End With
                    If Not TempFileInfo.Name.StartsWith("$(runtime.system32)\") And Not TempFileInfo.Name.StartsWith("$(runtime.bootdrive)\") And Not TempFileInfo.Name.StartsWith("$(runtime.drivers)\") Then
                        AddMessage("已忽略一个文件节点，因为它没有描述文件复制信息。")
                        Continue For
                    End If
                    With TempFileInfo
                        If .Name.StartsWith("$(runtime.system32)\") Then
                            .Name = .Name.Replace("$(runtime.system32)\", ResourceDirectory & "Windows\System32\")
                        End If
                        If .Name.StartsWith("$(runtime.bootdrive)\") Then
                            .Name = .Name.Replace("$(runtime.bootdrive)\", ResourceDirectory)
                        End If
                        If .Name.StartsWith("$(runtime.drivers)\") Then
                            .Name = .Name.Replace("$(runtime.drivers)\", ResourceDirectory & "Windows\System32\Drivers\")
                        End If
                        .CabPath = OutputDiectory & MUMFileName & "\" & .CabPath
                    End With
                    Dim CopyDest As String
                    If IsKeepStructure Then
                        CopyDest = FileElement.GetAttribute("name").ToString
                        If CopyDest.StartsWith("$(runtime.system32)\") Then
                            CopyDest = CopyDest.Replace("$(runtime.system32)\", OutputDiectory & "Windows\System32\")
                        End If
                        If CopyDest.StartsWith("$(runtime.bootdrive)\") Then
                            CopyDest = CopyDest.Replace("$(runtime.bootdrive)\", OutputDiectory)
                        End If
                        If CopyDest.StartsWith("$(runtime.drivers)\") Then
                            CopyDest = CopyDest.Replace("$(runtime.drivers)\", OutputDiectory & "Windows\System32\Drivers\")
                        End If
                    Else
                        CopyDest = TempFileInfo.CabPath
                    End If
                    Dim CopyDestDir As String = GetPathFromFile(CopyDest)
                    If Not Directory.Exists(CopyDestDir) Then
                        Directory.CreateDirectory(CopyDestDir)
                    End If
                    If File.Exists(CopyDest) Then
                        File.Delete(CopyDest)
                    End If
                    File.Copy(TempFileInfo.Name, CopyDest)
                    AddMessage("已成功从""" & TempFileInfo.Name & """复制文件到""" & CopyDest & """。")
                    DoEvents()
                Catch ex As Exception
                    AddMessage("已忽略一个文件节点，因为发生错误: " & ex.Message)
                    Continue For
                End Try
            Next
            AddMessage("对包描述文件""" & MUMFilePath & """的操作成功完成。")
            nSuccess += 1
            prgProgress.Value += 1
            SetTaskbarProgess(prgProgress.Maximum, 0, prgProgress.Value)
        Next

        MessageBox.Show("操作完成，共有 " & nSuccess.ToString & "个 MUM 文件被处理，有 " & nIgnored.ToString & " 个 MUM 文件被忽略，处理 " & nFail.ToString & " 个 MUM 文件时出错。", "大功告成!", MessageBoxButtons.OK, MessageBoxIcon.Information)
        UnlockUI()
        With prgProgress
            .Minimum = 0
            .Maximum = 100
            .Value = 0
        End With
        SetTaskbarProgess(100, 0, 0)
    End Sub

    Private Sub chkKeepStructure_Click(sender As Object, e As RoutedEventArgs) Handles chkKeepStructure.Click
        IsKeepStructure = chkKeepStructure.IsChecked
    End Sub
End Class
