Imports mshtml
Imports System.Text
Imports System.IO
Imports System.Net
Class MainWindow
    Dim EmptyList As New List(Of String)
    Dim LogList As New List(Of String)
    Dim SavePath As String
    Dim SiteUrl As String
    Dim PageListPageName As String
    Dim PageListPageUrl As String
    Dim WebBrowserLoadingOperationWaiter As New EventOrOperationWaiter
    Dim GeneralTimoutWaiter As New EventOrOperationWaiter
    Private Sub RefreshLogList()
        lstLog.ItemsSource = EmptyList
        lstLog.ItemsSource = LogList
        lstLog.SelectedIndex = lstLog.Items.Count - 1
        Try
            lstLog.ScrollIntoView(lstLog.SelectedItem)
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub
    Private Sub WriteLog(LogMessage As String)
        LogList.Add(LogMessage)
        RefreshLogList()
        DoEvents()
    End Sub
    Private Sub ClearLog()
        LogList.Clear()
        RefreshLogList()
    End Sub
    Private Sub LockUi()
        txtSiteUrl.IsEnabled = False
        txtListPageName.IsEnabled = False
        txtSavePath.IsEnabled = False
        btnBrowseSavePath.IsEnabled = False
        btnStartBackup.IsEnabled = False
        btnLogin.IsEnabled = False
    End Sub
    Private Sub UnlockUi()
        txtSiteUrl.IsEnabled = True
        txtListPageName.IsEnabled = True
        txtSavePath.IsEnabled = True
        btnBrowseSavePath.IsEnabled = True
        btnStartBackup.IsEnabled = True
        btnLogin.IsEnabled = True
    End Sub

    Private Sub btnBrowseSavePath_Click(sender As Object, e As RoutedEventArgs) Handles btnBrowseSavePath.Click
        Dim FolderBrowser As New System.Windows.Forms.FolderBrowserDialog
        With FolderBrowser
            .Description = "Please select the path to save the site backup."
        End With
        If FolderBrowser.ShowDialog() = Forms.DialogResult.OK Then
            SavePath = FolderBrowser.SelectedPath
            If SavePath(SavePath.Length - 1) <> "\" Then
                SavePath = SavePath & "\"
            End If
            txtSavePath.Text = SavePath
        End If
    End Sub

    Private Sub txtListPageName_KeyUp(sender As Object, e As Input.KeyEventArgs) Handles txtListPageName.KeyUp
        If e.Key = Key.Enter Then
            SiteUrl = txtSiteUrl.Text.Trim()
            PageListPageName = txtListPageName.Text.Trim()
            PageListPageUrl = SiteUrl & "/" & PageListPageName.Trim()
            LockUi()
            WebBrowserLoadingOperationWaiter.SetWaitingConditionNotSatisfied()
            wbbWikidotSiteContainer.Navigate(PageListPageUrl)
            WebBrowserLoadingOperationWaiter.WaitForEventOrOperationFinished(240 * 1000)
            UnlockUi()
        End If
    End Sub

    Private Sub txtSiteUrl_KeyUp(sender As Object, e As Input.KeyEventArgs) Handles txtSiteUrl.KeyUp
        If e.Key = Key.Enter Then
            SiteUrl = txtSiteUrl.Text.Trim()
            PageListPageName = txtListPageName.Text.Trim()
            PageListPageUrl = SiteUrl & "/" & PageListPageName.Trim()
            LockUi()
            WebBrowserLoadingOperationWaiter.SetWaitingConditionNotSatisfied()
            wbbWikidotSiteContainer.Navigate(PageListPageUrl)
            WebBrowserLoadingOperationWaiter.WaitForEventOrOperationFinished(240 * 1000)
            UnlockUi()
        End If
    End Sub

    Private Sub MainWindow_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        WebBrowserLoadingOperationWaiter.SetWaitingConditionAborted()
        GeneralTimoutWaiter.SetWaitingConditionAborted()
        End
    End Sub

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Dim ActiveX = wbbWikidotSiteContainer.GetType().InvokeMember("ActiveXInstance", Reflection.BindingFlags.GetProperty Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic, Nothing, wbbWikidotSiteContainer, Nothing)
        ActiveX.Silent = True
        SiteUrl = txtSiteUrl.Text.Trim()
        PageListPageName = txtListPageName.Text.Trim()
        PageListPageUrl = SiteUrl & "/" & PageListPageName.Trim()
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As RoutedEventArgs) Handles btnLogin.Click
        LockUi()
        WebBrowserLoadingOperationWaiter.SetWaitingConditionNotSatisfied()
        wbbWikidotSiteContainer.Navigate("https://www.wikidot.com/default--flow/login__LoginPopupScreen")
        WebBrowserLoadingOperationWaiter.WaitForEventOrOperationFinished(240 * 1000)
        UnlockUi()
    End Sub

    Private Sub wbbWikidotSiteContainer_LoadCompleted(sender As Object, e As NavigationEventArgs) Handles wbbWikidotSiteContainer.LoadCompleted
        WebBrowserLoadingOperationWaiter.SetWaitingConditionSatisfied()
    End Sub

    Private Sub btnStartBackup_Click(sender As Object, e As RoutedEventArgs) Handles btnStartBackup.Click
        'Lock UI
        ClearLog()
        LockUi()

        'Save variables
        SavePath = txtSavePath.Text.Trim()
        SiteUrl = txtSiteUrl.Text.Trim()
        PageListPageName = txtListPageName.Text.Trim()
        PageListPageUrl = SiteUrl & "/" & PageListPageName.Trim()

        'Create save directory
        If Not Directory.Exists(SavePath) Then
            Directory.CreateDirectory(SavePath)
            WriteLog("Created output dir: " + SavePath)
        End If

        'Load page list
        Dim RetryCounter As Integer
        For RetryCounter = 1 To 5
            WebBrowserLoadingOperationWaiter.SetWaitingConditionNotSatisfied()
            If RetryCounter = 1 Then
                wbbWikidotSiteContainer.Navigate(PageListPageUrl)
            Else
                wbbWikidotSiteContainer.Refresh()
            End If
            If WebBrowserLoadingOperationWaiter.WaitForEventOrOperationFinished(240 * 1000) Then
                RetryCounter = 0
                Exit For
            End If
            DoEvents()
        Next
        If RetryCounter <> 0 Then
            WriteLog("WARNING: Failed to load page list (timed out) after 5 tries, operations may fail.")
        End If

        'Create HTML document object and elements collection
        Dim CurrentHTMLDocument As HTMLDocument = Nothing
        Dim ChildrenHTMLDocument As HTMLDocument = Nothing
        Dim PossibleElements As IHTMLElementCollection = Nothing
        Dim ChildrenCollection As IHTMLElementCollection = Nothing
        Dim CurrentElement As IHTMLElement = Nothing
        Dim ChildElement As IHTMLElement = Nothing

        'Get page count of the page list page
        CurrentHTMLDocument = wbbWikidotSiteContainer.Document
        Dim PageListPageCount As Integer = 1
        PossibleElements = CurrentHTMLDocument.getElementsByTagName("div")
        Dim IsPagerFound As Boolean = False
        For Each CurrentElement In PossibleElements
            If CurrentElement.className = "pager" Then
                IsPagerFound = True
                Exit For
            End If
        Next
        If IsPagerFound Then
            PageListPageCount = 1
            ChildrenCollection = CurrentElement.children
            For Each CurrentElement In ChildrenCollection
                If CurrentElement.className = "target" And IsNumeric(CurrentElement.innerText) Then
                    PageListPageCount += 1
                End If
            Next
        Else
            PageListPageCount = 1
        End If
        WriteLog("The page list page has " & PageListPageCount & " sub page(s).")

        'Generate page list
        Dim PageList As New List(Of WikidotPage)
        For i As Integer = 1 To PageListPageCount
            WriteLog("Generating page list on page " & i.ToString() & " of the page list page.")
            'Load page list page
            For RetryCounter = 1 To 5
                WebBrowserLoadingOperationWaiter.SetWaitingConditionNotSatisfied()
                If RetryCounter = 1 Then
                    wbbWikidotSiteContainer.Navigate(PageListPageUrl & "/p/" & i.ToString())
                Else
                    wbbWikidotSiteContainer.Refresh()
                End If
                If WebBrowserLoadingOperationWaiter.WaitForEventOrOperationFinished(240 * 1000) Then
                    RetryCounter = 0
                    Exit For
                End If
                DoEvents()
            Next
            If RetryCounter <> 0 Then
                WriteLog("WARNING: Failed to load page list (timed out) after 5 tries, operations may fail.")
            End If
            'Enumerate links to pages
            CurrentHTMLDocument = wbbWikidotSiteContainer.Document
            Dim PageContentDiv As IHTMLElement = CurrentHTMLDocument.getElementById("page-content")
            If Not IsNothing(PageContentDiv) Then
                PossibleElements = PageContentDiv.all
                Dim IsListPagesBoxFound As Boolean = False
                For Each CurrentElement In PossibleElements
                    If CurrentElement.className = "list-pages-box" Then
                        IsListPagesBoxFound = True
                        Exit For
                    End If
                Next
                If IsListPagesBoxFound Then
                    ChildrenCollection = CurrentElement.all
                    For Each CurrentElement In ChildrenCollection
                        'Find a link
                        If CurrentElement.tagName.ToUpper = "A" And CurrentElement.parentElement.tagName.ToUpper = "P" Then
                            Dim LinkElement As IHTMLAnchorElement = CurrentElement
                            PageList.Add(New WikidotPage(LinkElement.href))
                        End If
                    Next
                End If
            End If
        Next
        WriteLog(PageList.Count.ToString & " page(s) found.")

        'Save page sources and attached files
        If PageList.Count = 0 Then
            MessageBox.Show("No pages to backup, exitting...", "No page", MessageBoxButton.OK, MessageBoxImage.Information)
            WriteLog("No pages to backup, exitting...")
            UnlockUi()
            Exit Sub
        End If
        prgProgress.Maximum = PageList.Count
        prgProgress.Minimum = 0
        Dim SucceededCount As Integer = 0
        Dim FailedCount As Integer = 0
        Dim IgnoredCount As Integer = 0
        If Not Directory.Exists(SavePath & "\source") Then
            Directory.CreateDirectory(SavePath & "\source")
        End If
        For i As Integer = 0 To PageList.Count - 1
            WriteLog("Backing up page '" & PageList(i).PageFullName & "', page " & i + 1 & " of " & PageList.Count)
            'Navigate to page
            For RetryCounter = 1 To 5
                WebBrowserLoadingOperationWaiter.SetWaitingConditionNotSatisfied()
                If RetryCounter = 1 Then
                    wbbWikidotSiteContainer.Navigate(PageList(i).PageUrl & "/norender/true")
                Else
                    wbbWikidotSiteContainer.Refresh()
                End If
                If WebBrowserLoadingOperationWaiter.WaitForEventOrOperationFinished(240 * 1000) Then
                    'Check if loaded page is valid
                    CurrentHTMLDocument = wbbWikidotSiteContainer.Document
                    If IsNothing(CurrentHTMLDocument.getElementById("more-options-button")) Then
                        Continue For
                    End If
                    RetryCounter = 0
                    Exit For
                End If
                DoEvents()
            Next
            If RetryCounter <> 0 Then
                WriteLog("Failed to load page '" & PageList(i).PageFullName & "' after 5 tries, page skipped.")
                FailedCount += 1
                Continue For
            End If
            'Save page source
            CurrentHTMLDocument = wbbWikidotSiteContainer.Document
            Dim MoreOptionsButton As IHTMLElement = Nothing
            For RetryCounter = 1 To 5
                MoreOptionsButton = CurrentHTMLDocument.getElementById("more-options-button")
                If Not IsNothing(MoreOptionsButton) Then
                    Exit For
                End If
                GeneralTimoutWaiter.SetWaitingConditionNotSatisfied()
                GeneralTimoutWaiter.WaitForEventOrOperationFinished(100)
                DoEvents()
            Next
            Try
                MoreOptionsButton.click()
                GeneralTimoutWaiter.SetWaitingConditionNotSatisfied()
                GeneralTimoutWaiter.WaitForEventOrOperationFinished(1000)
                Dim ViewSourceButton As IHTMLElement = Nothing
                For RetryCounter = 1 To 5
                    CurrentHTMLDocument = wbbWikidotSiteContainer.Document
                    ViewSourceButton = CurrentHTMLDocument.getElementById("view-source-button")
                    If Not IsNothing(ViewSourceButton) Then
                        Exit For
                    End If
                    GeneralTimoutWaiter.SetWaitingConditionNotSatisfied()
                    GeneralTimoutWaiter.WaitForEventOrOperationFinished(100)
                    DoEvents()
                Next
                ViewSourceButton.click()
                For RetryCounter = 1 To 5
                    GeneralTimoutWaiter.SetWaitingConditionNotSatisfied()
                    GeneralTimoutWaiter.WaitForEventOrOperationFinished(2500)
                    CurrentHTMLDocument = wbbWikidotSiteContainer.Document
                    PossibleElements = CurrentHTMLDocument.getElementsByTagName("div")
                    Dim IsPageSourceDivFound As Boolean = False
                    For Each CurrentElement In PossibleElements
                        If CurrentElement.className = "page-source" Then
                            IsPageSourceDivFound = True
                            Dim OutputFileName As String
                            If PageList(i).IsPageCategoryNotDefault Then
                                OutputFileName = PageList(i).PageCategory & "_" & PageList(i).PageUnixName
                            Else
                                OutputFileName = PageList(i).PageUnixName
                            End If
                            OutputFileName = OutputFileName.Replace(":", "_")
                            OutputFileName = SavePath & "\source\" & OutputFileName & ".txt"
                            Dim OutputStream As FileStream = File.Open(OutputFileName, FileMode.Create, FileAccess.Write)
                            Dim DataArray() As Byte = Encoding.UTF8.GetBytes(CurrentElement.innerText)
                            OutputStream.Write(DataArray, 0, DataArray.Length)
                            OutputStream.Flush()
                            OutputStream.Close()
                            Exit For
                        End If
                    Next
                    If IsPageSourceDivFound Then
                        RetryCounter = 0
                        Exit For
                    End If
                    DoEvents()
                Next
                If RetryCounter <> 0 Then
                    WriteLog("Failed to save page source of page '" & PageList(i).PageFullName & "' after 5 tries, page skipped.")
                    FailedCount += 1
                    Continue For
                End If
            Catch ex As Exception
                WriteLog("Failed to save page source of page '" & PageList(i).PageFullName & "', page skipped.")
                FailedCount += 1
                Continue For
            End Try
            'Save attached files
            CurrentHTMLDocument = wbbWikidotSiteContainer.Document
            Dim FilesButton As IHTMLElement = Nothing
            For RetryCounter = 1 To 5
                FilesButton = CurrentHTMLDocument.getElementById("files-button")
                If Not IsNothing(FilesButton) Then
                    Exit For
                End If
                GeneralTimoutWaiter.SetWaitingConditionNotSatisfied()
                GeneralTimoutWaiter.WaitForEventOrOperationFinished(100)
                DoEvents()
            Next
            Try
                FilesButton.click()
                GeneralTimoutWaiter.SetWaitingConditionNotSatisfied()
                GeneralTimoutWaiter.WaitForEventOrOperationFinished(2500)
                CurrentHTMLDocument = wbbWikidotSiteContainer.Document
                PossibleElements = CurrentHTMLDocument.getElementsByTagName("tr")
                For Each CurrentElement In PossibleElements
                    If IsNothing(CurrentElement.id) Then
                        Continue For
                    End If
                    If CurrentElement.id.StartsWith("file-row-") Then
                        ChildrenCollection = CurrentElement.children
                        ChildElement = ChildrenCollection(0) 'Get first <td>
                        ChildrenCollection = ChildElement.children
                        ChildElement = ChildrenCollection(0)
                        Dim FileLinkElement As IHTMLAnchorElement = ChildElement
                        Dim FileUrl As String = FileLinkElement.href
                        Dim FileName As String = ChildElement.innerText
                        Dim DownloadPath As String
                        If PageList(i).IsPageCategoryNotDefault Then
                            DownloadPath = PageList(i).PageCategory & "_" & PageList(i).PageUnixName
                        Else
                            DownloadPath = PageList(i).PageUnixName
                        End If
                        DownloadPath = DownloadPath.Replace(":", "_")
                        DownloadPath = SavePath & "\files\" & DownloadPath
                        Dim DownloadedFilePath As String = DownloadPath & "\" & FileName

                        WriteLog("Found attached file '" & FileName & "'.")
                        Try
                            If Not Directory.Exists(DownloadPath) Then
                                Directory.CreateDirectory(DownloadPath)
                            End If
                            Dim FileDownloader As New WebClient
                            FileDownloader.DownloadFile(FileUrl, DownloadedFilePath)
                            WriteLog("Downloaded file '" & FileName & "' from URL '" & FileUrl & "' to local path '" & DownloadedFilePath & "' successfully.")
                        Catch ex As Exception
                            WriteLog("Failed to download file '" & FileName & "' from URL '" & FileUrl & "' to local path '" & DownloadedFilePath & "'.")
                        End Try
                    End If
                    DoEvents()
                Next
            Catch ex As Exception
                WriteLog("Failed to check and save attached file(s) of page '" & PageList(i).PageFullName & "', skipped.")
                FailedCount += 1
                Continue For
            End Try
            'Count
            SucceededCount += 1
            prgProgress.Value = i + 1
        Next

        'Finalize
        MessageBox.Show("Operation finished." & vbCrLf & SucceededCount.ToString() & " pages processed successfully, " & FailedCount.ToString() & " pages was not backed up due to error(s).")
        WriteLog("Operation finished. " & SucceededCount.ToString() & " pages processed successfully, " & FailedCount.ToString() & " pages was not fully backed up due to error(s).")

        'Unlock UI
        UnlockUi()
        prgProgress.Value = 0
    End Sub
End Class
