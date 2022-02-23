''' <summary>
''' Protocol type of a URL
''' </summary>
''' <remarks></remarks>
Public Enum WikidotSiteHttpProtocol
    Unknown = 0
    HTTP = 1
    HTTPS = 2
End Enum
''' <summary>
''' Information about a Wikidot page
''' </summary>
''' <remarks></remarks>
Public Class WikidotPage
    ''' <summary>
    ''' Full URL of the page, looks like "http(s)://site.name/(category:)page-name(/../...)".
    ''' </summary>
    ''' <remarks></remarks>
    Public PageUrl As String
    ''' <summary>
    ''' Marks if the URL is valid.
    ''' </summary>
    ''' <remarks></remarks>
    Public IsPageUrlValid As Boolean
    ''' <summary>
    ''' Belonging site of this page.
    ''' </summary>
    ''' <remarks></remarks>
    Public PageSite As String
    ''' <summary>
    ''' Full name of this page, looks like "(category:)page-name".
    ''' </summary>
    ''' <remarks></remarks>
    Public PageFullName As String
    ''' <summary>
    ''' Unix name of the page.
    ''' </summary>
    ''' <remarks></remarks>
    Public PageUnixName As String
    ''' <summary>
    ''' Category of the page
    ''' </summary>
    ''' <remarks></remarks>
    Public PageCategory As String
    ''' <summary>
    ''' Protocol
    ''' </summary>
    ''' <remarks></remarks>
    Public SiteProtocol As WikidotSiteHttpProtocol
    Public Sub New(Url As String)
        ParsePageUrl(Url)
    End Sub
    Public Sub SetPageUrl(Url As String)
        ParsePageUrl(Url)
    End Sub
    Public Function IsPageCategoryNotDefault() As Boolean
        Return PageCategory <> ""
    End Function
    Private Sub ParsePageUrl(Url As String)
        'Avoid invalid URL
        If Url.Length = 0 Then
            PageUrl = ""
            IsPageUrlValid = False
            PageSite = ""
            PageFullName = ""
            PageUnixName = ""
            PageCategory = ""
            SiteProtocol = WikidotSiteHttpProtocol.Unknown
            Exit Sub
        End If

        'Replace all "\" to "/", and then split the URL
        Url = Url.Replace("\", "/")
        Dim UrlComponents() As String = Url.Split(New Char() {"/"c}, StringSplitOptions.RemoveEmptyEntries)

        'Analyze the URL
        PageUrl = Url
        If UrlComponents.Length >= 3 Then 'A valid URL shoud be like "http(s)://site.name/(category:)page-name(/../...)"
            'Get protocol
            Select Case UrlComponents(0).ToUpper
                Case "HTTP"
                    SiteProtocol = WikidotSiteHttpProtocol.HTTP
                    IsPageUrlValid = True
                Case "HTTPS"
                    SiteProtocol = WikidotSiteHttpProtocol.HTTPS
                    IsPageUrlValid = True
                Case Else
                    SiteProtocol = WikidotSiteHttpProtocol.Unknown
                    IsPageUrlValid = False
            End Select
            'Get site
            PageSite = UrlComponents(1)
            'Get full name
            PageFullName = UrlComponents(2)
            'Get category and unix name
            If PageFullName.Contains(":") Then
                Dim PageNameComponents() As String = PageFullName.Split(New Char() {":"c}, StringSplitOptions.RemoveEmptyEntries)
                If PageNameComponents.Length = 0 Then 'Only a sequence of ":"?
                    PageCategory = ""
                    PageUnixName = ""
                    IsPageUrlValid = False
                ElseIf PageNameComponents.Length = 1 Then 'Only unix name like ":page" or "page:"
                    PageCategory = ""
                    PageUnixName = PageNameComponents(0)
                ElseIf PageNameComponents.Length = 2 Then 'Typical full name like "category:page"
                    PageCategory = PageNameComponents(0)
                    PageUnixName = PageNameComponents(1)
                Else 'Full name like "category:subcat1:subcat2:page"
                    PageCategory = PageNameComponents(0)
                    Dim i As Integer
                    For i = 1 To PageNameComponents.GetUpperBound(0)
                        PageUnixName += PageNameComponents(i)
                        If i <> PageNameComponents.GetUpperBound(0) Then
                            PageUnixName += ":"
                        End If
                    Next
                End If
            Else
                PageCategory = ""
                PageUnixName = PageFullName
            End If
        End If
    End Sub
End Class
