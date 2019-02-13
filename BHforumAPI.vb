Imports System.Net
Imports System.Net.Http
Imports HtmlAgilityPack
Imports Newtonsoft.Json

Public Class BHforumAPI
    Private Shared cookies As New CookieContainer,
        client As New HttpClient(New HttpClientHandler() With {.CookieContainer = cookies}),
        cookieUri As New Uri("https://forum.gamer.com.tw"),
        web As New HtmlWeb,
        html As HtmlDocument

    ''' <summary>
    ''' 取得看板文章列表，每頁有30筆資料
    ''' </summary>
    ''' <param name="BoardID"></param>
    ''' <param name="Page"></param>
    ''' <returns></returns>
    Public Function GetTopics(ByVal BoardID As UInteger, Optional ByVal Page As UInteger = 1) As List(Of Topic) '平均每頁308.4ms
        'Documents:
        'http://html-agility-pack.net/select-nodes
        'https://stackoverflow.com/questions/1604471/how-can-i-find-an-element-by-css-class-with-xpath
        Dim url As String = "https://forum.gamer.com.tw/B.php?bsn=" & BoardID & "&page=" & Page
        html = Web.Load(url)
        Dim nodes = html.DocumentNode.SelectNodes("//tr[contains(@class, 'b-list__row')]"),
            result As New List(Of Topic)
        'parse data
        For Each i In nodes
            If i.InnerHtml.Contains("<div class=""attribution"">廣告</div>") Then Continue For '跳過巴哈廣告
            result.Add(ParseTopic(i))
        Next
        Return result
    End Function

    ''' <summary>
    ''' 使用哈拉板ID和文章ID取得特定文章的資料
    ''' </summary>
    ''' <param name="BoardID">哈拉板ID</param>
    ''' <param name="TopicID">文章ID</param>
    ''' <returns></returns>
    Public Function GetTopicByTopicID(ByVal BoardID As UInteger, ByVal TopicID As UInteger) As Topic
        html = Web.Load("https://forum.gamer.com.tw/C.php?bsn=" & BoardID & "&snA=" & TopicID)
        Dim title As String
        If html.DocumentNode.SelectSingleNode("//section").Attributes.Count = 1 Then
            '有多個頁數的文章，從第二個<section>取標題
            title = html.DocumentNode.SelectSingleNode("//section[2]/div[2]/div[1]/h1").InnerText
        Else
            title = html.DocumentNode.SelectSingleNode("//section[1]/div[2]/div[1]/h1").InnerText
        End If

        cookies.Add(cookieUri, New Cookie("ckFORUM_bsn", BoardID))
        cookies.Add(cookieUri, New Cookie("ckFORUM_stype", "title"))
        cookies.Add(cookieUri, New Cookie("ckFORUM_sval", Net.WebUtility.UrlEncode(title)))
        html.LoadHtml(client.GetStringAsync("https://forum.gamer.com.tw/B.php?bsn=" & BoardID & "&forumSearchQuery=" & title).Result)
        Dim nodes = html.DocumentNode.SelectNodes("//table[contains(@class, 'b-list')]/tr[contains(@class, 'b-list__row')]")
        For Each i In nodes
            If i.InnerHtml.Contains("<div class=""attribution"">廣告</div>") Then Continue For '跳過巴哈廣告
            If TopicID.ToString = i.SelectSingleNode("./td[contains(@class, 'b-list__summary')]/a").Attributes.Item(0).Value Then
                Return ParseTopic(i)
            End If
        Next
        Throw New TopicNotFound()
        Return Nothing
    End Function

    Private Function ParseTopic(TopicNode As HtmlNode) As Topic
        Dim result As New Topic
        result.BoardID = TopicNode.SelectSingleNode(".//p[contains(@class, 'b-list__summary__sort')]/a").Attributes.Item(0).Value.Split("?bsn=")(1).Replace("bsn=", "").Split("&")(0)
        result.TopicID = TopicNode.SelectSingleNode(".//p[contains(@class, 'b-list__time__edittime')]/a").Attributes.Item(2).Value.Split("&snA=")(1).Replace("snA=", "").Split("&")(0)
        result.AuthorID = TopicNode.SelectSingleNode(".//p[contains(@class, 'b-list__count__user')]/a").InnerText
        result.LastCommenterID = TopicNode.SelectSingleNode(".//p[contains(@class, 'b-list__time__user')]/a").InnerText
        result.LastCommentTime = TopicNode.SelectSingleNode(".//p[contains(@class, 'b-list__time__edittime')]/a").InnerText
        result.IsPinned = TopicNode.HasClass("b-list__row--sticky")
        result.IsExpertHighlight = TopicNode.SelectNodes("./a[contains(@class, 'is-expert-highlight')]") IsNot Nothing
        result.IsHighlight = TopicNode.SelectNodes(".//a[contains(@class, 'is-highlight')]") IsNot Nothing
        result.IsLocked = TopicNode.SelectNodes(".//i[contains(@class, 'icon-lock')]") IsNot Nothing
        result.SubBoardID = TopicNode.SelectSingleNode(".//p[contains(@class, 'b-list__summary__sort')]/a").Attributes.Item(0).Value.Split("&subbsn=")(1).Replace("subbsn=", "").Split("&")(0)
        result.ReplyCount = UShort.Parse(TopicNode.SelectSingleNode(".//p[contains(@class, 'b-list__count__number')]/span[1]").InnerText)
        result.ViewCount = UInteger.Parse(TopicNode.SelectSingleNode(".//p[contains(@class, 'b-list__count__number')]/span[2]").InnerText)

        '取得文章標題
        If TopicNode.HasClass("b-list__row--delete") Then
            '取得已刪除文章的標題
            result.Title = Web.Load(result.GetDesktopURL).DocumentNode.SelectSingleNode("//h1[contains(@class, 'c-disable__title')]").InnerText
            result.IsDeleted = True
        Else
            '一般的文章
            result.Title = TopicNode.SelectSingleNode(".//a[contains(@class, 'b-list__main__title')]").InnerText
        End If

        '取得GP數量
        If TopicNode.SelectSingleNode(".//span[contains(@class, 'b-gp')]") IsNot Nothing Then
            result.TotalGPCount = TopicNode.SelectSingleNode(".//span[contains(@class, 'b-gp')]").InnerText
        End If
        Return result
    End Function

    Public Class TopicNotFound : Inherits Exception
        Sub New()
            MyBase.New("找不到指定的文章")
        End Sub
    End Class
End Class

Public Class Topic
    Public BoardID As UInteger,                '哈拉板ID
        TopicID As UInteger,                   '文章ID
        Title As String,                       '標題
        AuthorID As String,                    '發文者ID
        LastCommenterID As String,             '最後回覆者的ID
        LastCommentTime As String,             '最後發表時間
        IsPinned As Boolean,                   '是否置頂
        IsExpertHighlight As Boolean,          '是否為達人文
        IsHighlight As Boolean,                '是否為高亮文章
        IsDeleted As Boolean,                  '是否為已刪除的文章
        IsLocked As Boolean,                   '是否為已鎖定的文章
        SubBoardID As Byte,                    '文章所在子板的ID
        ReplyCount As UShort,                  '文章的回覆數量
        ViewCount As UInteger,                 '文章的人氣(點擊人次)  計算公式未知
        TotalGPCount As UInteger               '文章所有樓層GP數的加總  (顯示在文章列表的數字)

    Private Shared web As New HtmlWeb,
        cookies As New CookieContainer,
        client As New HttpClient(New HttpClientHandler() With {.CookieContainer = cookies}),
        html As HtmlDocument

    Public Shared saveArticleContent As Boolean = True

    Sub New()
        cookies.Add(New Uri("https://forum.gamer.com.tw"), New Cookie("ckFORUM_setting", "000000000000000200"))
    End Sub

    ''' <summary>
    ''' 取得該文章裡面的貼文，一頁20個
    ''' </summary>
    ''' <param name="Page"></param>
    ''' <returns></returns>
    Public Function GetPosts(Optional ByVal Page As UInteger = 1) As List(Of Post)  '平均每頁250ms，此函數會觸發GC
        html = web.Load("https://forum.gamer.com.tw/C.php?bsn=" & BoardID & "&snA=" & TopicID & "&page=" & Page)  '大約每頁花220ms讀取
        Dim post As Post,
            result As New List(Of Post)
        For Each i In html.DocumentNode.SelectNodes("//section[contains(@class, 'c-section')]")
            If i.Attributes.Count = 2 Then '過濾掉無關的網頁元素
                If i.Id.StartsWith("post") Then
                    post = New Post
                    post.BoardID = BoardID
                    post.PostID = i.Attributes.Item(1).Value.Split("_")(1)
                    post.isModified = i.SelectSingleNode(".//div[contains(@class, 'c-post__header__info')]/a").InnerText.EndsWith(" 編輯")
                    post.PosterID = i.SelectSingleNode(".//a[contains(@class, 'userid')]").InnerText
                    post.PosterNick = i.SelectSingleNode(".//a[contains(@class, 'username')]").InnerText
                    post.PosterLV = i.SelectSingleNode(".//div[contains(@class, 'userlevel')]").InnerText.Split(ChrW(10))(2)
                    post.PosterGPCount = i.SelectSingleNode(".//div[contains(@class, 'usergp')]").Attributes.Item(1).Value
                    post.PosterCareer = i.SelectSingleNode(".//div[contains(@class, 'usercareer')]").Attributes.Item(1).Value
                    post.PosterRace = i.SelectSingleNode(".//div[contains(@class, 'userrace')]").Attributes.Item(1).Value
                    'post.GPCount = i.SelectSingleNode(".//span[contains(@class, 'postgp')]/span").InnerText
                    If i.SelectSingleNode(".//span[contains(@class, 'postgp')]/span").InnerText = "爆" Then
                        post.GPCount = post.GetGPUserList.Count
                    Else
                        post.GPCount = i.SelectSingleNode(".//span[contains(@class, 'postgp')]/span").InnerText
                    End If
                    post.PostTime = i.SelectSingleNode(".//div[contains(@class, 'c-post__header__info')]/a").Attributes.Item(3).Value
                    post.ModifyTime = i.SelectSingleNode(".//div[contains(@class, 'c-post__header__info')]/a").InnerText.Replace(" 編輯", "")
                    post.IP = i.SelectSingleNode(".//div[contains(@class, 'c-post__header__info')]/a").Attributes.Item(2).Value
                    post.Floor = i.SelectSingleNode(".//a[contains(@class, 'floor')]").Attributes.Item(2).Value
                    If saveArticleContent Then
                        post.ArticleContent = i.SelectSingleNode(".//div[contains(@class, 'c-article__content')]")
                    End If

                    '處理BP數量
                    If i.SelectSingleNode(".//span[contains(@class, 'postbp')]/span").InnerText = "-" Then
                        'BP數小於5不會顯示  當成數量=0
                        post.BPCount = 0
                    ElseIf i.SelectSingleNode(".//span[contains(@class, 'postbp')]/span").InnerText = "爆" Then
                        post.BPCount = post.GetBPUserList.Count
                    Else
                        post.BPCount = i.SelectSingleNode(".//span[contains(@class, 'postbp')]/span").InnerText
                    End If

                    '處裡簽名檔
                    If i.SelectSingleNode(".//div[contains(@class, 'post__body__signature')]/iframe") IsNot Nothing Then
                        post.SignatureURL = i.SelectSingleNode(".//div[contains(@class, 'post__body__signature')]/iframe").Attributes.Item(6).Value
                    End If
                    result.Add(post)
                Else
                    '貼文已經刪除
                    post = New Post
                    post.BoardID = BoardID
                    post.PostID = i.Attributes.Item(1).Value.Split("_")(1)
                    post.IsDeleted = True
                    post.Floor = i.SelectSingleNode(".//div[contains(@class, 'floor')]").Attributes.Item(1).Value
                    If saveArticleContent Then
                        post.ArticleContent = i.SelectSingleNode(".//div[contains(@class, 'hint')]")
                    End If
                    result.Add(post)
                End If
            ElseIf i.Attributes.Count = 3 Then
                '被摺疊的貼文

                '首先先刪除誤判的結果  摺疊會造成新增一筆認為貼文被刪除的紀錄
                result.RemoveAt(result.Count - 1)

                '處理貼文內容
                post = New Post
                post.BoardID = BoardID
                post.PostID = i.Attributes.Item(2).Value.Split("_")(1)
                post.isFolded = True
                post.isModified = i.SelectSingleNode(".//div[contains(@class, 'c-post__header__info')]/a").InnerText.EndsWith(" 編輯")
                post.PosterID = i.SelectSingleNode(".//a[contains(@class, 'userid')]").InnerText
                post.PosterNick = i.SelectSingleNode(".//a[contains(@class, 'username')]").InnerText
                post.PosterLV = i.SelectSingleNode(".//div[contains(@class, 'userlevel')]").InnerText.Split(ChrW(10))(2)
                post.PosterGPCount = i.SelectSingleNode(".//div[contains(@class, 'usergp')]").Attributes.Item(1).Value
                post.PosterCareer = i.SelectSingleNode(".//div[contains(@class, 'usercareer')]").Attributes.Item(1).Value
                post.PosterRace = i.SelectSingleNode(".//div[contains(@class, 'userrace')]").Attributes.Item(1).Value
                post.GPCount = i.SelectSingleNode(".//span[contains(@class, 'postgp')]/span").InnerText
                post.PostTime = i.SelectSingleNode(".//div[contains(@class, 'c-post__header__info')]/a").Attributes.Item(3).Value
                post.ModifyTime = i.SelectSingleNode(".//div[contains(@class, 'c-post__header__info')]/a").InnerText.Replace(" 編輯", "")
                post.IP = i.SelectSingleNode(".//div[contains(@class, 'c-post__header__info')]/a").Attributes.Item(2).Value
                post.Floor = i.SelectSingleNode(".//a[contains(@class, 'floor')]").Attributes.Item(2).Value
                If saveArticleContent Then
                    post.ArticleContent = i.SelectSingleNode(".//div[contains(@class, 'c-article__content')]")
                End If

                '處理BP數量
                If i.SelectSingleNode(".//span[contains(@class, 'postbp')]/span").InnerText = "-" Then
                    'BP數小於5不會顯示  當成數量=0
                    post.BPCount = 0
                Else
                    post.BPCount = i.SelectSingleNode(".//span[contains(@class, 'postbp')]/span").InnerText
                End If

                '處裡簽名檔
                If i.SelectSingleNode(".//div[contains(@class, 'post__body__signature')]/iframe") IsNot Nothing Then
                    post.SignatureURL = i.SelectSingleNode(".//div[contains(@class, 'post__body__signature')]/iframe").Attributes.Item(6).Value
                End If
                result.Add(post)
            End If
        Next
        Return result
    End Function

    ''' <summary>
    ''' 取得該文章裡的所有貼文，請注意這個函式執行之後會使用大量記憶體
    ''' </summary>
    ''' <returns></returns>
    Public Function GetAllPosts() As List(Of Post)
        Dim pageCount As UInteger = Math.Floor((ReplyCount + 1) / 20),
            result As New List(Of Post)
        If (ReplyCount + 1) Mod 20 > 0 Then
            pageCount += 1
        End If
        For i As UInteger = 1 To pageCount
            result.AddRange(GetPosts(i))
        Next
        Return result
    End Function

    ''' <summary>
    ''' 取得桌面版的文章網址
    ''' </summary>
    ''' <returns></returns>
    Public Function GetDesktopURL() As String
        Return "https://forum.gamer.com.tw/C.php?bsn=" & BoardID & "&snA=" & TopicID
    End Function

    ''' <summary>
    ''' 取得手機版的文章網址
    ''' </summary>
    ''' <returns></returns>
    Public Function GetMobileURL() As String
        Return "https://m.gamer.com.tw/forum/C.php?bsn=" & BoardID & "&snA=" & TopicID
    End Function

    Public Overrides Function ToString() As String
        Return "討論版ID        " & BoardID & vbCrLf &
            "文章ID          " & TopicID & vbCrLf &
            "標題            " & Title & vbCrLf &
            "發文者ID        " & AuthorID & vbCrLf &
            "最後回覆者ID    " & LastCommenterID & vbCrLf &
            "最後回覆時間    " & LastCommentTime & vbCrLf &
            "被置頂          " & IsPinned & vbCrLf &
            "達人文          " & IsExpertHighlight & vbCrLf &
            "高亮            " & IsHighlight & vbCrLf &
            "已刪文          " & IsDeleted & vbCrLf &
            "已鎖定          " & IsLocked & vbCrLf &
            "文章子版ID      " & SubBoardID & vbCrLf &
            "回覆數量        " & ReplyCount & vbCrLf &
            "人氣            " & ViewCount & vbCrLf &
            "PC版文章網址    " & GetDesktopURL() & vbCrLf &
            "手機版文章網址  " & GetMobileURL() & vbCrLf &
            "GP總數          " & TotalGPCount & vbCrLf
    End Function
End Class

Public Class Post
    Public Floor As UInteger,                  '樓層數
        BoardID As UInteger,                   '哈拉板ID
        PostID As UInteger,                    '貼文ID
        IsDeleted As Boolean,                  '是否被刪文
        isFolded As Boolean,                   '是否被折疊
        isModified As Boolean,                 '內容是否有修改過
        PosterID As String,                    '貼文者ID
        PosterNick As String,                  '貼文者的暱稱
        PosterLV As UInteger,                  '貼文者的巴哈等級
        PosterGPCount As UInteger,             '貼文者在巴哈得到的GP總數
        PosterCareer As String,                '貼文者的巴哈職業
        PosterRace As String,                  '貼文者的巴哈種族
        GPCount As UInteger,                   '這篇貼文得到的GP數
        BPCount As UInteger,                   '這篇貼文得到的BP數
        PostTime As String,                    '貼文發出的時間
        ModifyTime As String,                  '貼文修改的時間  (如果沒有修改過則會與發文時間一致)
        IP As String,                          '貼文者的IP  最後三碼是XXX
        SignatureURL As String,                '簽名檔的網址(如果沒有簽名檔則該欄位為空)
        ArticleContent As HtmlNode             '文章內容(html)

    Private Shared client As New HttpClient


    ''' <summary>
    ''' 取得該貼文的所有留言，數量可能為0並回傳空的List
    ''' </summary>
    ''' <returns></returns>
    Public Function GetComments() As List(Of Comment)
        Dim result As New List(Of Comment)
        Dim url = "https://forum.gamer.com.tw/ajax/moreCommend.php?bsn=" & BoardID & "&snB=" & PostID & "&returnHtml=0"
        Dim comments As Linq.JObject = (New JsonSerializer).Deserialize(New JsonTextReader(New IO.StreamReader(client.GetStreamAsync(url).Result)))
        comments.Remove("next_snC")
        For Each i In comments
            result.Add(i.Value.ToObject(Of Comment))
        Next
        '巴哈回傳的JSON格式是最後的留言在最前面  最早的留言在最後面
        '我覺得應該反過來  所以呼叫了.Reverse
        result.Reverse()
        Return result
    End Function

    ''' <summary>
    ''' 取得給該貼文GP的使用者資料
    ''' </summary>
    ''' <returns></returns>
    Public Function GetGPUserList() As List(Of GBPUser)
        Return GetGBPUserList(1)
    End Function

    ''' <summary>
    ''' 取得給該貼文BP的使用者資料
    ''' </summary>
    ''' <returns></returns>
    Public Function GetBPUserList() As List(Of GBPUser)
        Return GetGBPUserList(2)
    End Function

    Private Function GetGBPUserList(ByVal mode As UInteger) As List(Of GBPUser)
        'Mode=1  =>  取得GP資料
        'Mode=2  =>  取得BP資料

        Dim result As New List(Of GBPUser),
            pagecount As UInt16 = 1

        Dim Data As New Dictionary(Of String, String)
        Data.Add("t", mode)
        Data.Add("bsn", BoardID)
        Data.Add("snB", PostID)
        Data.Add("p", 1)

        Dim response As String,
            parsedResult As Linq.JObject,
            user As New GBPUser

        While True
            Data.Item("p") = pagecount
            response = client.PostAsync("https://forum.gamer.com.tw/ajax/GPBPlist.php", New FormUrlEncodedContent(Data)).Result.Content.ReadAsStringAsync.Result
            parsedResult = JsonConvert.DeserializeObject(response)
            If parsedResult.Value(Of String)("status") = "F" Then
                '已經沒有資料了
                Exit While
            End If
            parsedResult = JsonConvert.DeserializeObject(parsedResult("u"))
            For Each j In parsedResult
                user.UserID = j.Key
                user.UserNick = j.Value.First
                result.Add(user)
            Next
            pagecount += 1
        End While

        '巴哈回傳的JSON格式是最後G/BP的人在最前面  最早G/BP的人在最後面
        '我覺得應該反過來  所以呼叫了.Reverse
        result.Reverse()
        Return result
    End Function

    Public Overrides Function ToString() As String
        Return "Floor:           " & Floor & vbCrLf &
            "BoardID:         " & BoardID & vbCrLf &
            "PostID:          " & PostID & vbCrLf &
            "IsDeleted:       " & IsDeleted & vbCrLf &
            "isFolded:        " & isFolded & vbCrLf &
            "isModified:      " & isModified & vbCrLf &
            "PosterID:        " & PosterID & vbCrLf &
            "PosterNick:      " & PosterNick & vbCrLf &
            "PosterLV:        " & PosterLV & vbCrLf &
            "PosterGPCount:   " & PosterGPCount & vbCrLf &
            "PosterCareer:    " & PosterCareer & vbCrLf &
            "PosterRace:      " & PosterRace & vbCrLf &
            "GPCount:         " & GPCount & vbCrLf &
            "BPCount:         " & BPCount & vbCrLf &
            "PostTime:        " & PostTime & vbCrLf &
            "ModifyTime:      " & ModifyTime & vbCrLf &
            "IP:              " & IP & vbCrLf &
            "SignatureURL:    " & SignatureURL
    End Function
End Class

Public Class GBPUser
    Public UserID As String,                   '給予G/BP的使用者ID
        UserNick As String                     '給予G/BP的使用者暱稱  如果使用者被刪除帳號則該欄位為空
End Class

Public Class Comment
    Public bsn As UInteger,                    '哈拉板ID
        sn As UInteger,                        '留言ID
        userid As String,                      '留言者ID
        comment As String,                     '留言內容
        gp As UInteger,                        '該留言得到的 推 數量
        bp As UInteger,                        '該留言得到的 噓 數量
        wtime As String,                       '留言時間
        mtime As String,                       '修改時間  如果流言沒被修改則值為  "0000-00-00 00:00:00"
        delreason As String,                   '刪除理由  (?)
        state As String,                       '(不知道...)
        type As String,                        '(不知道...)
        content As String,                     '不知道，值與comment一樣
        snB As UInteger,                       '貼文ID
        time As String,                        '相較於現在的時間的發文時間  如 "38分前"  "前天 08:49"  "2018-08-22 15:38:21"
        nick As String                         '留言者暱稱
End Class