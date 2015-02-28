
Imports System.Net.Sockets
Imports System.Net
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices


Public Class sheet1
    Public a, b, yyy, lv As String
    Public tokenlog As CookieContainer = Login("めあど", "ぱすわーど", 190823874)

    Public token As String
    Public html, addr, port, thread As String
    'TextBox1.Text内で正規表現と一致する対象を1つ検索 
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        lv = getlv()
        Main()
        '   aaa(Integer.Parse(thread), addr, Integer.Parse(port))

    End Sub
    Public fuzai As Integer = 0

    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles Button2.Click


        fuzai = 1
        lv = getlv()
        Main()

        '   aaa(Integer.Parse(thread), addr, Integer.Parse(port))

    End Sub
    Public commanderas() As String = {Nothing, Nothing, Nothing, Nothing, Nothing, Nothing}
    Public oncommander As Integer
    Public commanderhash As String = Nothing

    Function newcommander(ByVal commentnumber As String, ByVal nextpre As Integer) As Integer


        ''''commanderにふさわしいナンバーを格納していく、preを受け取って
        ''''コマンダーファンクション
        Dim ichiji As Integer = nextpre
        newcommander = nextpre + 1

        commanderas(ichiji) = commentnumber

        If ichiji >= 3 Then
            'ｿｰﾄして...
            Array.Sort(commanderas)
            '比べる!!!

            If commanderas(5) = commanderas(4) Then
                If commanderas(4) = commanderas(3) Then


                    Dim hogehogehoge As String = ""
                    Dim search As New System.Text.RegularExpressions.Regex("[0-9]+", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                    Dim searchbox As System.Text.RegularExpressions.Match = search.Match(commanderas(5))

                    While searchbox.Success

                        '一致した対象が見つかったときキャプチャした部分文字列を表示 
                        hogehogehoge = searchbox.Value
                        searchbox = searchbox.NextMatch()

                    End While

                    Dim hashcolumn As Integer = Range("K1:K" + x.ToString).Find(hogehogehoge).Row

                    commanderhash = Cells(hashcolumn, 11).Value


                    If commanderhash = "" Then

                        commanderhash = "nothing hashman"
                        newcommander = 0
                        commanderas = {Nothing, Nothing, Nothing, Nothing, Nothing, Nothing}

                    Else
                        commanderas = {Nothing, Nothing, Nothing, Nothing, Nothing, Nothing}

                        newcommander = 0

                    End If

                End If
            End If
        End If



    End Function

    '例
    'http://watch.live.nicovideo.jp/api/broadcast/lv191612346?body=%E3%83%86%E3%82%B9%E3%83%88&mail=perm&token=4f339296c70b4164097bee560505251be6972e9a

    'http://live.nicovideo.jp/api/getpublishstatus




    Public Sub commderpost(ByVal body As String)
        '/perm コマンダー:aaaaaaa
        body = Replace(body, "a", "エー")
        body = Replace(body, "b", "エー")
        body = Replace(body, "A", "エー")
        body = Replace(body, "B", "ビー")
        body = Replace(body, "→", "ミギ")
        body = Replace(body, "右", "ミギ")
        body = Replace(body, "←", "ヒダリ")
        body = Replace(body, "左", "ヒダリ")
        body = Replace(body, "↑", "ウエ")
        body = Replace(body, "上", "ウエ")
        body = Replace(body, "↓", "シタ")
        body = Replace(body, "下", "シタ")
        body = Replace(body, "start", "スタート")
        body = Replace(body, "select", "セレクト")
        body = Replace(body, "ramdom", "ランダム")

        body = UrlEncodeUtf8(body)
        Dim POSTURL As String = "http://watch.live.nicovideo.jp/api/broadcast/lv" + lv + "?body=%e3%82%b3%e3%83%9e%e3%83%b3%e3%83%80%e3%83%bc%3a" + body + "&mail=perm&token=" + token
        html = Read(tokenlog, POSTURL)
    End Sub
    Sub Main()
        Dim cc As CookieContainer = Login("捨てアカウントメアド", "すてあかウントパスワード", 190823874)
        'マイページのHTMLを取得する
        a = Read(cc, "http://www.nicovideo.jp/my/top")
        Dim html As String
        html = Read(cc, "http://watch.live.nicovideo.jp/api/getplayerstatus?v=lv" & lv)

        'Dim client As System.Net.WebClient = New System.Net.WebClient()

        '  Dim url2 As String
        ' url2 = "http://watch.live.nicovideo.jp/api/getplayerstatus?v=lv" & lv
        '指定したURLからデータを取得する
        'Dim wkStream As System.IO.Stream = client.OpenRead(url2)
        'エンコード指定で文字列を取得する
        'サイトによってエンコードは異なる
        'Dim sr As StreamReader = New StreamReader(wkStream, System.Text.Encoding.GetEncoding("utf-8"))
        'Dim html As String = sr.ReadToEnd()
        'sr.Close()
        '    wkStream.Close()


        Dim raddr As New System.Text.RegularExpressions.Regex("<addr>[0-9.a-z]+</addr>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        Dim rthread As New System.Text.RegularExpressions.Regex("<thread>[0-9]+</thread>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        dim rport As New System.Text.RegularExpressions.Regex("<port>[0-9.a-z]+</port>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        Dim p As System.Text.RegularExpressions.Match = rport.Match(html)
        Dim q As System.Text.RegularExpressions.Match = rthread.Match(html)
        Dim g As System.Text.RegularExpressions.Match = raddr.Match(html)

        While p.Success
            '一致した対象が見つかったときキャプチャした部分文字列を表示 
            port = p.Value
            '次に一致する対象を検索 
            p = p.NextMatch()
        End While


        While q.Success
            '一致した対象が見つかったときキャプチャした部分文字列を表示 
            thread = q.Value
            '次に一致する対象を検索 
            q = q.NextMatch()
        End While
        While g.Success
            '一致した対象が見つかったときキャプチャした部分文字列を表示 
            addr = g.Value
            '次に一致する対象を検索 
            g = g.NextMatch()
        End While


        port = Replace(port, "<port>", "")
        port = Replace(port, "</port>", "")
        thread = Replace(thread, "<thread>", "")
        thread = Replace(thread, "</thread>", "")
        addr = Replace(addr, "<addr>", "")
        addr = Replace(addr, "</addr>", "")
        'getflv APIで組曲「ニコニコ動画」の動画保存場所を取得する
        ' b = Read(cc, "http://flapi.nicovideo.jp/api/getflv/sm500873")

        'http://live.nicovideo.jp/api/getpublishstatus

        aaa(Integer.Parse(thread), addr, Integer.Parse(port))

    End Sub

    'ニコニコ動画にログインして会員ページに必要なクッキーを返します。
    Public Function Login(ByVal Mail As String, ByVal Password As String, ByVal lv As Integer) As CookieContainer
        'データをPOSTする
        Dim content As String = "mail=" & Mail & "&password=" & Password
        Dim contentBytes As Byte() = Encoding.ASCII.GetBytes(content)
        Dim request As HttpWebRequest = HttpWebRequest.CreateHttp(
            "https://secure.nicovideo.jp/secure/login?site=niconico")
        request.CookieContainer = New CookieContainer
        request.Method = "POST"
        request.ContentType = "application/x-www-form-urlencoded"
        request.ContentLength = contentBytes.Length
        Using stream As Stream = request.GetRequestStream()
            stream.Write(contentBytes, 0, contentBytes.Length)
        End Using

        '応答を確認してcookieを取得する
        Using response As HttpWebResponse = request.GetResponse()
            Return request.CookieContainer
        End Using

    End Function

    'GETを送信してHTML、XML等を取得します。
    Public Function Read(ByRef cc As CookieContainer, ByVal URL As String) As String
        Dim request As HttpWebRequest = HttpWebRequest.CreateHttp(URL)
        request.CookieContainer = cc
        Using response As HttpWebResponse = request.GetResponse()
            Using reader As New StreamReader(response.GetResponseStream())
                Return reader.ReadToEnd
            End Using
        End Using
    End Function
    Public d, e, commentText, commentNo, commentUserID As String
    Public x As Integer

    Function aaa(ByVal thread As Integer, ByVal addr As String, ByVal port As Integer) As String

        x = 1

        'http://watch.live.nicovideo.jp/api/getplayerstatus?v=lv何とか
        'msの値
        '取ってくるのはまた別に書くよ！
        '  Dim thread As Integer = 1374838387
        '  Dim addr As String = "msg104.live.nicovideo.jp"
        ' Dim port As Integer = 2811
        '  winwin()
        '  System.Threading.Thread.Sleep(300)
        Dim jikkou As Integer = Cells(9, 5).value

        Dim more As String = ""
        'res_from で過去ログ取得 -10にすると、-10コメント前を取ってくる
        Dim req = String.Format("<thread thread=""{0}"" version=""20061206"" res_from=""-10"" /> ", thread)
        Dim tcp As New TcpClient(addr, Integer.Parse(port.ToString))
        Dim ns As NetworkStream = tcp.GetStream()
        Dim sendBytes As Byte() = Encoding.UTF8.GetBytes(req)
        sendBytes(sendBytes.Length - 1) = 0

        ns.Write(sendBytes, 0, sendBytes.Length)

        Dim pre As Integer = 0
        Dim resSize As Integer

        Do
            Dim resBytes As Byte() = New Byte(2048) {}

            resSize = ns.Read(resBytes, 0, resBytes.Length)
            If resSize = 0 Then
                Exit Do
            End If

            Dim message As String = Encoding.UTF8.GetString(resBytes)
            'Console.WriteLine(message)

            message = more & message
            more = ""
            Dim elements As String() = message.Split(New String() {vbNullChar}, StringSplitOptions.RemoveEmptyEntries)

            '<chat thread="1062986843" no="1" vpos="6300" date="1293718792" mail="184" user_id="mwR5e8FptFf6-O3gZtQ3ceHAgPU" premium="3" anonymity="1">テスト！</chat>

            For Each receiveData As String In elements

                '帰ってきたXMLに"<chat"から"</chat>"まで全部あったらで分ける
                '文字が多いいと、分割して帰ってくるから、
                If receiveData.StartsWith("<chat") AndAlso receiveData.EndsWith("</chat>") Then


                    OnReceiveChat(receiveData)


                    Cells(x, 11).value = commentNo
                    Cells(x, 12).value = commentText
                    'commentNo = comment.@no
                    'commentDate = comment.@date
                    'commentVpos = comment.@vpos
                    'commentMaill = comment.@mail
                    'commentUserID = comment.@user_id
                    'commentPremium = comment.@premium

                    Cells(x, 11) = commentUserID

                    '   If commentUserID = commanderhash Then

                    ' commderpost(commentText)

                    'End If



                    If commentText = "commentdoing" Then

                        AppActivate("bgb")

                    End If



                    '/disconnect

                    If commentText = "/disconnect" Then


                        html = Nothing
                        addr = Nothing
                        port = Nothing
                        thread = Nothing
                        lv = Nothing


                        System.Windows.Forms.Application.DoEvents()
                        savepoint()



                        lv = getlv()

                        Range(Cells(1, 11), Cells(x, 11)).Value = ""
                        Range(Cells(1, 13), Cells(x, 13)).Value = ""
                        Range(Cells(1, 12), Cells(x, 12)).Value = ""
                        Main()


                    End If

                    Cells(9, 3).value = x

                    System.Windows.Forms.Application.DoEvents()

                    Dim selcommand As New System.Text.RegularExpressions.Regex("[ABab←→↑↓左右上下]|start|select|random", System.Text.RegularExpressions.RegexOptions.IgnoreCase)

                    Dim hoge As Integer
                    Dim hidari, migi, ue, shita, ei, bii, start, selects, randoms As Integer

                    Dim rara As String
                    Dim qwe As Integer = 0
                    Dim rty As Integer = 0
                    Dim newcommand(100) As String
                    rara = commentText
                    Dim commander As System.Text.RegularExpressions.Match = selcommand.Match(rara)

                    While commander.Success

                        '一致した対象が見つかったときキャプチャした部分文字列を表示 
                        newcommand(rty) = commander.Value
                        rty = rty + 1
                        '次に一致する対象を検索 

                        commander = commander.NextMatch()

                    End While



                    '入力判定

                    If newcommand(0) = "a" Then
                        ei = ei + 1

                    ElseIf newcommand(0) = "A" Then
                        ei = ei + 1

                    ElseIf newcommand(0) = "b" Then
                        bii = bii + 1

                    ElseIf newcommand(0) = "B" Then
                        bii = bii + 1

                    ElseIf newcommand(0) = "左" Then
                        hidari = hidari + 1

                    ElseIf newcommand(0) = "←" Then
                        hidari = hidari + 1

                    ElseIf newcommand(0) = "右" Then
                        migi = migi + 1

                    ElseIf newcommand(0) = "→" Then
                        migi = migi + 1

                    ElseIf newcommand(0) = "上" Then
                        ue = ue + 1

                    ElseIf newcommand(0) = "↑" Then
                        ue = ue + 1

                    ElseIf newcommand(0) = "↓" Then
                        shita = shita + 1

                    ElseIf newcommand(0) = "下" Then
                        shita = shita + 1

                    ElseIf newcommand(0) = "start" Then
                        start = start + 1

                    ElseIf newcommand(0) = "select" Then
                        selects = selects + 1

                    ElseIf newcommand(0) = "random" Then
                        randoms = randoms + 1
                    End If


                    qwe = qwe + 1

                    '表示更新
                    Cells(1, 2).font.size = "28"
                    Cells(1, 4).font.size = "28"
                    Cells(1, 6).font.size = "28"
                    Cells(1, 8).font.size = "28"
                    Cells(1, 10).font.size = "28"
                    Cells(4, 2).font.size = "28"
                    Cells(4, 4).font.size = "28"
                    Cells(4, 6).font.size = "28"
                    Cells(4, 8).font.size = "28"


                    Cells(1, 2) = ei
                    Cells(1, 4) = bii
                    Cells(1, 6) = ue
                    Cells(1, 8) = shita
                    Cells(1, 10) = migi
                    Cells(4, 2) = hidari
                    Cells(4, 4) = start
                    Cells(4, 6) = selects
                    Cells(4, 8) = randoms

                    ''''''''コマンダーオプション'''''''''''''''''''''''''    '''''''''''''''''''''''''''''''''''''''    '''''''''''''''''''''''''''''''''''''''
                    Dim cocoa As String = ""
                    Dim selcommander As New System.Text.RegularExpressions.Regex("der:[0-9]+", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                    Dim regcommander As System.Text.RegularExpressions.Match = selcommander.Match(rara)

                    While regcommander.Success
                        '一致した対象が見つかったときキャプチャした部分文字列を表示 
                        cocoa = regcommander.Value
                        '次に一致する対象を検索 
                        regcommander = regcommander.NextMatch()
                    End While


                    If Not cocoa = Nothing Then

                        pre = newcommander(cocoa, pre)

                    End If

                    '''''''''''''''''''''''''''''''''''''''    '''''''''''''''''''''''''''''''''''''''    '''''''''''''''''''''''''''''''''''''''    '''''''''''''''''''''''''''''''''''''''


                    hoge = hidari + migi + ue + shita + ei + bii + start + selects + randoms


                    If x Mod jikkou = 0 Then

                        shuukei(hoge, hidari, migi, ue, shita, ei, bii, start, selects, randoms)

                        ei = 0
                        hidari = 0
                        migi = 0
                        ue = 0
                        shita = 0
                        bii = 0

                        start = 0
                        selects = 0
                        randoms = 0

                    End If



                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''7


                    System.Windows.Forms.Application.DoEvents()
                    x = x + 1

                    If x Mod 100 = 0 Then

                        savepoint()

                    End If


                Else
                    'チャットが分割で来たときの処理。
                    If receiveData.StartsWith("<thread resultcode=") = False Then
                        If receiveData.StartsWith("<view_counter") = True Then
                            'ここは、ニコニコ実況ようだから、スルーでおｋ
                        ElseIf receiveData.Contains("</chat>") = False Then
                            more = receiveData
                        End If
                    End If

                End If

            Next


        Loop While ns.CanRead
        Return e

    End Function

    Public Sub OnReceiveChat(ByVal receiveData As String)

        Dim commentVpos As String
        Dim commentDate As String
        Dim commentMaill As String
        'Dim commentUserID As String
        Dim commentPremium As String
        Dim commentAnonymity As String

        Dim commentThread As Integer
        Dim commentLoom As String = ""

        Dim userName As String = Nothing

        Dim comment = XElement.Parse(receiveData)
        commentNo = comment.@no
        commentDate = comment.@date
        commentVpos = comment.@vpos
        commentMaill = comment.@mail
        commentUserID = comment.@user_id
        commentPremium = comment.@premium

        '時々帰って来ない時がある。
        If commentPremium = Nothing Then
            commentPremium = "0"
        End If

        commentAnonymity = comment.@anonymity
        commentThread = CInt(comment.@thread)
        commentText = comment.Value

        'ユーザーレベル種分け
        If commentPremium = "0" Then
            commentPremium = "一般"
        ElseIf commentPremium = "1" Then
            commentPremium = "プレミアム"
        ElseIf commentPremium = "3" Then
            commentPremium = "放送主"
        ElseIf commentPremium = "7" Then
            commentPremium = "BSP"
        End If

        '0=一般
        '1=プレミアム
        '2=放送終了後に送られてくる、/disconnect のレベル
        '3=放送主
        '7=BSP

        '公式生見ると25とかある時がある。何故かはわからん

        d = commentNo & ":" & commentText

        'Console.Write(vbCrLf & CommentNo & " " & CommentVpos & " " & CommentDate & " " & CommentMaill & " " & CommentUserID & " " & CommentPremium & " " & CommentAnonymity & " " & CommentText)

    End Sub
    Public Sub shuukei(ByVal hoge1 As Integer, ByVal hidari1 As Integer, ByVal migi1 As Integer, ByVal ue1 As Integer, ByVal shita1 As Integer _
    , ByVal ei1 As Integer, ByVal bii1 As Integer, ByVal start1 As Integer, ByVal selects1 As Integer, ByVal randoms1 As Integer)


        If (hoge1 > 0) Then

            Dim arrayman() As Integer = {hidari1, migi1, ue1, shita1, ei1, bii1, start1, selects1, randoms1}
            Dim ingmax As Integer = 1
            Dim maxsoeji As Integer = 0

            Dim onaji(4), kaj, toto As Integer
            kaj = 0
            For lngCounter = 0 To 8

                If arrayman(lngCounter) >= ingmax Then


                    If arrayman(lngCounter) = ingmax Then

                        onaji(kaj) = lngCounter '添え字
                        kaj = kaj + 1
                        toto = arrayman(lngCounter)

                    End If

                    ingmax = arrayman(lngCounter)
                    maxsoeji = lngCounter

                End If

            Next lngCounter


            If toto = ingmax Then
                'random
                '0 ~ kaj分random
                Dim rr As New System.Random(1000)
                Dim i1 As Integer = rr.Next(kaj - 1)
                maxsoeji = onaji(i1)

            End If

            Cells(9, 1) = sentaku(maxsoeji)

        End If



    End Sub

    Public Function IsArrayEx(varArray As Object) As Long
        On Error GoTo ERROR_

        If IsArray(varArray) Then
            IsArrayEx = IIf(UBound(varArray) >= 0, 1, 0)
        Else
            IsArrayEx = -1
        End If

        Exit Function

ERROR_:
        If Err.Number = 9 Then
            IsArrayEx = 0
        End If
    End Function

    Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, _
                     ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Function sentaku(ByVal you As Integer) As String
        '  {hidari, migi, ue, shita, ei, bii, start, selects, randoms}
        '  A	B button
        'S	A button
        'Shift	Select button
        'Enter	Start button

        'Const VK_LBUTTON = &H1  '[LeftClick]
        'Const VK_RBUTTON = &H2  '[RightClick]
        'Const VK_CANCEL = &H3  '[Cancel]
        'Const VK_MBUTTON = &H4  '[MiddleClick]
        'Const VK_BACK = &H8  '[BackSpace]
        'Const VK_TAB = &H9  '[TAB]
        'Const VK_CLEAR = &HC  '[Clear]
        'Const VK_RETURN = &HD  '[Enter]
        'Const VK_SHIFT = &H10 '[LeftShift]
        'Const VK_RSHIFT = &HA1 '[RightShift]
        'Const VK_CONTROL = &H11 '[LeftControl]
        'Const VK_RCONTROL = &HA3 '[RightControl]
        'Const VK_MENU = &H12 '[LeftMenu(Alt)]
        'Const VK_RMENU = &HA5 '[RightMenu(Alt)]
        'Const VK_PAUSE = &H13 '[Pause]
        'Const VK_CAPITAL = &H14 '[CapsLock]
        'Const VK_ESCAPE = &H1B '[Esc]
        'Const VK_SPACE = &H20 '[Space]
        'Const VK_PRIOR = &H21 '[PageUp]
        'Const VK_NEXT = &H22 '[PageDown]
        'Const VK_END = &H23 '[End]
        'Const VK_HOME = &H24 '[Home]
        'Const VK_LEFT = &H25 '[←]
        'Const VK_UP = &H26 '[↑]
        'Const VK_RIGHT = &H27 '[→]
        'Const VK_DOWN = &H28 '[↓]
        'Const VK_SELECT = &H29 '[Select]
        'Const VK_PRINT = &H2A '[PrintScreen]
        'Const VK_EXECUTE = &H2B '[Execute]
        'Const VK_SNAPSHOT = &H2C '[SnapShot]
        'Const VK_INSERT = &H2D '[Insert]
        'Const VK_DELETE = &H2E '[Delete]
        'Const VK_HELP = &H2F '[Help]
        'Const VK_0 = &H30 '[0]
        'Const VK_1 = &H31 '[1]
        'Const VK_2 = &H32 '[2]
        'Const VK_3 = &H33 '[3]
        'Const VK_4 = &H34 '[4]
        'Const VK_5 = &H35 '[5]
        'Const VK_6 = &H36 '[6]
        'Const VK_7 = &H37 '[7]
        'Const VK_8 = &H38 '[8]
        'Const VK_9 = &H39 '[9]
        Const VK_A = &H41 '[A]
        'Const VK_B = &H42 '[B]
        'Const VK_C = &H43 '[C]
        'Const VK_D = &H44 '[D]
        'Const VK_E = &H45 '[E]
        'Const VK_F = &H46 '[F]
        'Const VK_G = &H47 '[G]
        'Const VK_H = &H48 '[H]
        Const VK_I = &H49 '[I]
        Const VK_J = &H4A '[J]
        'Const VK_K = &H4B '[K]
        Const VK_L = &H4C '[L]
        Const VK_M = &H4D '[M]
        'Const VK_N = &H4E '[N]
        'Const VK_O = &H4F '[O]
        'Const VK_P = &H50 '[P]
        'Const VK_Q = &H51 '[Q]
        'Const VK_R = &H52 '[R]
        Const VK_S = &H53 '[S]
        'Const VK_T = &H54 '[T]
        'Const VK_U = &H55 '[U]
        'Const VK_V = &H56 '[V]
        'Const VK_W = &H57 '[W]
        Const VK_X = &H58 '[X]
        'Const VK_Y = &H59 '[Y]
        Const VK_Z = &H5A '[Z]
        'Const VK_NUMPAD0 = &H60 'テンキー[0]
        'Const VK_NUMPAD1 = &H61 'テンキー[1]
        'Const VK_NUMPAD2 = &H62 'テンキー[2]
        'Const VK_NUMPAD3 = &H63 'テンキー[3]
        'Const VK_NUMPAD4 = &H64 'テンキー[4]
        'Const VK_NUMPAD5 = &H65 'テンキー[5]
        'Const VK_NUMPAD6 = &H66 'テンキー[6]
        'Const VK_NUMPAD7 = &H67 'テンキー[7]
        'Const VK_NUMPAD8 = &H68 'テンキー[8]
        'Const VK_NUMPAD9 = &H69 'テンキー[9]
        'Const VK_MULTIPLY = &H6A 'テンキー[*]
        'Const VK_ADD = &H6B 'テンキー[+]
        'Const VK_SEPARATOR = &H6C 'テンキー[Enter]
        'Const VK_SUBTRACT = &H6D 'テンキー[-]
        'Const VK_DECIMAL = &H6E 'テンキー[.]
        'Const VK_DIVIDE = &H6F 'テンキー[/]
        'Const VK_F1 = &H70 '[F1]
        'Const VK_F2 = &H71 '[F2]
        'Const VK_F3 = &H72 '[F3]
        'Const VK_F4 = &H73 '[F4]
        'Const VK_F5 = &H74 '[F5]
        'Const VK_F6 = &H75 '[F6]
        'Const VK_F7 = &H76 '[F7]
        'Const VK_F8 = &H77 '[F8]
        'Const VK_F9 = &H78 '[F9]
        'Const VK_F10 = &H79 '[F10]
        'Const VK_F11 = &H7A '[F11]
        'Const VK_F12 = &H7B '[F12]
        'Const VK_F13 = &H7C '[F13]
        'Const VK_F14 = &H7D '[F14]
        'Const VK_F15 = &H7E '[F15]
        'Const VK_F16 = &H7F '[F16]
        'Const VK_F17 = &H80 '[F17]
        'Const VK_F18 = &H81 '[F18]
        'Const VK_F19 = &H82 '[F19]
        'Const VK_F20 = &H83 '[F20]
        'Const VK_F21 = &H84 '[F21]
        'Const VK_F22 = &H85 '[F22]
        'Const VK_F23 = &H86 '[F23]
        'Const VK_F24 = &H87 '[F24]
        'Const VK_NUMLOCK = &H90 '[Num Lock]
        'Const VK_WIN = &H5B '[Win]


        Const yajirushitime = 300
        Const buttontime = 350
        ' AppActivate("bgb")

            System.Windows.Forms.Application.DoEvents()
            sentaku = "0"



            If you = 0 Then
                sentaku = "←"
            System.Windows.Forms.Application.DoEvents()

            Call keybd_event(VK_J, 0, 0, 0) '｢M｣キーを押して
            ' AppActivate("bgb")
            System.Windows.Forms.Application.DoEvents()
            System.Threading.Thread.Sleep(yajirushitime)
                Call keybd_event(VK_J, 0, 2, 0) '｢M｣キーを放す
            System.Windows.Forms.Application.DoEvents()
            End If
            If you = 1 Then
                sentaku = "→"
            System.Windows.Forms.Application.DoEvents()

                Call keybd_event(VK_L, 0, 0, 0) '｢M｣キーを押して

                System.Threading.Thread.Sleep(yajirushitime)
                Call keybd_event(VK_L, 0, 2, 0) '｢M｣キーを放す
            System.Windows.Forms.Application.DoEvents()
            End If
            If you = 2 Then
                sentaku = "↑"
            System.Windows.Forms.Application.DoEvents()

                Call keybd_event(VK_I, 0, 0, 0) '｢M｣キーを押して
            System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(yajirushitime)
                Call keybd_event(VK_I, 0, 2, 0) '｢M｣キーを放す
            System.Windows.Forms.Application.DoEvents()
            End If
            If you = 3 Then
                sentaku = "↓"
            System.Windows.Forms.Application.DoEvents()

                Call keybd_event(VK_M, 0, 0, 0) '｢M｣キーを押して
                System.Threading.Thread.Sleep(yajirushitime)
            System.Windows.Forms.Application.DoEvents()
                Call keybd_event(VK_M, 0, 2, 0) '｢M｣キーを放す
            System.Windows.Forms.Application.DoEvents()
            End If
            If you = 4 Then
                sentaku = "A"
       
                     Call keybd_event(VK_S, 0, 0, 0) '｢M｣キーを押して
                System.Threading.Thread.Sleep(buttontime)
            System.Windows.Forms.Application.DoEvents()
                Call keybd_event(VK_S, 0, 2, 0) '｢M｣キーを放す
            System.Windows.Forms.Application.DoEvents()
            End If
            If you = 5 Then
                sentaku = "B"
            System.Windows.Forms.Application.DoEvents()

                Call keybd_event(VK_A, 0, 0, 0) '｢M｣キーを押して
            System.Windows.Forms.Application.DoEvents()
                Call keybd_event(VK_A, 0, 2, 0) '｢M｣キーを放す
                System.Threading.Thread.Sleep(buttontime)
            System.Windows.Forms.Application.DoEvents()
            End If
            If you = 6 Then
                sentaku = "start"
            System.Windows.Forms.Application.DoEvents()

                Call keybd_event(VK_X, 0, 0, 0) '｢M｣キーを押して
                System.Threading.Thread.Sleep(buttontime)
            System.Windows.Forms.Application.DoEvents()
                Call keybd_event(VK_X, 0, 2, 0) '｢M｣キーを放す
            System.Windows.Forms.Application.DoEvents()
            End If
            If you = 7 Then
                sentaku = "select"
            System.Windows.Forms.Application.DoEvents()

                Call keybd_event(VK_Z, 0, 0, 0) '｢M｣キーを押して
                System.Threading.Thread.Sleep(buttontime)
            System.Windows.Forms.Application.DoEvents()
                Call keybd_event(VK_Z, 0, 2, 0) '｢M｣キーを放す
            System.Windows.Forms.Application.DoEvents()
            End If


            If you = 8 Then
                Dim o1 As New System.Random(1000)
                Dim i1 As Integer = o1.Next(7)
                sentaku = sentaku(i1)
            End If



             System.Windows.Forms.Application.DoEvents()




    End Function


    Sub winwin()
        Dim WSH As Object
        WSH = CreateObject("WScript.Shell")
        WSH.Popup("閉じた60秒後に動作を開始します。", 2, "プログラムの歩調を調整しています。", vbInformation)
        WSH = Nothing
    End Sub
    Function getlv() As String

        Dim OKflag As Integer = 0
        Dim aaa(0) As Integer
        Dim aaaaaa As String = Nothing

        Dim log As CookieContainer = Login("ステアカウントメアド", "パスワード", 190823874)

        While OKflag = 0

            System.Threading.Thread.Sleep(30000)



            Dim textman As String = Read(log, "http://com.nicovideo.jp/community/co2497791")
            Dim lvreg As New System.Text.RegularExpressions.Regex("<h2><a href=""http://live.nicovideo.jp/watch/lv[0-9]+", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            Dim commander As System.Text.RegularExpressions.Match = lvreg.Match(textman)
         


            While commander.Success

                '一致した対象が見つかったときキャプチャした部分文字列を表示 
                aaaaaa = commander.Value

                commander = commander.NextMatch()

            End While


            System.Windows.Forms.Application.DoEvents()

            If Not aaaaaa = Nothing Then
                OKflag = 1

            End If

            aaaaaa = Replace(aaaaaa, "<h2><a href=""http://live.nicovideo.jp/watch/lv", "")


        End While


        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim tokenhtml As String = Read(tokenlog, "http://live.nicovideo.jp/api/getpublishstatus/lv" + lv)

        Dim tokenreg As New System.Text.RegularExpressions.Regex("<token>[0-9a-z]+", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        Dim tokenbox As System.Text.RegularExpressions.Match = tokenreg.Match(tokenhtml)

        While tokenbox.Success

            '一致した対象が見つかったときキャプチャした部分文字列を表示 
            token = tokenbox.Value

            tokenbox = tokenbox.NextMatch()

        End While
        token = Replace(token, "<token>", "")
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        getlv = aaaaaa


    End Function


    Sub savepoint()

        Const VK_F2 = &H71 '[F2]

        System.Windows.Forms.Application.DoEvents()
        Call keybd_event(VK_F2, 0, 0, 0) '｢M｣キーを押して
        System.Threading.Thread.Sleep(350)
        Call keybd_event(VK_F2, 0, 2, 0) '｢M｣キーを放す
        System.Windows.Forms.Application.DoEvents()



    End Sub

    '
    Public Function UrlEncodeUtf8(ByRef strSource As String) As String
        Dim objSC As Object
        objSC = CreateObject("ScriptControl")
        objSC.Language = "Jscript"
        UrlEncodeUtf8 = objSC.CodeObject.encodeURIComponent(strSource)
        objSC = Nothing
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub
End Class

