Imports System.IO
Imports System.Text

Module CreDAT


#Region "共通ヘッダ作成"


    ''' <summary>
    ''' 共通ヘッダ作成
    ''' </summary>
    ''' <param name="RecNo" >レコードNo.</param>
    ''' <param name="series" >シリーズNo.</param>
    ''' <remarks></remarks>
    Function CreCommonRec(ByVal RecNo As String, ByVal series As String) As String

        '共通ヘッダ作成
        Dim str As New StringBuilder(18)    '戻り値



        'ジョブNo.
        str.Append(Job)

        'レコードNo.
        str.Append(RecNo)

        'シリーズNo.
        str.Append(CreSpace(ComDic("シリーズNo") - LenB(series), "0"))
        str.Append(series)

        '空白
        str.Append(CreSpace(ComDic("空白"), " "))

        '子番号=群番
        str.Append(Gunban)
        str.Append(CreSpace(ComDic("子番号") - LenB(Gunban), " "))

        '盤No
        str.Append(BanNo)
        str.Append(CreSpace(ComDic("盤No") - LenB(BanNo), " "))

        CreCommonRec = str.ToString
    End Function

#End Region


#Region "Create Record1"

    ''' <summary>
    ''' レコード1作成
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreRec1()

        Dim Line As New StringBuilder

        Dim Coordcnt As String = CStr(ArrCoord.Count - 1)     '座標数


        'Kouban = TrimOrder(Dic1("工番"))        '工番


        '共通ヘッダ
        Line.Append(CreCommonRec(CStr(1), CStr(1)))

        '前背
        Line.Append(FB)

        'SCT
        Line.Append(CreSpace(RecDic1("SCT"), " "))

        'ゾーン
        Line.Append(CreSpace(RecDic1("ゾーン"), " "))

        '工番
        Line.Append(Kouban)
        Line.Append(CreSpace(RecDic1("工番") - LenB(Kouban), " "))

        '?
        Line.Append(CreSpace(RecDic1("?"), " "))

        '空白2
        Line.Append(CreSpace(RecDic1("空白2"), " "))

        '客先名


        Line.Append(CustomerName)
        Line.Append(CreSpace(RecDic1("客先名") - LenB(CustomerName), " "))

        '客先コード
        Line.Append(CreSpace(RecDic1("客先コード"), " "))

        '空白3
        Line.Append(CreSpace(RecDic1("空白3"), " "))

        '座標数
        Dim spacecnt = RecDic1("座標数") - Coordcnt.Length

        If spacecnt <> 0 Then
            '0つめ
            Line.Append(CreSpace(spacecnt, "0"))
        End If

        Line.Append(Coordcnt)

        'マークバンド
        Line.Append(MarkBand)

        '空白4
        Line.Append(CreSpace(RecDic1("空白4"), " "))

        '結果をPublicに保存
        Rec1 = Line.ToString

    End Sub

#End Region


#Region "Create Record2"


    Public Sub CreRec2()
        '切り捨て
        Dim max As Integer = CInt(Math.Floor((CSVDataDic.Count - 1) / 3)) - 1
        Dim Line As New StringBuilder



        For i As Integer = 0 To max
            '共通ヘッダ
            Line.Append(CreCommonRec(CStr(2), CStr(i + 1)))


            '前背
            Line.Append(FB)

            'SCT
            Line.Append(CreSpace(CInt(RecDic2("SCT")), " "))

            'ゾーン
            Line.Append(CreSpace(CInt(RecDic2("ゾーン")), " "))

            '空白2
            Line.Append(CreSpace(CInt(RecDic2("空白2")), " "))

            For j As Integer = 1 To 3
                SupRec2(i * 3 + j, Line)
            Next j

            '空白3
            Line.Append(CreSpace(CInt(RecDic2("空白3")), " "))
            '結果をRec2に保存
            Rec2.Add(Line.ToString)
            '初期化
            Line.Clear()
        Next i

        '3で割って余りが存在する場合
        If CSVDataDic.Count Mod 3 <> 0 Then

            For i As Integer = Rec2.Count * 3 To CSVDataDic.Count
                If i = Rec2.Count * 3 Then
                    '共通ヘッダ
                    Line.Append(CreCommonRec(CStr(2), CStr(Rec2.Count + 1)))

                    '前背
                    Line.Append(FB)

                    'SCT
                    Line.Append(CreSpace(CInt(RecDic2("SCT")), " "))

                    'ゾーン
                    Line.Append(CreSpace(CInt(RecDic2("ゾーン")), " "))

                    '空白2
                    Line.Append(CreSpace(CInt(RecDic2("空白2")), " "))
                End If

                SupRec2(i, Line)

            Next

            ''データ数が1,2の場合、空白を追加(第1引数が0の場合、CreSpaceは何もせず終了)
            'Line.Append(CreSpace(Set2Area * (3 - (CSVDataDic.Count - Rec2.Count * 3)), " "))

            ''空白3
            'Line.Append(CreSpace(CInt(RecDic2("空白3")), " "))

            Line.Append(CreSpace(RecordLength - Line.Length, " "))

            Rec2.Add(Line.ToString)
            '初期化
            Line.Clear()
        End If '(ArrCoord.Length - 1) Mod 3 <> 0
    End Sub

    Sub SupRec2(ByVal index As Integer, ByRef Line As StringBuilder)
        Dim coord As String = ""
        Dim data() As String

        coord = CSVDataDic.Keys(index - 1)
        '0=略称、1=形式、2=端子種類
        data = CSVDataDic(coord)

        '座標名+分割名
        Line.Append(coord)

        '略称

        OverSpace(data(0), SetDic2("略称"), Line)


        '形式
        OverSpace(data(1), SetDic2("形式"), Line)

        'フレーム
        Line.Append(CreSpace(SetDic2("フレーム"), " "))

        '空白
        Line.Append(CreSpace(SetDic2("空白"), " "))

        '端子種類
        OverSpace(data(2), SetDic2("端子種類"), Line)
    End Sub

    Public Sub OverSpace(ByVal data As String, ByVal Mark As Integer, ByRef Line As StringBuilder)
        Dim str As String = ""

        If LenB(data) < Mark Then
            str = data
            Line.Append(data)
            '空白詰め

        Else
            str = MidB(data, 1, Mark)
            Line.Append(str)

        End If

        Line.Append(CreSpace(Mark - LenB(str), " "))
    End Sub

#End Region


#Region "Create Record3"


    Sub CreRec3()

        For i As Integer = 0 To SortedRec3_4_Data.Keys.Count - 1
            '共通ヘッダの作成
            If Rec3Cnt = 0 Then

                '共通ヘッダ
                Line3.Append(CreCommonRec(CStr(3), CStr(Row3Cnt)))

                '前背
                Line3.Append(FB)

                'SCT
                Line3.Append(CreSpace(CInt(RecDic3("SCT")), " "))

                'ゾーン
                Line3.Append(CreSpace(CInt(RecDic3("ゾーン")), " "))

                '空白2
                Line3.Append(CreSpace(CInt(RecDic3("空白2")), " "))

                '行数カウントアップ
                Row3Cnt = Row3Cnt + 1
            End If

            Line3.Append(SortedRec3_4_Data.Keys(i))
            Rec3Cnt += 1

            If Rec3Cnt = 10 Then
                 '10データまでデータ入力できたとき
                'レコード長を揃えるための空白を用意
                Line3.Append(CreSpace(CInt(RecDic3("空白3")), " "))
                'データを追加
                Rec3.Add(Line3.ToString)
                Line3.Clear()
                Rec3Cnt = 0
            ElseIf i = SortedRec3_4_Data.Keys.Count - 1 Then
                '10データまでの入力ができていない状態で、最終行を迎えたとき
                '不足分を空白で埋める
                Line3.Append(CreSpace((10 - Rec3Cnt) * 8, " "))
                'レコード長を揃えるための空白を用意
                Line3.Append(CreSpace(CInt(RecDic3("空白3")), " "))
                'データを追加
                Rec3.Add(Line3.ToString)
                Line3.Clear()
                Rec3Cnt = 0
            End If

        Next i

    End Sub

#End Region


#Region "Create Record4"



    Sub CreRec4()
        Dim Key As String
        Dim Value As List(Of String)
        Dim str As String
        Dim FromStr As String = ""
        Dim LineSign As String = ""
        Dim ToStr As String = ""

        CreRec4Header()

        State4 = "To"

        '最初の0詰め
        Line4.Append(CreSpace(Set4Area, "0"))

        State4 = "From"
        Rec4Cnt += 1

        For i As Integer = 0 To SortedRec3_4_Data.Keys.Count - 1
            'Keyの数だけ繰り返す
            Key = SortedRec3_4_Data.Keys(i)
            Value = SortedRec3_4_Data(Key)

            For j As Integer = 0 To Value.Count - 1
                str = Value(j)                                          'From+線符号+To

                FromStr = Mid(str, 1, Set4Area)                         'From
                LineSign = Mid(str, Set4Area + 1, RecDic4("線符号1"))   '線符号
                ToStr = Mid(str, Set4Area + RecDic4("線符号1") + 1)     'To

SelectCase:

                Select Case State4

                    Case "From"
                        Line4.Append(FromStr)
                        Line4.Append(LineSign)
                        State4 = "To"
                        Rec4Cnt += 1

                    Case "To"
                        Line4.Append(ToStr)
                        State4 = "From"
                        Rec4Cnt += 1
                End Select

                '一行分のデータが入力されていれば、次の行へ
                If Rec4Cnt = 6 Then NextRow()

                'state4="To"のときはToのデータ入力をおこなっていないのでSelectCaseまで戻る
                If State4 = "To" Then GoTo SelectCase
            Next j


            If i = SortedRec3_4_Data.Keys.Count - 1 Then
                'データ入力が最後までおこなっている場合

                'To分だけ0詰め
                Line4.Append(CreSpace(Set4Area, "0"))

                'レコード長になるまで空白詰め
                Line4.Append(CreSpace(RecordLength - Line4.Length, " "))

                Rec4.Add(Line4.ToString)    'データ追加
                Line4.Clear()               'データクリア
            Else
                'データ入力が最後までおこなっていない場合
                '線種、線サイズ、色が変わる
                'From分だけ0詰め
                Line4.Append(CreSpace(Set4Area, "0"))
                '線符号を記入
                Line4.Append(LineSign)
                Rec4Cnt += 1

                'レコードが最後まで行っていれば次の行に移動
                If Rec4Cnt = 6 Then NextRow()

                'To分だけ0詰め
                Line4.Append(CreSpace(Set4Area, "0"))

                Rec4Cnt += 1

                'レコードが最後まで行っていれば次の行に移動
                If Rec4Cnt = 6 Then NextRow()
            End If
        Next i


    End Sub 'CreRec4

    ''' <summary>
    ''' 次の行へ移行する
    ''' </summary>
    ''' <remarks></remarks>
    Sub NextRow()

        Rec4Cnt = 0

        Rec4.Add(Line4.ToString)

        Line4.Clear()
        CreRec4Header()
    End Sub


    Sub CreRec4Header()

        '共通ヘッダ
        Line4.Append(CreCommonRec(CStr(4), CStr(Row4Cnt)))

        '前背
        Line4.Append(FB)

        'SCT
        Line4.Append(CreSpace(CInt(RecDic4("SCT")), " "))

        'ゾーン
        Line4.Append(CreSpace(CInt(RecDic4("ゾーン")), " "))

        '空白2
        Line4.Append(CreSpace(CInt(RecDic4("空白2")), " "))

        '行数カウントアップ
        Row4Cnt = Row4Cnt + 1

    End Sub




#End Region


#Region "Create Record3 and 4 Data"

    Sub CreRec3_4_Data()

        Dim StrB As New StringBuilder
        Dim Key As String
        Dim Value As New StringBuilder


        'SQCord,YoriCord,ColorCordについてはGlobal.vb参照
        'SQコード
        If SQCord.ContainsKey(SQ.Trim) = True Then
            'SQがGlobal.vbで定義されているDictionary:SQCordのKeyに存在する場合
            StrB.Append(SQCord(SQ.Trim))
        Else
            'SQがKeyに存在しない場合
            StrB.Append(CreSpace(2, " "))
        End If

        '肉厚コード
        StrB.Append(CreSpace(2, " "))  '空白

        '撚りコード
        If YoriCord.ContainsKey(LnKnd.Trim) = True Then
            'LnKndがKeyに存在する場合
            StrB.Append(SQCord(LnKnd.Trim))
        Else
            'LnKndがKeyに存在しない場合
            StrB.Append(CreSpace(2, " "))
        End If

        '色コード
        If ColorCord.ContainsKey(Color.Trim) = True Then
            'ColorがKeyに存在する場合
            StrB.Append(ColorCord(Color.Trim))
        Else
            'ColorがKeyに存在しない場合、直接入力
            StrB.Append(Color)
        End If

        Key = StrB.ToString

        'Public変数に値を導入
        SetPubFT()

        '----------------------------From----------------------------------
        '0詰め(1 then 001)
        Value.Append(CreSpace(SetDic4("仮座標") - LenB(FromTemp), "0"))
        '仮座標
        Value.Append(FromTemp)
        'カラー
        Value.Append(FromClr)
        '空白
        Value.Append(CreSpace(SetDic4("空白"), " "))
        '端子番号
        Value.Append(FromPin)
        'ベンチNo
        Value.Append(CreSpace(SetDic4("ベンチNo"), " "))

        '----------------------------From----------------------------------

        '線符号
        Value.Append(LFSign)

        '-----------------------------To-----------------------------------
        '0詰め
        Value.Append(CreSpace(SetDic4("仮座標") - LenB(ToTemp), "0"))

        '仮座標
        Value.Append(ToTemp)
        'カラー
        Value.Append(ToClr)
        '空白
        Value.Append(CreSpace(SetDic4("空白"), " "))
        '端子番号
        Value.Append(ToPin)
        'ベンチNo
        Value.Append(CreSpace(SetDic4("ベンチNo"), " "))
        '-----------------------------To-----------------------------------

        If Rec3_4_Data.ContainsKey(Key) = False Then
            'KeyがDictionaryに存在しない場合
            'データを追加
            Rec3_4_Data.Add(Key, New List(Of String)(New String() {Value.ToString}))
        Else
            'KeyがDictionaryに存在しない場合
            '既存のValue(List)を取得し、そのデータに今回をのデータを追加する。

            Dim temp As List(Of String) = Rec3_4_Data(Key)

            If temp.Contains(Value.ToString) = False Then
                temp.Add(Value.ToString)
                Rec3_4_Data(Key) = temp
            End If

        End If

    End Sub
#End Region


#Region "Create Record5"

    ''' <summary>
    ''' レコードNo.5作成
    ''' </summary>
    ''' <remarks></remarks>
    Sub CreRec5()
        Dim max As Integer = CInt(Math.Floor(Seq.Count / 10) - 1)
        Dim Line As New StringBuilder
        Dim str As String
        Dim place As Integer
        Dim i As Integer, j As Integer

        For i = 0 To max
            '共通ヘッダ
            Line.Append(CreCommonRec(CStr(5), CStr(i + 1)))

            '前背
            Line.Append(FB)

            '9999
            Line.Append("9999")


            '空白2
            Line.Append(CreSpace(CInt(RecDic5("空白2")), " "))

            For j = 0 To 9
                str = Seq(i * 10 + j)

                Line.Append(str)
                Line.Append(CreSpace(10 - LenB(str), " "))
            Next j

            If i = max Then
                'jが9まではFor文が動くので、最後にカウントアップされてj=10
                place = i * 10 + j
            End If

            '空白3
            Line.Append(CreSpace(CInt(RecDic5("空白3")), " "))
            '結果をRec5に保存
            Rec5.Add(Line.ToString)
            '初期化
            Line.Clear()
        Next i

        If place <> Seq.Count Then
            'データ数が10で割り切れないとき
            '共通ヘッダ
            Line.Append(CreCommonRec(CStr(5), CStr(i + 1)))

            '前背
            Line.Append(FB)

            '9999
            Line.Append("9999")


            '空白2
            Line.Append(CreSpace(CInt(RecDic5("空白2")), " "))

            For i = place + 1 To Seq.Count


                str = Seq(i - 1)
                Line.Append(str)
                Line.Append(CreSpace(10 - LenB(str), " "))
            Next

            Line.Append(CreSpace((10 - (Seq.Count - place)) * 10, " "))

            '空白3
            Line.Append(CreSpace(CInt(RecDic5("空白3")), " "))
            '結果をRec5に保存
            Rec5.Add(Line.ToString)
            '初期化
            Line.Clear()
        End If
    End Sub

#End Region


#Region "Create Record6"


    Sub CreRec6_Data()
        Dim Key As String
        Dim StrB As New StringBuilder
        Dim Value As New StringBuilder
        Dim LineKind As String = ConvertStr(Trim(LnKnd), ConvLnKnd)
        Dim Square As String = ConvertStr(Trim(SQ), ConvSQ)

        '線種、線サイズ、線色のKeyを作成
        '電線種
        StrB.Append(LineKind)
        '空白詰め
        StrB.Append(CreSpace(RecDic6("電線種") - LenB(LineKind), " "))
        '太さ
        StrB.Append(Square)
        '空白詰め
        StrB.Append(CreSpace(RecDic6("太さ") - LenB(Square), " "))
        '線色
        StrB.Append(Color)
        '空白詰め
        StrB.Append(CreSpace(RecDic6("線色") - LenB(Color), " "))

        '変数Keyに代入
        Key = StrB.ToString


        SetPubFT()

        '----------------------------From----------------------------------
        '0詰め(1 then 001)
        Value.Append(CreSpace(SetDic6("仮座標") - LenB(FromTemp), "0"))
        '仮座標
        Value.Append(FromTemp)
        '*
        Value.Append(CreSpace(SetDic6("**"), "*"))
        '端子番号
        Value.Append(FromPin)

        '----------------------------From----------------------------------

        '-----------------------------To-----------------------------------
        '0詰め
        Value.Append(CreSpace(SetDic6("仮座標") - LenB(ToTemp), "0"))

        '仮座標
        Value.Append(ToTemp)

        '*
        Value.Append(CreSpace(SetDic6("**"), "*"))
        '端子番号
        Value.Append(ToPin)
        '-----------------------------To-----------------------------------

        '線符号
        Value.Append(LFSign)
        'カラー
        Value.Append(FromClr)


        If Rec6_Data.ContainsKey(Key) = False Then
            'KeyがDictionaryに存在しない場合
            'データを追加
            Rec6_Data.Add(Key, New List(Of String)(New String() {Value.ToString}))
        Else
            'KeyがDictionaryに存在する場合
            '既存のValue(List)を取得し、そのデータに今回のデータを追加する。

            Dim temp As List(Of String) = Rec6_Data(Key)

            If temp.Contains(Value.ToString) = False Then
                temp.Add(Value.ToString)
                Rec6_Data(Key) = temp
            End If

        End If

    End Sub


    Sub CreRec6()

        Dim Key As String = ""
        Dim Value As New List(Of String)

        For i As Integer = 0 To SortedRec6_Data.Keys.Count - 1

            Key = SortedRec6_Data.Keys(i)
            Value = SortedRec6_Data(Key)

            For j As Integer = 0 To Value.Count - 1

                            '共通ヘッダの作成
                If Rec6Cnt = 0 Then

                    '共通ヘッダ
                    Line6.Append(CreCommonRec(CStr(6), CStr(Row6Cnt)))

                    '前背
                    Line6.Append(FB)

                    'SCT
                    Line6.Append(CreSpace(CInt(RecDic6("SCT")), " "))

                    'ゾーン
                    Line6.Append(CreSpace(CInt(RecDic6("ゾーン")), " "))

                    '空白2
                    Line6.Append(CreSpace(CInt(RecDic6("空白2")), " "))


                    Line6.Append(Key)

                    '行数カウントアップ
                    Row6Cnt = Row6Cnt + 1
                End If

                Line6.Append(Value(j))
                Rec6Cnt += 1

                If Rec6Cnt = 3 Then
                     '10データまでデータ入力できたとき
                    'レコード長を揃えるための空白を用意
                    Line6.Append(CreSpace(RecordLength - Line6.Length, " "))
                    'データを追加
                    Rec6.Add(Line6.ToString)
                    Line6.Clear()
                    Rec6Cnt = 0
                ElseIf j = Value.Count - 1 Then
                    '10データまでの入力ができていない状態で、最終行を迎えたとき
                    Line6.Append(CreSpace(RecordLength - LenB(Line6.ToString), " "))
                    'データを追加
                    Rec6.Add(Line6.ToString)
                    Line6.Clear()
                    Rec6Cnt = 0
                End If

            Next j
        Next i
    End Sub
#End Region


#Region "Create Record7"

    Public Sub CreRec7_Data()

        '{略称, 形式, 端子種類, 器具番号}
        Dim FromArr() As String = SearchCSVDataDic(FromCoord)
        Dim ToArr() As String = SearchCSVDataDic(ToCoord)

        'From端末の略称がTBの時
        If Left(FromArr(0), 2) = "TB" Then
           '略称頭二文字がTB=端子台のとき(From)

           InsertData(FromCoord, FromArr, FromPin)
        End If

        'To端末の略称がTBの時
        If Left(ToArr(0), 2) = "TB" Then
            '略称頭二文字がTB=端子台のとき(To)

            InsertData(ToCoord, ToArr, ToPin)
        End If
    End Sub

    Sub InsertData(ByVal Coord As String, ByVal Arr() As String, ByVal Pin As String)
        Dim TempArr(2) As String        '追加する配列データ
        Dim Seach(2) As String          '同じのがあるかどうかを探すために使用
        Dim Ident As Boolean = False    '同一データがあるかどうかのフラグ

        'マークバンド、形式、座標名、端子番号、線符号
        TempArr = {MarkBand, Arr(1), Coord, Pin, LFSign}


        'データがあれば
        If Rec7_DataList.Count <> 0 Then
            For i As Integer = 0 To Rec7_DataList.Count - 1
                Seach = Rec7_DataList(i)
                If Seach(1) = TempArr(1) And _
                   Seach(2) = TempArr(2) And _
                   Seach(3) = TempArr(3) And _
                   Seach(4) = TempArr(4) Then

                    'まったく同じデータが見つかればフラグON
                   Ident = True
                   Exit For
                End If
            Next

            If Ident = False Then
                '同一データがなければ、データ追加
                Rec7_DataList.AddLast(TempArr)
            End If
        Else
            'データがなければ、データ追加
            Rec7_DataList.AddLast(TempArr)
        End If



    End Sub

    Public Function SearchCSVDataDic(ByVal CoordiNate As String) As String()
        Dim key As String
        Dim temp As String
        Dim TrimCoord As String

        For i As Integer = 0 To CSVDataDic.Count - 1
            key = CSVDataDic.Keys(i)
            temp = Replace(key, " ", "")
            TrimCoord = CoordiNate.Trim
            If temp = TrimCoord Then
                SearchCSVDataDic = CSVDataDic(key)
                Exit Function
            End If
        Next

        Using Err As New StreamWriter("C:\Users\94503\Desktop\err.txt", True, Encoding.Default)
            Err.WriteLine(CoordiNate)
        End Using

        Return {"    ", "            ", "  ", "          "}
    End Function


    Sub CreRec7()
        Dim arr() As String
        Dim Mark As String          'マークバンド
        Dim Format As String        '形式
        Dim Coord As String         '座標
        Dim LSign As String         '線符号
        Dim Term As String          '端子No
        Dim Moji As String = ""     '文字板
        Dim Max As Integer

        Line7.Clear()

        'LinkedListをArrayListに変換
        Dim lst As New List(Of String())(Rec7_DataList)
        lst.Sort(AddressOf Sort.CompareArrStr)

        For i As Integer = 0 To Rec7_DataList.Count - 1


            'マークバンド、形式、座標名、線符号
            arr = Rec7_DataList(i)
            'マークバンド、形式
            Mark = arr(0) : Format = arr(1)
            '座標、線符号
            Coord = arr(2) : LSign = arr(3)

            '形式からループ数を求める
            Max = FindCount(arr(1))

            For j As Integer = 1 To Max

                If Rec7Cnt = 0 Then
                    '共通ヘッダ
                    Line7.Append(CreCommonRec(7, Row7Cnt))
                    '前背
                    Line7.Append(FB)
                    '9999
                    Line7.Append(CreSpace(RecDic7("9999"), "9"))
                    '空白2
                    Line7.Append(CreSpace(RecDic7("空白2"), " "))
                    '形式
                    Line7.Append(Format)
                    '空白埋め
                    Line7.Append(CreSpace(RecDic7("形式") - LenB(Format), " "))

                    '座標
                    If LenB(Coord) < RecDic7("座標") Then
                        '4Byte以下なら
                        Line7.Append(Coord)
                        Line7.Append(CreSpace(RecDic7("座標") - LenB(Coord), " "))
                    Else
                        '4Byte以上なら
                        Line7.Append(MidB(Coord, 1, RecDic7("座標")))
                    End If

                    '行数カウントアップ
                    Row7Cnt += 1
                End If

                Term = CStr(j)  '端子No.
                Line7.Append(CreSpace(SetDic7("端子No") - LenB(Term), "0"))     '0詰め
                Line7.Append(Term)  '端子No.

                Select Case arr(0)
                    'HOTなら端子番号、FMKなら線符号を文字板に出力
                    Case "HOT"
                        Moji = Term
                    Case ("FMK")
                        Moji = LSign
                End Select
                '文字板
                Line7.Append(Moji)
                '空白詰め
                Line7.Append(CreSpace(SetDic7("文字板") - LenB(Moji), " "))

                Rec7Cnt += 1

                If Rec7Cnt = 6 Then
                    Line7.Append(CreSpace(RecDic7("空白3"), " "))
                    Rec7Cnt = 0
                    Rec7.Add(Line7.ToString)
                    Line7.Clear()
                ElseIf j = Max Then
                    Line7.Append(CreSpace(RecordLength - LenB(Line7.ToString), " "))
                    Rec7Cnt = 0
                    Rec7.Add(Line7.ToString)
                    Line7.Clear()
                End If
            Next j
        Next i
    End Sub

    ''' <summary>
    ''' 形式に含まれる数字を見つけて、それを返す
    ''' </summary>
    ''' <param name="str">形式</param>
    ''' <remarks></remarks>
    Function FindCount(ByVal str As String) As Integer
        Dim temp1 As String = ""
        Dim temp2 As String = ""
        Dim count As Integer = 0

        For i = str.Length - 1 To 0 Step -1
            temp1 = str.Substring(i, 1)

            If IsNumeric(temp1) = True Then
                '数字の時
                If i <> 0 Then
                    '数字を見つけた位置が文字列の初めでないとき
                    temp2 = str.Substring(i - 1, 1)
                    If IsNumeric(temp2) = True Then
                        '数字の←隣が数字のとき
                        count = CInt(temp2 & temp1)
                        If count <= MojiBaseMax Then
                            'count=24以下のとき
                            Return count
                        Else
                            'count=24超過のとき
                            Return CInt(temp1)
                        End If
                    Else
                        '数字の←隣が数字ではないとき
                        Return temp1
                    End If
                Else
                    '数字を見つけた位置が文字列の初めのとき
                    Return temp1
                End If
            End If
        Next

        '数字が一つも見つからなかった場合
        Return 0

    End Function
#End Region


#Region "Create Record9"


    Public Sub CreRec9()
        '線サイズ
        Dim TrimSQ As String = SQ.Trim
        '線種
        Dim LK As String = Left(LnKnd.Trim, 2)

        Select Case LK

            Case "1C", "2C", "Fu"
                '1芯、2芯シールド、付属線の場合
                CountRec9(LK) = CountRec9(LK) + 1
            Case Else
                Select Case TrimSQ
                    Case "0.3", "0.5", "1.25", "2.0", _
                        "3.5", "5.5", "8", "14", "22", _
                        "38", "60", "100", "150"
                        '線サイズが上記の場合
                        CountRec9(TrimSQ) = CountRec9(TrimSQ) + 1

                    Case Else
                        '線サイズが規定以外の場合
                        CountRec9("ETC") = CountRec9("ETC") + 1

                End Select
        End Select

        '最終行の時
        If LastFlg = True Then
            'カウント結果を文字列にして保存する

            Dim Line As New StringBuilder
            Dim key As String = ""

            '共通ヘッダ
            Line.Append(CreCommonRec(CStr(9), CStr(Rec9Cnt)))
            'シリーズNo.カウントアップ
            Rec9Cnt = Rec9Cnt + 1

            '9詰
            Line.Append(CreSpace(RecDic9("9詰"), "9"))

            'FB
            Line.Append(FB)

            '空白2
            Line.Append(CreSpace(RecDic9("空白2"), " "))

            '0.3-付属線まで
            For i = 0 To CountRec9.Keys.Count - 1
                key = CountRec9.Keys(i)

                '空白詰
                Line.Append(CreSpace(RecDic9(key) - LenB(CStr(CountRec9(key))), " "))
                '0.3
                Line.Append(CountRec9(key))
            Next



            '空白3
            Line.Append(CreSpace(RecDic9("空白3") - 1, " "))
            Line.Append("*")

            '空白4
            Line.Append(CreSpace(RecDic9("空白4"), " "))

            'データ追加
            Rec9.Add(Line.ToString)

        End If
    End Sub

    Sub ClearRec9()
        Dim key As String
        For i As Integer = 0 To CountRec9.Count - 1
            key = CountRec9.Keys(i)
            CountRec9(key) = 0
        Next

    End Sub
#End Region

''' <summary>
''' 
''' </summary>
''' <param name="key">指定されたキー</param>
''' <param name="Dic">キーがあるかどうかを調べるDictionary</param>
''' <returns>キーが存在すれば、キーに対応した値。キーが存在しなければキーを返す</returns>
''' <remarks></remarks>

Public Function ConvertStr(ByVal key As String, ByRef Dic As Dictionary(Of String, String)) As String

    If Dic.ContainsKey(key) = True Then
        'strがKeyとして存在する場合

        ConvertStr = Dic(key)
        Exit Function
    End If

    ConvertStr = key

End Function

End Module
