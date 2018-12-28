Imports System.IO
Imports System.Text

Module CreZWD

#Region " Create FLP"

    Sub CreFLP()

        Dim StrB As New StringBuilder

        If RecFLP.Count = 0 Then

            CreFLPHeader(StrB)

        End If

        StrB.Append(CreSpace(FLPDic("空白1"), " "))

        FLPCnt += 1
        'シリーズNo.
        StrB.Append("(")
        StrB.Append(CreSpace(FLPDic("シリーズNo") - LenB(CStr(FLPCnt)), "0"))
        StrB.Append(CStr(FLPCnt))
        StrB.Append(")")

        StrB.Append(CreSpace(FLPDic("空白2"), " "))

        '群番

        StrB.Append("GRP---")
        StrB.Append(Gun)
        StrB.Append(CreSpace(FLPDic("GRP---") - LenB(Gun), " "))

        '盤No.
        Dim Ban As String = BanNo.Trim()
        StrB.Append("BAN---")
        StrB.Append(Ban)
        StrB.Append(CreSpace(FLPDic("BAN---") - LenB(Ban), " "))

        '前背
        StrB.Append("FB---")
        StrB.Append(FB)
        StrB.Append(CreSpace(FLPDic("FB---") - LenB(FB), " "))

        StrB.Append("SCT---")
        StrB.Append(CreSpace(FLPDic("SCT---"), " "))

        StrB.Append("ZONE---")
        StrB.Append(CreSpace(FLPDic("ZONE---"), " "))

        StrB.Append("LOCATION---")
        StrB.Append(CreSpace(FLPDic("LOCATION---"), " "))


        StrB.Append(CreSpace(FLPDic("空白3"), " "))

        RecFLP.Add(StrB.ToString)
        StrB.Clear()
    End Sub

    Sub CreFLPHeader(ByRef StrB As StringBuilder)

        StrB.Append(FLP_Header)

        '工番名
        StrB.Append("KOOBAN=")
        StrB.Append(Kouban)
        '空白詰め
        StrB.Append(CreSpace(FLPDic("KOOBAN=") - LenB(Kouban), " "))

        '客先名
        StrB.Append("KYAKU=")
        StrB.Append(CustomerName)
        '空白詰め
        StrB.Append(CreSpace(FLPDic("KYAKU=") - LenB(CustomerName), " "))

        '日時
        StrB.Append("DATE:")
        StrB.Append(NowDate)
        '空白詰め
        StrB.Append(CreSpace(FLPDic("DATE:") - LenB(NowDate), " "))

        '従番
        StrB.Append("(")
        StrB.Append(CreSpace(FLPDic("従番"), " "))
        StrB.Append(")")

        'ジョブNo.
        StrB.Append("JOB=")
        StrB.Append(Job)
        '空白詰め
        StrB.Append(CreSpace(FLPDic("JOB=") - LenB(Job), " "))

        RecFLP.Add(StrB.ToString)
        StrB.Clear()


        For Each s In Border
            RecFLP.Add(s)
        Next
    End Sub


#End Region


#Region " Create HON"

    Sub CreHON()

        Dim StrB As New StringBuilder

        If LastFlg = True Then
            If RecHON.Count = 0 Then

                StrB.Append(HON_Header1)

                '工番名
                StrB.Append(Kouban)
                '空白詰め
                StrB.Append(CreSpace(HONDic("KOOBAN=") - LenB(Kouban), " "))

                StrB.Append(HON_Header2)
                StrB.Append(CreSpace(HONDic("空白1"), " "))

                '日時
                StrB.Append("DATE:")
                StrB.Append(NowDate)
                '空白詰め
                StrB.Append(CreSpace(HONDic("DATE:") - LenB(NowDate), " "))

                '従番
                StrB.Append("(")
                StrB.Append(CreSpace(HONDic("従番"), " "))
                StrB.Append(")")

                'ジョブNo.
                StrB.Append("JOB=")
                StrB.Append(Job)
                '空白詰め
                StrB.Append(CreSpace(HONDic("JOB=") - LenB(Job), " "))
                RecHON.Add(StrB.ToString)
                StrB.Clear()

                'ヘッダー二行目
                StrB.Append(HON_Header3)
                RecHON.Add(StrB.ToString)
                StrB.Clear()
            End If


            StrB.Append(CreSpace(HONDic("空白2"), " "))
            StrB.Append(Gun)
            StrB.Append(CreSpace(HONDic("GROUP") - LenB(Gun), " "))

            '盤No.
            Dim Ban As String = BanNo.Trim()
            StrB.Append(Ban)
            StrB.Append(CreSpace(HONDic("BAN") - LenB(Ban), " "))

            StrB.Append(FB)
            StrB.Append(CreSpace(HONDic("ZENHAI") - LenB(FB), " "))

            '0.3-付属線まで
            Dim Key As String
            For i = 0 To CountRec9.Keys.Count - 1
                Key = CountRec9.Keys(i)

                '空白詰
                StrB.Append(CreSpace(HONDic(Key) - LenB(CStr(CountRec9(Key))), " "))
                '0.3-TOBI
                StrB.Append(CountRec9(Key))

                'Part
                PartHON(Key) = PartHON(Key) + CountRec9(Key)
                'Total
                TotalHON(Key) = TotalHON(Key) + CountRec9(Key)
            Next

            RecHON.Add(StrB.ToString)
            StrB.Clear()

            LastFlg = False
            ClearRec9()                 'CountRec9の値の初期化

        End If
    End Sub

    Sub SetPartHON()
            Dim StrB As New StringBuilder

        StrB.Append("     ---------------    ")

        Dim Key As String
        Dim AllCnt As Integer
        For i = 0 To PartHON.Keys.Count - 1
            Key = PartHON.Keys(i)

            '空白詰
            StrB.Append(CreSpace(HONDic(Key) - LenB(CStr(PartHON(Key))), " "))
            '0.3-TOBI
            StrB.Append(PartHON(Key))

            AllCnt += PartHON(Key)
            PartHON(Key) = 0
        Next i

        RecHON.Add(StrB.ToString)
        StrB.Clear()

    End Sub

    Sub SetTotalHON()
        Dim StrB As New StringBuilder

        SetPartHON()

        StrB.Append("     ** JOB TOTAL **    ")

        Dim Key As String
        Dim AllCnt As Integer
        For i = 0 To TotalHON.Keys.Count - 1
            Key = TotalHON.Keys(i)

            '空白詰
            StrB.Append(CreSpace(HONDic(Key) - LenB(CStr(TotalHON(Key))), " "))
            '0.3-TOBI
            StrB.Append(TotalHON(Key))

            AllCnt += TotalHON(Key)
        Next i

        RecHON.Add(StrB.ToString)
        StrB.Clear()
        StrB.Append("             (")
        StrB.Append(CreSpace(6 - LenB(CStr(AllCnt)), " "))
        StrB.Append(CStr(AllCnt))
        StrB.Append(")                                                                                                               ")
        RecHON.Add(StrB.ToString)
    End Sub


#End Region


#Region " KIK"

    Sub CreKIK()
        Dim StrB As New StringBuilder
        Dim key As String

        'ヘッダー作成
        CreKIKHeader()

        For i = 1 To UBound(ArrCoord)

            StrB.Append(" ")
            key = ArrCoord(i)
            '配列{略称,形式,端子種類,器具番号}
            Dim arr() As String = CSVDataDic(key)
            '0詰め
            StrB.Append(CreSpace(KIKDic("仮座標") - LenB(CStr(i)), "0"))
            '仮座標
            StrB.Append(CStr(i))
            '
            StrB.Append(CreSpace(KIKDic("空白2"), " "))

            '座標名
            Dim temp = Left(key, 6).Trim
            StrB.Append(temp)
            StrB.Append(CreSpace(KIKDic("座標名") - LenB(temp), " "))

            '分割名
            Dim str = Right(key, 1)
            StrB.Append(str)
            StrB.Append(CreSpace(KIKDic("分割名") - LenB(str), " "))

            '端子種類
            StrB.Append("(")
            StrB.Append(arr(2))
            StrB.Append(CreSpace(KIKDic("端子種類") - LenB(arr(2)), " "))
            StrB.Append(")")

            StrB.Append(CreSpace(KIKDic("空白3"), " "))

            '形式
            StrB.Append(arr(1))
            StrB.Append(CreSpace(KIKDic("形式") - LenB(arr(1)), " "))

            'フレーム
            StrB.Append(CreSpace(KIKDic("フレーム"), " "))

            '略称
            Dim Ryaku As String = arr(0)
            If LenB(Ryaku) > KIKDic("略称") Then
                Ryaku = MidB(Ryaku, 1, KIKDic("略称"))
            Else
                Ryaku = Ryaku & CreSpace(KIKDic("略称") - LenB(arr(0)), " ")
            End If

            StrB.Append(Ryaku)


            StrB.Append(CreSpace(KIKDic("No"), " "))

            If Left(temp, 2) = "YY" Then
                '座標名+分割名の左から二文字がYYだったとき

                Dim chukei As String = Gun & CreSpace(6 - LenB(Gun), " ") & NextBanDic(BanNo)
                StrB.Append(chukei)
                StrB.Append(CreSpace(KIKDic("CHUKEI") - LenB(chukei), " "))
            Else
                StrB.Append(CreSpace(KIKDic("CHUKEI"), " "))
            End If

            'KIGU_No
            StrB.Append(CreSpace(KIKDic("KIGU_No"), " "))

            'デバイス名
            StrB.Append(arr(3))
            StrB.Append(CreSpace(KIKDic("デバイス名") - LenB(arr(3)), " "))

            'SOU
            StrB.Append(CreSpace(KIKDic("SOU"), " "))

            '(T)
            StrB.Append(CreSpace(KIKDic("(T)"), " "))

            RecKIK.Add(StrB.ToString)
            StrB.Clear()
        Next i

    End Sub

#Region "Header"
    Sub CreKIKHeader()
        Dim StrB As New StringBuilder
        Dim jobs As String = "JOB=" & CreSpace(KIKHeaderDic("JOB=") - LenB(CStr(Job)), " ") & Job

        StrB.Append(KIKHeader1)

        '日時
        StrB.Append("DATE:")
        StrB.Append(NowDate)
        '空白詰め
        StrB.Append(CreSpace(KIKHeaderDic("DATE:") - LenB(NowDate), " "))

        '従番
        StrB.Append("(")
        StrB.Append(CreSpace(KIKHeaderDic("従番"), " "))
        StrB.Append(")")

        'ジョブNo.
        StrB.Append(jobs)

        RecKIK.Add(StrB.ToString)
        StrB.Clear()

        StrB.Append(" ")
        StrB.Append(jobs)
        StrB.Append(CreSpace(KIKHeaderDic("空白2"), " "))

        '工番
        StrB.Append("KOOBAN=")
        StrB.Append(Kouban)
        StrB.Append(CreSpace(KIKHeaderDic("KOOBAN=") - LenB(Kouban), " "))

        '群番

        StrB.Append("GRP=")
        StrB.Append(Gun)
        StrB.Append(CreSpace(KIKHeaderDic("GRP=") - LenB(Gun), " "))

        '盤No.

        StrB.Append("BAN=")
        StrB.Append(BanNo)
        StrB.Append(CreSpace(KIKHeaderDic("BAN=") - LenB(BanNo), " "))

        '前背
        StrB.Append("FB=")
        StrB.Append(FB)
        StrB.Append(CreSpace(KIKHeaderDic("FB=") - LenB(FB), " "))

        StrB.Append("SCT=")
        StrB.Append(CreSpace(KIKHeaderDic("SCT="), " "))

        StrB.Append("ZONE=")
        StrB.Append(CreSpace(KIKHeaderDic("ZONE="), " "))

        StrB.Append(CreSpace(KIKHeaderDic("空白3"), " "))

        RecKIK.Add(StrB.ToString)
        StrB.Clear()

        RecKIK.Add(CreSpace(KIKRecLen, " "))
        RecKIK.Add(KIKHeader2)
    End Sub

#End Region


#End Region


#Region " TEA･TEB･WAT共通"

'Key:文字列、Value:DictionaryとするDictionary
'ValueのDictionaryのKey,Valueはそれぞれ文字列、配列{}
Public Root As New Dictionary(Of String, Dictionary(Of String, String()))

    Public Sub CreRoot()

        'From,Toの座標名+分割名
        Dim key As String = Dic2("FT機器座標").Trim
        Dim value As String = Dic2("TT機器座標").Trim

        '線符号
        Dim Sign As String = Dic2("FT線番号")

        'From,Toの端子番号
        Dim FPin As String = Dic2("FT器具端子").Trim
        Dim TPin As String = Dic2("TT器具端子").Trim


        Dim arr As String()
        'To座標名 線サイズ 線色 線符号
        Dim str As String = value.Trim & " " & SQ.Trim & " " & Color.Trim & " " & Sign.Trim

        'Toの仮座標
        Dim tem As String = SearchTemp(value, ArrCoord)
        Dim temp As String = CreSpace(3 - LenB(tem), "0") & tem


        If Root.ContainsKey(key) = False Then
            '(0)=座標名、(1)=カウント、(2)=From器具端子、(3)=To器具端子、(4)=To座標名 線サイズ 線色 線符号
            Dim dic As New Dictionary(Of String, String()) _
             From {
                    {str, {value, "1", FPin, TPin, temp}}
                  }
            Root.Add(key, dic)
        Else
            Dim tempdic As Dictionary(Of String, String()) = Root(key)

            If tempdic.ContainsKey(str) = True Then
                'キーが存在する場合
                '(0)=座標名、(1)=カウント、(2)=From器具端子、(3)=To器具端子、(4)=To座標名 線サイズ 線色 線符号
                '配列を取得して、カウントを追加
                arr = tempdic(str)

                If arr(2) = FPin And arr(3) = TPin And arr(4) = str Then
                    arr(1) += 1
                Else

                End If


                tempdic(str) = arr
            Else
                'キーが存在しない場合
                'データを追加
                '(0)=座標名、(1)=カウント、(2)=From器具端子、(3)=To器具端子、(4)=To座標名 線サイズ 線色 線符号
                tempdic.Add(str, {value, "1", FPin, TPin, temp})
            End If

            Sort_Dic_A(tempdic)
        End If

    End Sub



     Sub CreTEABWAT()
        Dim key As String   'From座標名
        Dim dic As Dictionary(Of String, String())
        Dim str As String
        Dim temp As String
        Dim arr As String()

        PageCnt += 1

        For i As Integer = 0 To Root.Count - 1
            '座標名の分だけ繰り返す
            key = Root.Keys(i)      'From座標名
            dic = Root(key)

            If i = 0 Then
                CreTEHeaderAorB(HeaderTEA, PageCnt, RecTEA)
                CreTEHeaderAorB(HeaderTEB, PageCnt, RecTEB)
            End If

            For j As Integer = 0 To dic.Count - 1
                '仮座標の分だけ繰り返す
                str = dic.Keys(j)

                'arr() = (0)=座標名、(1)=カウント、(2)=From器具端子、(3)=To器具端子、(4)=To仮座標
                arr = dic(str)

                temp = arr(4)
                If Left(key, 2) = "YY" Or Left(arr(0), 2) = "YY" Then
                    '中継線なら
                    If TEBCount = 0 And TEBExtenFlg = False Then
                        'From側を作成
                        StartTE(i, key, TEBDic, LineTEB)
                    ElseIf TEBExtenFlg = True Then
                        TEBExtenFlg = False
                    End If

                    'To側を作成
                    CreTEB(arr(0), temp, arr(1))

                    If j = dic.Count - 1 Then
                        If LineTEA.Length <> 0 Then
                            EndTE(SetTEA, LineTEA, TEACount, TotalTEACount, TotalTEA, RecTEA_Data)
                        End If

                        EndTE(SetTEB, LineTEB, TEBCount, TotalTEBCount, TotalTEB, RecTEB_Data)
                    End If
                ElseIf key = arr(0) Then
                    '渡線の時(Fromの座標とToの座標が同じとき)



                    CreWAT_Data(arr(2), arr(3), str)
                Else

                    If TEACount = 0 And TEAExtenFlg = False Then
                        'From側を作成
                        StartTE(i, key, TEADic, LineTEA)
                    ElseIf TEAExtenFlg = True Then
                        TEAExtenFlg = False
                    End If

                    'To側を作成
                    CreTEA(arr(0), arr(1))

                    If j = dic.Count - 1 Then
                        EndTE(SetTEA, LineTEA, TEACount, TotalTEACount, TotalTEA, RecTEA_Data)

                        If LineTEB.Length <> 0 Then
                            EndTE(SetTEB, LineTEB, TEBCount, TotalTEBCount, TotalTEB, RecTEB_Data)
                        End If
                    End If
                End If
            Next
        Next i

        Root.Clear()
        RecTEA_Data.Sort()
        RecTEB_Data.Sort()

        For Each d In RecTEA_Data
            RecTEA.Add(d)
        Next

        For Each d In RecTEB_Data
            RecTEB.Add(d)
        Next

        RecTEA_Data.Clear() : RecTEB_Data.Clear()

        CreRecWAT()
     End Sub

#End Region


#Region " TEA･TEB 共通Header"

        Public Sub CreTEHeaderAorB(ByVal AorBH As String, ByVal Page As String, ByRef RecTE As List(Of String))
        Dim StrB As New StringBuilder
        Dim jobs As String = "JOB=" & CreSpace(TEHeaderDic("JOB=") - LenB(CStr(Job)), " ") & CStr(Job)
        Dim ReList As New List(Of String)

        StrB.Append(AorBH)

        StrB.Append("DATE:")
        StrB.Append(NowDate)
        StrB.Append(CreSpace(TEHeaderDic("DATE:") - LenB(NowDate), " "))

        StrB.Append("PAGE=")
        StrB.Append(CreSpace(TEHeaderDic("PAGE=") - LenB(Page), "0"))
        StrB.Append(Page)

        StrB.Append(CreSpace(TEHeaderDic("空白1"), " "))

        '従番
        StrB.Append("(")
        StrB.Append(CreSpace(TEHeaderDic("従番"), " "))
        StrB.Append(")")

        StrB.Append(jobs)

        RecTE.Add(StrB.ToString)
        StrB.Clear()

        StrB.Append(CreSpace(TEHeaderDic("空白2"), " "))

        'ジョブNo.
        StrB.Append(jobs)

        StrB.Append(CreSpace(TEHeaderDic("空白3"), " "))

        '工番
        StrB.Append("KOOBAN---")
        StrB.Append(Kouban)
        StrB.Append(CreSpace(TEHeaderDic("KOOBAN---") - LenB(Kouban), " "))

        '群番

        StrB.Append("GRP---")
        StrB.Append(Gun)
        StrB.Append(CreSpace(TEHeaderDic("GRP---") - LenB(Gun), " "))

        '盤No.
        Dim Ban As String = BanNo.Trim()
        StrB.Append("BAN---")
        StrB.Append(Ban)
        StrB.Append(CreSpace(TEHeaderDic("BAN---") - LenB(Ban), " "))

        '前背
        StrB.Append("FB---")
        StrB.Append(FB)
        StrB.Append(CreSpace(TEHeaderDic("FB---") - LenB(FB), " "))

        StrB.Append("SCT---")
        StrB.Append(CreSpace(TEHeaderDic("SCT---"), " "))

        StrB.Append("ZONE---")
        StrB.Append(CreSpace(TEHeaderDic("ZONE---"), " "))

        StrB.Append(CreSpace(TEHeaderDic("空白4"), " "))

        RecTE.Add(StrB.ToString)
    End Sub


     Sub StartTE(ByVal i As Integer, ByVal key As String, ByVal Dic As Dictionary(Of String, Object), ByRef Line As StringBuilder)

        Dim temp As String = SearchTemp(key, ArrCoord)
        Line.Append(" ")
        Line.Append(CreSpace(Dic("From仮座標") - LenB(temp), "0"))  '0詰め
        Line.Append(temp)

        Line.Append(CreSpace(Dic("空白2"), " "))

        Line.Append(key)
        Line.Append(CreSpace(Dic("From座標") - LenB(key), " "))

        Line.Append(",")
        Line.Append(CreSpace(Dic("From分割名"), " "))
        Line.Append(CreSpace(Dic("-"), "-"))
     End Sub


     Sub EndTE(ByVal Dic As Dictionary(Of String, Integer), ByRef Line As StringBuilder, ByRef TECount As Integer, ByRef TotalTECount As Integer, ByVal Totalstr As String, ByRef RecTE As List(Of String))

        Line.Append(Totalstr)
        Line.Append("(")
        Line.Append(CreSpace(Dic("本数") - LenB(TotalTECount), "0"))
        Line.Append(TotalTECount)
        Line.Append(")")

        Line.Append(CreSpace(RecTELen - LenB(Line.ToString), " "))
        RecTE.Add(Line.ToString)
        Line.Clear()
        TECount = 0 : TotalTECount = 0
     End Sub

#End Region


#Region " Create TEA"

    Sub CreTEA(ByVal TCoord As String, ByVal Hon As String)




        LineTEA.Append(" ")
        LineTEA.Append(TCoord)
        LineTEA.Append(CreSpace(SetTEA("To座標") - LenB(TCoord), " "))

        LineTEA.Append(",")
        LineTEA.Append(CreSpace(SetTEA("To分割名"), " "))

        LineTEA.Append("(")
        LineTEA.Append(CreSpace(SetTEA("本数") - LenB(Hon), "0"))  '0詰め
        LineTEA.Append(Hon)
        LineTEA.Append(")")

        TotalTEACount += Hon

        TEACount += 1

        If TEACount = TEAMAX Then

            LineTEA.Append(" ")
            RecTEA.Add(LineTEA.ToString)
            LineTEA.Clear()
            LineTEA.Append(NextSpace)
            TEAExtenFlg = True
            TEACount = 0

        End If
    End Sub

#End Region


#Region " Create TEB"


    Sub CreTEB(ByVal TCoord As String, ByVal TempT As String, ByVal Hon As String)




        LineTEB.Append(" ")
        LineTEB.Append(TCoord)
        LineTEB.Append(CreSpace(SetTEB("To座標") - LenB(TCoord), " "))

        LineTEB.Append(",")
        LineTEB.Append(CreSpace(SetTEB("To分割名"), " "))

        LineTEB.Append("<<")
        LineTEB.Append(CreSpace(SetTEB("To仮座標") - LenB(TempT), "0"))  '0詰め
        LineTEB.Append(TempT)
        LineTEB.Append(">>")

        LineTEB.Append("(")
        LineTEB.Append(CreSpace(SetTEB("本数") - LenB(Hon), "0"))  '0詰め
        LineTEB.Append(Hon)
        LineTEB.Append(")")

        TotalTEBCount += Hon

        TEBCount += 1

        If TEBCount = 5 Then
            '不足分空白詰
            LineTEB.Append(CreSpace(RecTELen - LineTEB.Length, " "))
            RecTEB.Add(LineTEB.ToString)
            LineTEB.Clear()
            LineTEB.Append(NextSpace)
            TEBExtenFlg = True
            TEBCount = 0

        End If
    End Sub

#End Region


#Region " Create WAT"




    Sub CreWATHeader()
        Dim StrB As New StringBuilder
        Dim jobs As String = "JOB=" & CreSpace(WATHeader("JOB=") - LenB(CStr(Job)), " ") & CStr(Job)


        StrB.Append(TitleWAT)

        '工番
        StrB.Append("KOOBAN=")
        StrB.Append(Kouban)
        StrB.Append(CreSpace(WATHeader("KOOBAN=") - LenB(Kouban), " "))

        '群番
        StrB.Append("GRP=")
        StrB.Append(Gunban)
        StrB.Append(CreSpace(WATHeader("GRP=") - LenB(Gunban), " "))

        '盤No.
        Dim Ban As String = BanNo.Trim()
        StrB.Append("BAN=")
        StrB.Append(Ban)
        StrB.Append(CreSpace(WATHeader("BAN=") - LenB(Ban), " "))

        '前背
        StrB.Append("FB=")
        StrB.Append(FB)
        StrB.Append(CreSpace(WATHeader("FB=") - LenB(FB), " "))

        'SCT
        StrB.Append("SCT=")
        StrB.Append(CreSpace(WATHeader("SCT="), " "))

        'DATE
        StrB.Append("DATE:")
        StrB.Append(NowDate)
        StrB.Append(CreSpace(WATHeader("DATE:") - LenB(NowDate), " "))

        '従番
        StrB.Append("(")
        StrB.Append(CreSpace(WATHeader("従番"), " "))
        StrB.Append(")")

        'JOBNO.
        StrB.Append(jobs)

        RecWAT.Add(StrB.ToString)
        RecWAT.Add(BorderWAT)
    End Sub

    Sub CreWAT_Data(ByVal FromPin As String, ByVal ToPin As String, ByVal str As String)

        Dim arr As String() = str.Split(" ")

        If WATDataDic.ContainsKey(str) = False Then

            Dim list As List(Of String) = New List(Of String)({FromPin, ToPin})
            list.Sort()
            'From,Toの端子番号を追加
            WATDataDic.Add(str, list)

        Else
            '端子番号を追加
            Dim templist As List(Of String) = WATDataDic(str)

            If templist.Contains(FromPin) = False Then
                templist.Add(FromPin)
            End If

            If templist.Contains(ToPin) = False Then
                templist.Add(ToPin)
            End If

            templist.Sort()

            WATDataDic(str) = templist
        End If

    End Sub

    Public OldCoord As String


    Sub CreRecWAT()
        Dim StrB As New StringBuilder
        Dim arr() As String
        Dim list As List(Of String) '端子番号一覧リスト
        Dim coord As String         '座標
        Dim sq As String            '線サイズ
        Dim color As String         'カラー
        Dim sign As String          '線符号

        'ヘッダー作成
        CreWATHeader()


        WATDataDic = Sort_Dic_L(WATDataDic)

        For Each s In WATDataDic.Keys

            arr = s.Split(" ")

            '座標:線サイズ:カラー:線符号
            coord = arr(0) : sq = arr(1) : color = arr(2) : sign = arr(3)

            list = WATDataDic(s)    '端子番号一覧リスト

            If OldCoord = coord Then
                StrB.Append(CreSpace(WATDic("座標名"), " "))
            Else
                If OldCoord IsNot Nothing Then
                    RecWAT.Add(BorderWAT)
                End If

                StrB.Append(CreSpace(WATDic("座標名") - LenB(coord), " "))
                StrB.Append(coord)
            End If


            StrB.Append(" ")

            StrB.Append(sq)
            StrB.Append(CreSpace(WATDic("線サイズ") - LenB(sq), " "))

            StrB.Append(" ")

            StrB.Append(color)
            StrB.Append(CreSpace(WATDic("色") - LenB(color), " "))

            StrB.Append(CreSpace(WATDic("空白3"), " "))

            For i As Integer = 0 To list.Count - 1
                '端子番号を追加
                StrB.Append(list(i))

                If i <> list.Count - 1 Then
                    '最後のデータじゃない時
                    StrB.Append(" * ")
                End If

                If StrB.Length + NextData > PartLen Then
                    '領域オーバーする恐れがある場合
                    RecWAT.Add(StrB.ToString)
                    StrB.Clear()
                    StrB.Append("                     ")
                End If

                If i = list.Count - 1 Then
                    '線符号の領域まで空白詰め
                    StrB.Append(CreSpace(PartLen - StrB.Length, " "))
                End If
            Next i

            StrB.Append(sign)
            StrB.Append(CreSpace(WATDic("線符号") - LenB(sign), " "))

            RecWAT.Add(StrB.ToString)
            StrB.Clear()
            OldCoord = coord
        Next
        RecWAT.Add(CreSpace(RecWATLength, " "))
        WATDataDic.Clear()
    End Sub
#End Region

End Module
