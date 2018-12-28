Imports System.IO
Imports System.Text

Module WriteData



#Region "結果をファイルに出力"

''' <summary>
''' DATファイルのデータ出力
''' </summary>
''' <param name="FinalFlg">最終出力かどうかのフラグ。最終出力ならTrue</param>
''' <returns>エラーが発生せずに終了すればTrue。エラーが発生すればFalse</returns>
''' <remarks></remarks>

    Public Function OutPutDATFiles(ByVal FinalFlg As Boolean) As Boolean

        ' 出力ファイルパスを組み立てる
        Dim opWD As String = OutPutWDPath & "WD" & Job & ".DAT"
        Dim opNC As String = OutPutNCPath & "NC" & Job & ".DAT"
        Dim opDEV As String = OutPutNCPath & "DEV" & Job & ".DAT"
        Dim opMRK As String = OutPutNCPath & "MRK" & Job & ".DAT"
        Dim opMJ As String = OutPutNCPath & "MJ" & Job & ".DAT"
        Dim opHON As String = OutPutNCPath & "HON" & Job & ".DAT"

        Dim i As Integer

        Try


            '作成するStreamWriterの宣言
            Using WD As New StreamWriter(opWD, FirstOutFlg, Encoding.Default)
                Using NC As New StreamWriter(opNC, FirstOutFlg, Encoding.Default)
                    Using DEV As New StreamWriter(opDEV, FirstOutFlg, Encoding.Default)
                        Using MRK As New StreamWriter(opMRK, FirstOutFlg, Encoding.Default)
                            Using MJ As New StreamWriter(opMJ, FirstOutFlg, Encoding.Default)
                                Using HON As New StreamWriter(opHON, FirstOutFlg, Encoding.Default)
                                    'レコード1出力
                                    WD.WriteLine(Rec1)
                                    NC.WriteLine(Rec1)
                                    DEV.WriteLine(Rec1)
                                    MRK.WriteLine(Rec1)
                                    MJ.WriteLine(Rec1)
                                    HON.WriteLine(Rec1)

                                    'レコード2出力
                                    For i = 0 To Rec2.Count - 1
                                        WD.WriteLine(Rec2(i))
                                        NC.WriteLine(Rec2(i))
                                    Next

                                    'レコード3出力
                                    For i = 0 To Rec3.Count - 1
                                        WD.WriteLine(Rec3(i))
                                        NC.WriteLine(Rec3(i))
                                    Next

                                    'レコード4出力
                                    For i = 0 To Rec4.Count - 1
                                        WD.WriteLine(Rec4(i))
                                        NC.WriteLine(Rec4(i))
                                    Next

                                    'レコード5出力
                                    For i = 0 To Rec5.Count - 1
                                        WD.WriteLine(Rec5(i))
                                        DEV.WriteLine(Rec5(i))
                                    Next

                                    'レコード6出力
                                    For i = 0 To Rec6.Count - 1
                                        WD.WriteLine(Rec6(i))
                                        MRK.WriteLine(Rec6(i))
                                    Next

                                    'レコード7出力
                                    For i = 0 To Rec7.Count - 1
                                        WD.WriteLine(Rec7(i))
                                        MJ.WriteLine(Rec7(i))
                                    Next

                                    If FinalFlg = True Then
                                        For i = 0 To Rec9.Count - 1
                                            'レコード9出力
                                            WD.WriteLine(Rec9(i))
                                            HON.WriteLine(Rec9(i))

                                        Next i
                                        Rec9.Clear()

                                    End If

                                    'データのクリア
                                    Rec1 = "" : Rec2.Clear() : Rec3.Clear()
                                    Rec4.Clear() : Rec5.Clear() : Rec6.Clear()
                                    Rec7.Clear()
                                    Rec3_4_Data.Clear()
                                    SortedRec3_4_Data.Clear()
                                    Rec6_Data.Clear()
                                    SortedRec6_Data.Clear()
                                    Rec7_DataList.Clear()

                                    'ファイルを閉じる
                                    WD.Close() : NC.Close() : DEV.Close() : MRK.Close() : MJ.Close() : HON.Close()
                                End Using   'HON    
                            End Using   'MJ
                        End Using   'MRK
                    End Using   'DEV
                End Using   'NC
            End Using   'WD

            If FirstOutFlg = False Then
                'FirstOutFlgの初期値はFalse
                'TrueからFalseへの値の変更処理はなし
                '一回目の出力の時のみ、ファイルの上書きをおこなう
                '一回目の出力以降はファイルの末尾に追加する。
                FirstOutFlg = True
            End If


            Row3Cnt = 1 : Rec3Cnt = 0
            Row4Cnt = 1 : Rec4Cnt = 0

            OutPutDATFiles = True
        Catch ex As System.IO.IOException
            'ファイルが開けなかったエラーの場合
            Console.WriteLine(opWD & vbCrLf & opNC & vbCrLf & opDEV & vbCrLf & opMRK & vbCrLf & opMJ & vbCrLf & opHON)
            Console.WriteLine("いずれかのファイルが開かれている可能性があります。")
            OutPutDATFiles = False
        Catch ex As Exception
            'その他の場合のエラー
            OutPutDATFiles = False
        End Try
    End Function

    ''' <summary>
    ''' ZWDファイルのデータ出力
    ''' </summary>
    ''' <remarks></remarks>
    Sub OutPutZWDFiles()
        Dim opFLP As String = OutPutLISTPath & "ZWD05FLP." & Job
        Dim opHON As String = OutPutLISTPath & "ZWD05HON." & Job
        Dim opKIK As String = OutPutLISTPath & "ZWD05KIK." & Job
        Dim opTBP As String = OutPutLISTPath & "ZWD05TBP." & Job
        Dim opTEA As String = OutPutLISTPath & "ZWD05TEA." & Job
        Dim opTEB As String = OutPutLISTPath & "ZWD05TEB." & Job
        Dim opWAT As String = OutPutLISTPath & "ZWD05WAT." & Job

        Dim i As Integer


        '作成するStreamWriterの宣言
        Using FLP As New StreamWriter(opFLP, False, Encoding.Default)
            Using HON As New StreamWriter(opHON, False, Encoding.Default)
                Using KIK As New StreamWriter(opKIK, False, Encoding.Default)
                    Using TBP As New StreamWriter(opTBP, False, Encoding.Default)
                        Using TEA As New StreamWriter(opTEA, False, Encoding.Default)
                            Using TEB As New StreamWriter(opTEB, False, Encoding.Default)
                                Using WAT As New StreamWriter(opWAT, False, Encoding.Default)

                                    For i = 0 To RecFLP.Count - 1
                                        FLP.WriteLine(RecFLP(i))
                                    Next

                                    For i = 0 To RecHON.Count - 1
                                        HON.WriteLine(RecHON(i))
                                    Next

                                    For i = 0 To RecKIK.Count - 1
                                        KIK.WriteLine(RecKIK(i))
                                    Next

                                    For i = 0 To RecTEA.Count - 1
                                        TEA.WriteLine(RecTEA(i))
                                    Next

                                    For i = 0 To RecTEB.Count - 1
                                        TEB.WriteLine(RecTEB(i))
                                    Next


                                    For i = 0 To RecWAT.Count - 1
                                        WAT.WriteLine(RecWAT(i))
                                    Next

                                End Using
                            End Using
                        End Using
                    End Using
                End Using
            End Using
        End Using


    End Sub


#End Region


#Region "工番体系調整"

    ''' <summary>
    ''' 工番体系を整える。
    ''' </summary>
    ''' <param name="str">調べる文字列</param>
    ''' <returns>X**-****のような工番体系を返す</returns>
    ''' <remarks></remarks>
    Public Function TrimOrder(ByVal str As String) As String
        Dim result As Integer   '変換結果
        '半角+大文字にして空白全消し
        Dim rtn As String = Replace(StrConv(str, VbStrConv.Narrow + VbStrConv.Uppercase), " ", "")


        If Integer.TryParse(Left(rtn, 2), result) = True Then
            '引数文字列左から二文字がIntegerへの変換が成功したとき
            '99X16-338などの場合
            '3文字目からの文字列を返す。
            TrimOrder = Mid(rtn, 3)
        Else
            '変換が成功しなかったとき
            'X16-338 →　X1　などの時
            TrimOrder = rtn

        End If

        Dim f As String = Mid(TrimOrder, 1, 1)
        Dim s As String = Mid(TrimOrder, 2, 1)

        If f >= "A" And f <= "Z" And s >= "A" And s <= "Z" Then
            '2文字ともアルファベットの場合
            'XP等の場合→Xに
            TrimOrder = TrimOrder.Replace(f & s, f)
        End If

    End Function



    ''' <summary>
    ''' 群番体系を整える
    ''' </summary>
    ''' <param name="gun">群番</param>
    ''' <returns>体系を整えた群番</returns>
    ''' <remarks></remarks>
    Public Function TrimGroup(ByVal gun As String) As String

        Dim rtn As String = ""  '戻り値

        '(0)=アルファベット、(1)=数字
        Dim arr() = SepaStr(gun)

        '
        TrimGroup = arr(0) & CStr(CInt(arr(1)))
    End Function

    ''' <summary>
    ''' 0付工番を返す
    ''' </summary>
    ''' <param name="kouban"></param>
    ''' <remarks></remarks>
    Public Function KoubanAdd0(ByVal kouban As String) As String

        Dim arr As String() = kouban.Split("-")

        If arr(1).Length <> 4 Then

            kouban = arr(0) & "-" & CreSpace(4 - arr(1).Length, "0") & arr(1)
        End If

        KoubanAdd0 = kouban
    End Function

    ''' <summary>
    ''' 工番をデータベース用に体系を整える
    ''' </summary>
    ''' <param name="Order"></param>
    ''' <remarks></remarks>
    Public Sub DBOrder(ByVal Order As String)
        'OrderはX**-****がくることを想定
        Dim arr As String() = Nothing

        arr = Order.Split("-")

        While arr(1).Length <> 4
            arr(1) = "0" & arr(1)
        End While

        arr(0) = "  " & arr(0)
        arr(0) = arr(0).Insert(3, " ")

        DBKouban = arr(0) & arr(1)
    End Sub
#End Region


#Region "FromTo関係"

    ''' <summary>
    ''' FromTo関係のPublic変数セット
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
    Sub SetPubFT()

        '線符号
        LFSign = Dic2("FT線番号")
        LTSign = Dic2("TT線番号")

        'From
        FromCoord = Dic2("FT機器座標")  '座標名
        FromPin = Dic2("FT器具端子")    '端子番号
        FromClr = Dic2("FT色別")        'カラー
        FromTemp = SearchTemp(FromCoord, ArrCoord) '仮座標

        'To
        ToCoord = Dic2("TT機器座標")    '座標名
        ToPin = Dic2("TT器具端子")      '端子番号
        ToClr = Dic2("TT色別")          'カラー
        ToTemp = SearchTemp(ToCoord, ArrCoord)     '仮座標
    End Sub



    ''' <summary>
    ''' 座標名+分割名から仮座標を求める
    ''' </summary>
    ''' <param name="coord">座標名＋分割名(空白無)</param>
    ''' <returns>仮座標</returns>
    ''' <remarks></remarks>

    Public Function SearchTemp(ByVal coord As String, ByVal arr As String()) As String

        '見つからなければ、空白を返す
        SearchTemp = "   "

        For i As Integer = 1 To UBound(arr)

            coord = Replace(coord, " ", "")
            arr(i) = Replace(arr(i), " ", "")

            If coord = arr(i) Then
                SearchTemp = CStr(i)   '0始まりなので+1
                Exit For
            End If
        Next

        'If SupRec4 = "   " Then
        '    MsgBox("仮座標が見つかりませんでした")
        'End If

    End Function

#End Region


End Module
