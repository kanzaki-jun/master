Imports System.IO
Imports System.Text

Module Main



#Region "メイン"

#Region "工番、群番、JobNo.を取得"

    Sub GetKouGunJob(ByVal CmdArg() As String)


        Dim tempArr() As String

        For Each s In CmdArg

            tempArr = s.Split(",")

            If (tempArr Is Nothing) = False Then

                '工番、群番、ジョブNo.を取得
                Kouban = TrimOrder(tempArr(0))
                Kouban0 = WriteData.KoubanAdd0(Kouban)
                Gunban = TrimGroup(tempArr(1))
                Job = CreSpace(ComDic("ジョブNo") - LenB(CStr(tempArr(2))), "0") & tempArr(2)
            End If
        Next
    End Sub
#End Region

#Region "パス関係"

''' <summary>
''' CSVファイルのパスを作成し、ファイルの存在を確認する
''' </summary>
''' <returns>ファイルが存在すればTrue、ファイルが存在しなければFalse</returns>
''' <remarks></remarks>
    Public Function SelectCSVPath() As Boolean
        Dim index As String = Dic1("インデックス").Trim
        Dim path As String = E3Path & "\" & Kouban0 & "-" & Gunban & "-" & BanNo

        'パス作成
        CoordPath = path & "_端点座標_" & IndexToArea(BanNo & "-" & index) & ".csv"  '自動配線後_端点座標へのパス
        CSVDataPath = path & "_Conv_" & IndexToArea(BanNo & "-" & index) & ".csv"    '座標＋分割名に対する略称、形式、端子種類、器具番号の対応CSVファイルへのパス


        If System.IO.File.Exists(CoordPath) = True And System.IO.File.Exists(CSVDataPath) Then
            'ファイルの存在を確認して、両方存在すれば
            Console.WriteLine("端点座標、座標データCSVファイル存在確認")
            SelectCSVPath = True
        Else
            MsgBox("CSVファイルが見つかりませんでした。")
            SelectCSVPath = False
        End If

    End Function

    ''' <summary>
    ''' E3出力先へのパスを設定
    ''' </summary>
    ''' <remarks>パスデータはE3Pathに保存</remarks>
    Sub SetE3OutPath()
        Dim typestr As String = ""
        Dim path As String = DefE3Path

        Select Case Type

            Case "産"
                typestr = "産業\"

            Case "水"
                typestr = "水\"

            Case "鉄"
                typestr = "電鉄\"
            Case "電"
                typestr = "電力\"

            Case "道"
                typestr = "道路\"

        End Select

        path = path & typestr & "E3\"

        'X16-338 なら　0=X,1=16338
        Dim KI As String = Left(SepaStr(Kouban)(1), 2)

        E3Path = path & KI & "期\" & Kouban0 & "\" & Gunban


    End Sub

    ''' <summary>
    ''' E3から出力される配線データと変換後のデータの更新日を比較して
    ''' E3から出力される配線データの方が新しければTrueを返す。
    ''' それ以外ならFalseを返す。
    ''' </summary>
    ''' <param name="E3Path"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckFiles(ByVal E3Path As String) As Boolean

        'データ出力先へのパスが正しいかチェック
        If CheckDir(OutPutWDPath) = True And _
            CheckDir(OutPutNCPath) = True And _
            CheckDir(OutPutLISTPath) = True Then
            'パスが正しい場合

            'E3のファイルとデータ変換のファイルを比較して、
            'E3のファイルのほうが新しければTrueを返す
            Dim E3Date As Date = System.IO.File.GetLastWriteTime(E3Path)
            Dim OutDate As Date = System.IO.File.GetLastWriteTime(OutPutWDPath & "\WD" & Job & ".DAT")


            If E3Date > OutDate Then
                CheckFiles = True
            Else
                MsgBox("E3出力データの新規更新がないので、処理を中断します。", , "処理中断")
                CheckFiles = False
            End If
        Else
            'パスが正しくない場合
            CheckFiles = False
        End If
    End Function

    ''' <summary>
    ''' フォルダがあるかどうか
    ''' </summary>
    ''' <param name="ChkPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CheckDir(ByVal ChkPath As String) As Boolean

        Dim temp As Boolean = System.IO.Directory.Exists(ChkPath)
        CheckDir = True

        If temp = False Then
            '
            MsgBox(ChkPath & "のフォルダがありません。")

            CheckDir = False
        End If

    End Function


#End Region


    '前提として
    'E3出力ファイル
    'エリアとインデックスの対応表
    '端点座標のCSVファイル
    '座標名、略称、形式、端子種類、器具番号を一行とするCSVファイル
    '以下4ファイルがあると想定したうえで動作。

     Sub Main(ByVal CmdArg() As String)


        Dim Line As String                          'E3出力ファイルの一行の文字列
        Dim Max As Integer                          'For文の最大ループ数
        Dim ArrArea As New ArrayList                'インデックスとエリアの対応表に使用する仮変数
        Dim inputFile As String = ""                '入力用ファイル


        '工番、群番、JobNo.を引数から取得
        GetKouGunJob(CmdArg)

        'データベース用の工番DBKoubanを求める
        DBOrder(Kouban)

        'データベースから客先名を取得
        ConnectODB(DBKouban)

        'E3出力先のパスを設定
        SetE3OutPath()



        'E3データ出力先の工番と群番に一致したファイルのパスを複数取得
        Dim E3Files As String() = _
          System.IO.Directory.GetFiles(E3Path, Kouban0 & Gunban & "*" & ".*")



        If E3Files.Length = 0 Then
            'E3出力ファイルが存在しない場合
            'メッセージを出して処理を中断

            MsgBox("E3の出力データを見つけることができませんでした。" & vbCrLf _
            & "JobNo.を間違えて入力している可能性があります。" & vbCrLf & _
            "もしくは、JobNo.登録先の工番、群番の入力ミスの可能性があります。", , "エラーにより処理中断")

            Exit Sub
        End If

        Console.WriteLine("E3出力ファイル読み込み開始")

        'データ変換の出力先のフォルダがない、もしくはE3出力ファイルがデータ変換よりも古いかどうかチェック
        'If CheckFiles(E3Files(0)) = False Then
        '    ErrEndFlg = True
        '    Console.WriteLine("E3出力先フォルダがないもしくは、E3出力ファイルが変換データより古いです。")
        '    GoTo BefEnd
        'End If


        Max = UBound(E3Files)



        '取得したファイルの分だけ繰り返す
        For i As Integer = 0 To Max
            '複数ファイルのデータを１つのListに纏める

            '複数取得したファイルパスのうちの一つを選択
            inputFile = E3Files(i)

            Try
                'パスのファイルを開く

                Using reader As New StreamReader(inputFile, Encoding.Default)
                    Console.WriteLine(inputFile & " 読み込み")
                    While (reader.Peek() > -1)
                        '最終行まで一行ずつ読み込む
                        Line = reader.ReadLine()

                        If Left(Line, 2) = "01" Then
                            'ヘッダー行の時箱番を取得
                            BanNo = Trim(Right(Line, 4))
                        End If
                        '読み込んだ一行をPublic listのReadListに保存
                        ReadList.Add(Line)
                    End While
                    reader.Close()
                End Using
            Catch ex As System.IO.IOException
                Console.WriteLine(inputFile & " が読み込めませんでした。")
                Console.WriteLine(inputFile & " が開かれている可能性があります。")

            End Try
        Next i

        Console.WriteLine("E3出力ファイル読み込み終了")

        Console.WriteLine("エリア情報読み込み開始")

        'エリア情報CSVファイルへのパス
        AreaCSVPath = E3Path & "\" & Kouban0 & "-" & Gunban & "-" & BanNo & "_Area.csv"

        If System.IO.File.Exists(AreaCSVPath) = False Then

            'エリア情報CSVファイルが存在しない場合
            Console.WriteLine("エリア情報の取得に失敗しました。")
            ErrEndFlg = True
            GoTo BefEnd
        End If

        'インデックスとエリアの対応表作成
        'ArrayListに読み込み結果をArrayList型に変換して保存
        ArrArea = CType(ReadCSV(AreaCSVPath, False), System.Collections.ArrayList)

        For i = 1 To ArrArea.Count - 1
            Dim temp() As String = CType(ArrArea(i), String())
            'temp(0)=インデックス、temp(1)=エリア
            IndexToArea.Add(BanNo & "-" & temp(1), temp(2))
        Next

        Console.WriteLine("エリア情報読み込み終了")

        '製作情報ファイルへのパス作成
        ProdInfo = E3Path & "\" & "製作情報_" & Kouban0 & "_" & Gunban & ".xlsx"




        '製作情報ファイルからマークバンド、盤配列情報を読み込む
        If ReadData.LoadXLSX() = False Then GoTo BefEnd


        Console.WriteLine("変換処理開始")

        'メインループ内で問題が起きれば
        If MainLoop() = False Then
            GoTo BefEnd
        End If


        SortedRec3_4_Data = Sort_Dic_L(Rec3_4_Data)
        SortedRec6_Data = Sort_Dic_L(Rec6_Data)
        CreRec3()
        CreRec4()
        CreRec6()
        CreRec7()
        CreFLP()
        SetTotalHON()
        CreTEABWAT()
        OutPutDATFiles(True)
        OutPutZWDFiles()

        Console.WriteLine("変換処理終了")
BefEnd:
        If ErrEndFlg = False Then
            'エラーが起きずに終了した場合
            Console.WriteLine("工番:" & Kouban & " 群番:" & Gunban & " のデータを作成しました。" & vbCrLf & "正常終了")
        Else
            Console.WriteLine("工番:" & Kouban & " 群番:" & Gunban & "のデータ作成失敗" & vbCrLf & "異常終了")
        End If
        Console.WriteLine("終了するには何かキーを押してください．．．")
        Console.ReadKey()
    End Sub


    Public Function MainLoop() As Boolean
        Dim Line As String                          'E3出力ファイルの一行の文字列
        Dim Header As String                        'E3出力ファイルのヘッダの値(01=ヘッダ行、02=布線行)
        Dim Place As Integer                        '現在の文字列内での位置

        'ListをFor文でまわす
        For i = 0 To ReadList.Count - 1
            '1データの読み込み
            Line = ReadList(i)
            '左から2文字はHeader行
            Header = Left(Line, 2)
            '開始位置を初期値へ
            Place = 3

            'データとして最終行の場合、最終行フラグを立てる
            If i <> 0 And i <> ReadList.Count - 1 Then
                If Left(ReadList(i + 1), 2) = "01" Then
                    'ヘッダが切り替わったとき
                    LastFlg = True
                End If
            ElseIf i = ReadList.Count - 1 Then
                '最後のデータの場合
                LastFlg = True
            End If


            Select Case Header
                Case "01"
                    'ヘッダ行の時

                    If i <> 0 Then
                        'i≠0のときで、頭二文字が01の時
                        '初回以外のヘッダ行の時

                        SortedRec3_4_Data = Sort_Dic_L(Rec3_4_Data)
                        SortedRec6_Data = Sort_Dic_L(Rec6_Data)

                        CreRec3() : CreRec4() : CreRec6() : CreRec7()
                        CreFLP() : CreTEABWAT()
                        If OutPutDATFiles(False) = False Then
                            'データの書き込み途中でエラーが発生した場合
                            ErrEndFlg = True
                            Exit For        'For文を抜ける
                        End If
                    End If


                    'キーと値の削除
                    Dic1.Clear()

                    'ヘッダ行に関するDictionaryを作成
                    CreDic(Dic1, Area1, AreaName1, Line, Place)

                    '盤No.
                    BanNo = Dic1("盤No").Trim()

                    If i = 0 Then
                        AllBanNo.Add(BanNo)
                    ElseIf AllBanNo.Contains(BanNo) = False Then
                        AllBanNo.Add(BanNo)
                        SetPartHON()
                    End If

                    '前背
                    FB = IndexToArea(BanNo & "-" & Dic1("インデックス").Trim)
                    FB = CreSpace(2 - LenB(FB), " ") & FB

                    'マークバンド
                    MarkBand = MarkDic(Kouban0 & "-" & Gunban)

                    '群番(Space付)
                    Dim arr As String() = SepaStr(Gunban)
                    Gun = arr(0) & " " & CreSpace(2 - LenB(arr(1)), "0") & arr(1)

                    'csvファイルのパスを作成し、ファイルが存在すればTrue、存在しなければFalseを返す
                    If SelectCSVPath() = False Then
                        ErrEndFlg = True
                        Exit For
                    End If

                    'CSVファイルを読み込み、座標名＋分割名をArrCoordに保存
                    CreCoord(CType(ReadCSV(CoordPath, True), System.Collections.ArrayList))

                    '座標名,分割名,略称,形式,端子種類,器具番号のCSVデータを読み込み
                    'キー:座標名+分割名
                    '値  :{略称,形式,端子種類,器具番号}
                    'CSVDataDicを作成
                    ParseCSVData(CType(ReadCSV(CSVDataPath, True), System.Collections.ArrayList))



                    'レコードNo.1　作成
                    CreRec1()

                    'レコードNo.2　作成
                    CreRec2()

                    'レコードNo.5　作成
                    CreRec5()

                    'ZWD05KIK.job 作成
                    CreKIK()
                Case "02"

                    '布線行の時

                    'キーと値の削除
                    Dic2.Clear()

                    '布線行に関するDictionaryを作成
                    CreDic(Dic2, Area2, AreaName2, Line, Place)


                    SQ = Dic2("線サイズ")       '線サイズ(太さ)(空白付)
                    LnKnd = Dic2("線種")           '線種=電線種(空白付)
                    Color = Dic2("色")             'カラー=線色(空白付)



                    CreRoot()

                    'レコードNo.3　作成
                    'CreRec3(job)

                    '非特殊電線ならレコードNo.4を、特殊電線ならレコードNo.6を作成
                    Select Case SPCCheck(SQ.Trim, LnKnd.Trim)

                        Case "非特殊電線"
                            'レコードNo.3,4のデータ作成
                            CreRec3_4_Data()
                        Case "特殊電線"
                            'レコードNo.6　作成
                            CreRec6_Data()
                    End Select

                    'レコードNo.7のデータを作成
                    CreRec7_Data()

                    'レコードNo.9作成
                    CreRec9()
                    CreHON()
            End Select
        Next

        If ErrEndFlg = True Then
            'エラーがあれば
            MainLoop = False
        Else
            'エラーがなければ
            MainLoop = True
        End If
    End Function
#End Region






    ''' <summary>
    ''' アルファべットと数字を分けた配列を返す
    ''' </summary>
    ''' <param name="str">対象文字列</param>
    ''' <returns>(0):アルファべット,(1):数字　　</returns>
    ''' <remarks></remarks>
    Function SepaStr(ByVal str As String) As String()
        Dim temp As String
        Dim ret(1) As String
        For i = 1 To str.Length
            temp = Mid(str, i, 1)
            Select Case temp
                Case "a" To "z", "A" To "Z"
                    ret(0) = ret(0) & temp
                Case "0" To "9"
                    ret(1) = ret(1) & temp
            End Select
        Next

        Return ret
    End Function


End Module
