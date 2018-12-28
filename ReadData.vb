Option Strict On

Imports System.IO
Imports System.Text


Imports ClosedXML.Excel

Module ReadData






#Region "CSV読み込み"


    ''' <summary>
    ''' CSVを読み込む
    ''' </summary>
    ''' <param name="csvFilePath" >CSVファイルパス</param>
    ''' <param name="Trimflg" >読み込んだデータのスペースを削除するかどうかのフラグ</param>
    ''' <returns >読み込んだ結果を二次元オブジェクトで返す</returns>
    ''' <remarks>参考　https://dobon.net/vb/dotnet/file/ReadCSVfile.html　</remarks>

    Function ReadCSV(ByVal csvFilePath As String, ByVal Trimflg As Boolean) As Object

        Dim csvRecords As New System.Collections.ArrayList()



        'Shift JISでCSVをTextFieldParserで読み込む
        'System.Text.Encoding.GetEncoding(932)でエンコーディングオブジェクトを取得
        'コードページIDについては下記URL参照
        'https://dobon.net/vb/dotnet/string/getencodingobject.html#section2

        Dim tfp As New FileIO.TextFieldParser(csvFilePath, _
            System.Text.Encoding.GetEncoding(932))


        'フィールドが文字で区切られているとする
        'デフォルトでDelimitedだが、念のため指定
        tfp.TextFieldType = FileIO.FieldType.Delimited

        '区切り文字を,とする
        tfp.Delimiters = New String() {","}

        'フィールドを"で囲み、改行文字、区切り文字を含めることができるか
        'デフォルトでtrueだが、念のため指定
        tfp.HasFieldsEnclosedInQuotes = True

        'フィールドの前後からスペースを削除する設定
        '第2引数の結果から判断する。
        tfp.TrimWhiteSpace = Trimflg

        While Not tfp.EndOfData
            'フィールドを配列として読み込む
            Dim fields As String() = tfp.ReadFields()
            'ArrayListに保存
            csvRecords.Add(fields)
        End While

        '後始末
        tfp.Close()

        ReadCSV = csvRecords
    End Function


#End Region


#Region "XLSX読み込み"

    Function ReadXLSX(ByVal Path As String, ByVal SheetName As String) As String(,)

        Dim Arr(,) As String
        Using wkbk As New ClosedXML.Excel.XLWorkbook(Path)
            'シート名を指定してシートを取得
            Dim sheet As ClosedXML.Excel.IXLWorksheet = wkbk.Worksheet(SheetName)

            '使用されている最終セルの位置を取得
            Dim address = sheet.RangeUsed().RangeAddress.LastAddress
            '最終セルの位置からから最終行、最終列を取得
            Dim row As Integer = address.RowNumber : Dim col As Integer = address.ColumnNumber

            ReDim Arr(row - 1, col - 1)
            For i As Integer = 1 To row
                For j As Integer = 1 To col
                    'Dim tempAd As ClosedXML.Excel.IXLAddress = sheet.Cell(i, j).Address
                    '任意のセルを取得
                    Dim tempVal As ClosedXML.Excel.IXLCell = sheet.Cell(i, j)
                    'Console.WriteLine(tempAd.ColumnLetter & tempAd.RowNumber & "=" & CStr(tempVal.Value))

                    Arr(i - 1, j - 1) = CStr(tempVal.Value)
                Next j
            Next i
            ReadXLSX = Arr
        End Using
    End Function

    Public Function LoadXLSX() As Boolean


        If System.IO.File.Exists(ProdInfo) = False Then

            'エリア情報CSVファイルが存在しない場合
            Console.WriteLine("製作情報の取得に失敗しました。")
            ErrEndFlg = True

            LoadXLSX = False
        Else
            Console.WriteLine("製作情報の読み込みを開始")

            '製作情報ファイルからデータを読み込み、Publicの配列にデータをセット
            SideBan = ReadXLSX(ProdInfo, BanArrName)   '盤配列シートのデータ読み込む
            Mark = ReadXLSX(ProdInfo, MarkName)        '工番群番マークシートのデータ読み込み

            Console.WriteLine("製作情報の読み込みを終了")


            Console.WriteLine("読み込んだ製作情報からデータを作成　開始")

            '隣の盤がどの盤かを示す辞書型のNextBanDicを作成
            CreNextBanDic()
            '工番-群番をキー、HOTもしくはFMKが値の辞書型のMarkDicを作成
            CreMarkDic()

            Console.WriteLine("読み込んだ製作情報からデータを作成　終了")

            LoadXLSX = True

        End If

    End Function


    Public Sub CreMarkDic()
        Dim i As Integer, j As Integer
        Dim KouPos As Integer = 0, GunPos As Integer = 0
        Dim key As String = ""
        Dim imax As Integer = UBound(mark, 1)
        Dim jmax As Integer = UBound(mark, 2)

        For i = 0 To imax
            For j = 0 To jmax
                If i = 0 Then
                    Select Case Mark(i, j)

                        Case "工番"
                            KouPos = j
                        Case "群番"
                            GunPos = j
                    End Select
                Else
                    If Mark(i, j) = "HOT" Or Mark(i, j) = "FMK" Then
                        key = Mark(i, KouPos) & "-" & Mark(i, GunPos)
                        If MarkDic.ContainsKey(key) = False Then
                            MarkDic.Add(key, Mark(i, j))
                        End If
                    End If
                End If


            Next j
        Next i


    End Sub



''' <summary>
''' 現在の盤の隣の盤を示すNextBanDicを作成する
''' </summary>
''' <remarks></remarks>
    Public Sub CreNextBanDic()
        Dim i As Integer, j As Integer
        Dim iMax As Integer = UBound(SideBan, 1), jMax As Integer = UBound(SideBan, 2)
        Dim BanNoP As Integer  '盤No.位置
        Dim ChukeiP As Integer '中継端子台取付位置
        Dim Threshold As Integer
        Dim NextBan As String

        Dim Ban As String                   '値:盤No
        Dim Chukei As String                '値:中継端子台取付位置
        Dim LRFlg As Boolean = False    'Chukeiの値がL、RではないときTrue



        For j = 0 To jMax
            If SideBan(0, j) = "盤No" Then BanNoP = j
            If SideBan(0, j) = "中継端子台取付位置" Then ChukeiP = j
        Next j

        For i = 1 To iMax

            Ban = SideBan(i, BanNoP).Trim
            Chukei = SideBan(i, ChukeiP).Trim
            Select Case Chukei
                Case "L"
                    Threshold = i - 1
                Case "R"
                    Threshold = i + 1
                Case Else
                    'LR等の場合
                    LRFlg = True
            End Select

            '最初の盤の中継端子台取付位置がLか,
            '最後の盤の中継端子台取付位置がRか, 中継端子台取付位置がLRの場合
            If Threshold = 0 Or Threshold = iMax + 1 Or LRFlg = True Then
                '隣の盤のデータをなしにする
                NextBan = ""
                LRFlg = False
            Else
                '隣の盤データを保存
                NextBan = SideBan(Threshold, BanNoP)
            End If

            NextBanDic.Add(Ban, NextBan)
        Next i


    End Sub


#End Region

''' <summary>
''' 
''' </summary>
''' <param name="SQ">線サイズ</param>
''' <param name="LnKnd">線種</param>
''' <returns></returns>
''' <remarks></remarks>
    Function SPCCheck(ByVal SQ As String, ByVal LnKnd As String) As String
        Dim imax As Integer, jmax As Integer
        Dim i As Integer, j As Integer

        Dim Arl As ArrayList
        'csvをArrayListに変換
        Arl = CType(ReadCSV(SPCPath, True), System.Collections.ArrayList)
        'ArrayListをジャグ配列に変換
        Dim jag()() As String = CType(Arl.ToArray(GetType(String())), String()())

        imax = jag.Length

        '線種どの位置かを調べる
        For i = 0 To imax - 1
            If LnKnd.Trim = jag(i)(0) Then
                Exit For
            End If
        Next

        If i > imax - 1 Then
            Console.WriteLine("線種が見つかりませんでした")
            SPCCheck = ""
            Exit Function
        End If

        jmax = jag(0).Length

        '線サイズがどの位置かを調べる
        For j = 0 To jmax - 1
            If CStr(SQ) = jag(0)(j) Then
                Exit For
            End If
        Next

        If j > jmax - 1 Then
            Console.WriteLine("線サイズが見つかりませんでした")
            SPCCheck = ""
            Exit Function
        End If

        Select Case jag(i)(j)

            Case "○"
                'Console.WriteLine("非特殊電線")
                SPCCheck = "非特殊電線"
            Case "×"
                'Console.WriteLine("特殊電線")
                SPCCheck = "特殊電線"
            Case "-"
                MsgBox("電線種=" & LnKnd & ",線サイズ=" & SQ)
                SPCCheck = ""
            Case Else
                SPCCheck = ""
        End Select
    End Function




    ''' <summary>
    ''' 引数のArrayListを
    ''' Key:座標名＋分割名
    ''' Value:配列{略称,形式,端子種類,器具番号}　とするDictionaryに変換
    ''' </summary>
    ''' <param name="csvRecords">ArryList</param>
    ''' <remarks></remarks>
    Sub ParseCSVData(ByRef csvRecords As System.Collections.ArrayList)

        Dim arr() As String
        Dim key As String = ""
        Dim value() As String
        Dim space As String = ""
        Dim divide As String = ""

        For i As Integer = 1 To csvRecords.Count - 1
            '(0)=座標名、(1)=分割名、(2)=略称、(3)=形式、(4)=端子種類、(5)=器具番号
            arr = CType(csvRecords(i), String())

            space = CreSpace(SetDic2("座標名") - arr(0).Length, " ")

            If arr(1) = "" Then
                divide = " "
            Else
                divide = arr(1)
            End If

            'キー:座標名＋分割名
            key = arr(0).Trim & space & divide

            '値:配列{略称,形式,端子種類,器具番号}
            value = {arr(2), arr(3), arr(4), arr(5)}

            If CSVDataDic.ContainsKey(key) = False Then
                CSVDataDic.Add(key, value)
                Seq.Add(arr(5))

            End If
        Next

        key = Kouban & " " & Gunban & " " & BanNo & " " & FB
        'ALLCSVDataDic.Add(key, CSVDataDic)
    End Sub

    ''' <summary>
    ''' 座標名＋分割名を保存しておく配列ArrCoordを作成
    ''' </summary>
    ''' <param name="csvRecords"></param>
    ''' <remarks></remarks>

    Sub CreCoord(ByRef csvRecords As System.Collections.ArrayList)

        Dim arr() As String             'オブジェクトの変換用
        Dim Coordinate As String = ""   '座標
        Dim Divide As String = ""       '分割名
        Dim Space As String             '空白

        ReDim ArrCoord(csvRecords.Count - 1)

        '座標、ピングループ=分割名のデータを抜き出し、配列に保存
        '(0)X座標(1)Y座標(2)配線端番号(3)座標(4)ピン高さ(5)ピングループ
        For i As Integer = 1 To csvRecords.Count - 1
            'ObjectをString配列に変換
            arr = CType(csvRecords(i), String())
            '座標、ピングループのデータを取得
            Coordinate = arr(3) : Divide = arr(5)

            Space = CreSpace(SetDic2("座標名") - Coordinate.Length, " ")

            If Divide = "" Then Divide = " "

            ArrCoord(i) = Coordinate & Space & Divide
        Next

        '重複する要素を削除
        '参考 https://dobon.net/vb/dotnet/programing/arraydistinct.html

        'HashSetクラスを宣言
        Dim HashSet As New System.Collections.Generic.HashSet(Of String)(ArrCoord)

        '重複がなくなったHashSetのカウントを利用して配列再定義
        ReDim ArrCoord(HashSet.Count - 1)

        '配列にHashSetの要素をコピー
        HashSet.CopyTo(ArrCoord)
    End Sub







    ''' <summary>
    ''' Dictionaryを作成する
    ''' </summary>
    ''' <param name="Dic">作成するDictionary</param>
    ''' <param name="area">フィールド</param>
    ''' <param name="areaname">フィールド名</param>
    ''' <param name="line">読み込まれた1行</param>
    ''' <param name="place">開始位置</param>
    ''' <remarks></remarks>
    ''' 
    Public Sub CreDic(ByVal Dic As Dictionary(Of String, String), _
                      ByRef area() As Integer, ByRef areaname() As String, _
                      ByVal line As String, ByRef place As Integer)

        Dim i As Integer    'For文カウンタ

        For i = 0 To area.Length - 1
            '読み込まれた1行をplace=開始位置からエリアの分だけ切り取った文字列を
            '値とするDictionaryを作成する
            Dic.Add(areaname(i), Mid(line, place, area(i)))
            '開始位置を更新
            place = place + area(i)
        Next

    End Sub

    ''' <summary>
    ''' 第1引数で指定された数だけ第2引数の文字を追加した文字列を返す
    ''' </summary>
    ''' <param name="cnt">返す文字列の長さ</param>
    ''' <param name="tag" >文字あるいは空白</param>
    ''' <returns>文字列</returns>
    ''' <remarks>cnt = 0ならば戻り値=""</remarks>

    Public Function CreSpace(ByVal cnt As Integer, ByVal tag As String) As String

        '返す文字列の長さが０の場合は戻り値=""として終了
        If cnt <= 0 Then
            CreSpace = ""
            Exit Function
        End If

        '参考
        'https://dobon.net/vb/dotnet/string/stringbuilder.html

        'キャパシティ指定のStringBuilderを宣言
        Dim str As New StringBuilder(cnt)


        For i = 1 To cnt
            '第2引数を第1引数の数だけ追加
            str.Append(tag)
        Next

        'String型に変換
        CreSpace = str.ToString
    End Function





End Module
