Imports System
Imports System.Collections.Generic



Module Sort

    'Public Function Sort_ArrayList(ByRef ArrList As ArrayList)


    '    Dim rtn As New ArrayList

    '    Array.Sort(ArrList, AddressOf Sort.CompareArrStr)
    'End Function

    Function CompareArrStr(ByVal x As String(), _
                            ByVal y As String()) As Integer

        Dim i As Integer, max As Integer = UBound(x)
        Dim rtn As Integer
        For i = 0 To max

            If x(i) <> y(i) Then

                Exit For
            End If
        Next i

        If i > max Then i = max

        Dim xi As Integer, yi As Integer
        If Integer.TryParse(x(i), xi) = True And Integer.TryParse(y(i), yi) = True Then

            If xi = yi Then
                rtn = 0
            ElseIf xi > yi Then
                rtn = 1
            Else
                rtn = -1
            End If


        Else
            rtn = String.Compare(x(i), y(i))

        End If

        Return rtn

    End Function

    Function Sort_Dic(ByVal Dic As Dictionary(Of String, String)) As Dictionary(Of String, String)

        '戻り値
        Dim SortedDic As New Dictionary(Of String, String) '戻り値


        'Keyのソート
        'Dictionaryの内容をコピーして、List(Of KeyValuePair(Of String, Integer))に変換

        Dim pairs As New List(Of KeyValuePair(Of String, String))(Dic)

        ' List.Sortメソッドを使ってソート
        ' (引数に比較方法を定義したメソッドを指定する)
        pairs.Sort(AddressOf CompareKeyValuePair)

        'KeyValuePairからDictionaryに変換
        For Each TDic In pairs

            '戻り値にデータ追加
            SortedDic.Add(TDic.Key, TDic.Value)
        Next

        'ソート後のDictionaryを返す
        Return SortedDic

    End Function

    'Keyのソート部分(value = string)
    Function CompareKeyValuePair(ByVal x As KeyValuePair(Of String, String), _
                                      ByVal y As KeyValuePair(Of String, String)) As Integer
        ' Keyで比較した結果を返す
        Return String.Compare(x.Key, y.Key)

    End Function


    ''' <summary>
    '''　Keyが文字列、Valueが文字列型のリストのDictionaryのソート
    ''' </summary>
    ''' <param name="Dic">ソートする対象のDictionary</param>
    ''' <remarks>KeyとValueの両方のソートをおこなう</remarks>
    Function Sort_Dic_L(ByVal Dic As Dictionary(Of String, List(Of String))) As Dictionary(Of String, List(Of String))

        Dim SortedDic As New Dictionary(Of String, List(Of String)) '戻り値
        Dim TempList As New List(Of String)        'Valueのソートに使用

        'Keyのソート
        'Dictionaryの内容をコピーして、List(Of KeyValuePair(Of String, Integer))に変換

        Dim pairs As New List(Of KeyValuePair(Of String, List(Of String)))(Dic)

        ' List.Sortメソッドを使ってソート
        ' (引数に比較方法を定義したメソッドを指定する)
        pairs.Sort(AddressOf CompareKeyValuePair_L)


        For Each TDic In pairs
            'Valueのソート
            TempList = TDic.Value
            TempList.Sort()

            '戻り値にデータ追加
            SortedDic.Add(TDic.Key, TempList)
        Next

        'ソート後のDictionaryを返す
        Return SortedDic

    End Function

    'Keyのソート部分(value = list(Of String))
    Function CompareKeyValuePair_L(ByVal x As KeyValuePair(Of String, List(Of String)), _
                                      ByVal y As KeyValuePair(Of String, List(Of String))) As Integer
        ' Keyで比較した結果を返す
        Return String.Compare(x.Key, y.Key)

    End Function



    Function Sort_Dic_A(ByVal Dic As Dictionary(Of String, String())) As Dictionary(Of String, String())

        Dim SortedDic As New Dictionary(Of String, String()) '戻り値


        'Keyのソート
        'Dictionaryの内容をコピーして、List(Of KeyValuePair(Of String, Integer))に変換

        Dim pairs As New List(Of KeyValuePair(Of String, String()))(Dic)

        ' List.Sortメソッドを使ってソート
        ' (引数に比較方法を定義したメソッドを指定する)
        pairs.Sort(AddressOf CompareKeyValuePair_A)

        'KeyValuePairからDictionaryに変換
        For Each TDic In pairs

            '戻り値にデータ追加
            SortedDic.Add(TDic.Key, TDic.Value)
        Next

        'ソート後のDictionaryを返す
        Return SortedDic

    End Function






    'Keyのソート部分(value = string())
    Function CompareKeyValuePair_A(ByVal x As KeyValuePair(Of String, String()), _
                                      ByVal y As KeyValuePair(Of String, String())) As Integer
        ' Keyで比較した結果を返す
        Return String.Compare(x.Key, y.Key)

    End Function

End Module
