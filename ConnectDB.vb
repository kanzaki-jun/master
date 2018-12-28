
Imports System
Imports System.Data
Imports Oracle.DataAccess.Client

Module ConnectDB


    Public Msg As String
    Public MsgTitle As String

    Public cn As OracleConnection = Nothing

    ''' 
    ''' データベース接続
    ''' 
    ''' 
    Public Sub ConnectODB(ByVal DBK As String)

        Dim ConnectStr As String



        Try
            'ユーザー名、パスワード、接続文字列を入れた文字列を作成(Globalにて定義)
            ConnectStr = "User ID=" & UserID & ";" & "password=" & Password & ";" _
                                & "Data Source=" & Source & ";"

            Try

                '接続オブジェクトを作成
                cn = New OracleConnection(ConnectStr)
                '実際にデータベースに接続
                cn.Open()



                'SQL文
                Dim cmdQuery As String = "SELECT HANNYUU_MEI" & " " & _
                                         "FROM DREAMPRJ.DAINITTEI_A" & " " & _
                                         "WHERE KOUBAN=" & "'" & DBK & "'"

                Dim cmd2Query As String = "SELECT FUNJYUYOUKAKUBUN5(ss_tantou_ss||ss_tantou_ka,kouban) 区分" & " " & _
                            "FROM DREAMPRJ.DAINITTEI_A" & " " & _
                            "WHERE KOUBAN=" & "'" & DBK & "'"


                ' Create the OracleCommand object
                Dim cmd As OracleCommand = New OracleCommand(cmdQuery)

                cmd.Connection = cn
                cmd.CommandType = CommandType.Text

                Dim cmd2 As OracleCommand = New OracleCommand(cmd2Query)
                cmd2.Connection = cn
                cmd2.CommandType = CommandType.Text

                Try
                    ' Execute command, create OracleDataReader object
                    Dim reader As OracleDataReader = cmd.ExecuteReader()

                    While (reader.Read())

                        '30byteだけ取り込む

                        CustomerName = MidB(TrimCus(reader.GetString(0)), 1, 30)

                    End While


                    ' Execute command, create OracleDataReader object
                    Dim reader2 As OracleDataReader = cmd2.ExecuteReader()
                    While (reader2.Read())

                        Type = reader2.GetString(0).Trim

                    End While


                Catch ex As Exception

                    Console.WriteLine(ex.Message)

                Finally
                    '例外が発生するしないに関わらず実行
                    ' Dispose OracleCommand object
                    cmd.Dispose()

                    ' Close and Dispose OracleConnection object
                    cn.Close()
                    cn.Dispose()
                End Try

            Catch ex As OracleException
                Msg = ex.Message
                MsgTitle = "データベース接続エラーです(Oracle)"

            End Try


        Catch ex As Exception
            Msg = ex.Message
            MsgTitle = "データベース接続エラーです"

        End Try

        ConnectionClose()


    End Sub

    ''' 
    ''' データベース切断
    ''' 
    ''' 
    Private Sub ConnectionClose()

        Try
            If Not IsNothing(cn) Then
                If cn.State = ConnectionState.Open Then
                    cn.Close()

                End If
                cn.Dispose()
            End If

        Catch ex As OracleException
            Msg = ex.Message
            MsgTitle = "データベース切断エラーです"


        End Try
    End Sub

    Public Function TrimCus(ByVal Cus As String) As String

        '半角にして、前後の空白削除
        Cus = StrConv(Trim(Cus), VbStrConv.Narrow)

        While InStr(Cus, "  ") <> 0
            Cus = Replace(Cus, "  ", " ")
        End While

        TrimCus = Cus
    End Function


    ''' -----------------------------------------------------------------------------------------
    ''' <summary>
    '''     文字列の指定されたバイト位置から、指定されたバイト数分の文字列を返します。</summary>
    ''' <param name="stTarget">
    '''     取り出す元になる文字列。</param>
    ''' <param name="iStart">
    '''     取り出しを開始する位置。</param>
    ''' <param name="iByteSize">
    '''     取り出すバイト数。</param>
    ''' <returns>
    '''     指定されたバイト位置から指定されたバイト数分の文字列。</returns>
    ''' -----------------------------------------------------------------------------------------
    Public Function MidB _
    (ByVal stTarget As String, ByVal iStart As Integer, ByVal iByteSize As Integer) As String
        Dim hEncoding As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
        Dim btBytes As Byte() = hEncoding.GetBytes(stTarget)
        Dim size As Integer = 0

        size = iByteSize

        If iByteSize > stTarget.Length Then
            If iByteSize = stTarget.Length + 1 And _
                LenB(Right(stTarget, 1)) = 2 Then
                '元の文字列のバイト数が取り出すバイト数+1の時
                'かつ、元の文字列の最後の文字が２バイト文字の場合
                'オーバーしている最後の文字まで返してしまうので返すバイト数を-1
                size = iByteSize - 1
            Else
                size = LenB(stTarget)
            End If

        End If

        Dim str As String = hEncoding.GetString(btBytes, iStart - 1, size)
        Return str
    End Function

     ''' -----------------------------------------------------------------------------------------
    ''' <summary>
    '''     半角 1 バイト、全角 2 バイトとして、指定された文字列のバイト数を返します。</summary>
    ''' <param name="stTarget">
    '''     バイト数取得の対象となる文字列。</param>
    ''' <returns>
    '''     半角 1 バイト、全角 2 バイトでカウントされたバイト数。
    '''     ""の場合、0を返す
    ''' </returns>
    ''' -----------------------------------------------------------------------------------------
    Public Function LenB(ByVal stTarget As String) As Integer
        Return System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(stTarget)
    End Function


End Module
