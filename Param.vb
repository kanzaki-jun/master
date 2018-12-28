Option Strict On

Imports System.Text

Module Publics

#Region " レコードデータ"

    Public Rec1 As String
    Public Rec2 As New List(Of String)
    Public Rec3 As New List(Of String)
    Public Rec4 As New List(Of String)
    Public Rec5 As New List(Of String)
    Public Rec6 As New List(Of String)
    Public Rec7 As New List(Of String)
    Public Rec9 As New List(Of String)

    Public Rec3_4_Data As New Dictionary(Of String, List(Of String))
    Public SortedRec3_4_Data As New Dictionary(Of String, List(Of String))
    Public Rec6_Data As New Dictionary(Of String, List(Of String))
    Public SortedRec6_Data As New Dictionary(Of String, List(Of String))
    Public Rec7_DataList As New LinkedList(Of String())     '


#End Region


#Region " データベース"

    Public Const UserID As String = "dreamprj"
    Public Const Password As String = "dreamprj"
    Public Const Source As String = "x1svr"

#End Region


    Public Kouban As String                     '工番(デフォルト)
    Public Kouban0 As String                    '0つき工番 ex)X16-280 →　X16-0280
    Public DBKouban As String                        'データベース用の工番
    Public Gunban As String                     '群番(デフォルト)
    Public Gun As String                       '群番(B01ならB 01となるような間にスペースを入れた群番)
    Public Job As String                        'JobNo.(デフォルト)
    Public BanNo As String                      '盤No.
    Public AllBanNo As New List(Of String)
    Public CustomerName As String                    '客先名
    Public Type As String

    Public Const RecordLength As Integer = 158              'レコード長

    Public ReadList As New List(Of String)                      'E3からの出力されたテキストファイルのデータを一つにまとめたもの
    Public ArrCoord() As String                                 '座標名と分割名をまとめた配列。(0)はNothing
    Public Seq As New List(Of String)                           '器具番号まとめ
    Public IndexToArea As New Dictionary(Of String, String)     'インデックスをキー、エリアを値とするDictionary
    Public CSVDataDic As New Dictionary(Of String, String())    'Key:座標名(空白含む)+分割名　Value:配列{略称,形式,端子種類,器具番号}
    'Public ALLCSVDataDic As New Dictionary(Of String, Dictionary(Of String, String())) '上記を工番 群番 盤No. 前背でキーとするDictionary

    Public MarkBand As String           'マークバンド
    Public FB As String                 '前背
    Public SQ As String                 '線サイズ
    Public LnKnd As String              '線種
    Public Color As String              'カラー


    '基本的にLFSignとLTSignは同一
    Public LFSign As String             'From線符号
    Public LTSign As String             'To線符号

    'From
    Public FromCoord As String          '座標名
    Public FromPin As String            '端子番号
    Public FromClr As String            'カラー
    Public FromTemp As String           '仮座標

    'To
    Public ToCoord As String            '座標名
    Public ToPin As String              '端子番号
    Public ToClr As String              'カラー
    Public ToTemp As String             '仮座標


    Public Rec4flg As Boolean = False
    Public Rec6flg As Boolean = False

    Public LastFlg As Boolean = False                                '最終行フラグ
    Public FirstOutFlg As Boolean = False                            '初めて出力したかどうかのフラグ(出力していなければFalse、出力したらTrue)

    Public Const MojiBaseMax = 24                                   '文字板の最大数

    Public NowDate As String = DateTime.Now.ToString("yy-MM-dd/HH:mm:ss")   '現在時刻



    Public ErrEndFlg As Boolean = False                                  'Trueの場合、異常終了。Falseなら正常終了

#Region " パス関係"

    'PCAD内のパス
    Public Const pcadpath As String = "\\nissin-fs.nissin-group.local\share\shisusou\B0.特定部署\01.PCAD\課内プロジェクト\新布線システム\E3出力データ変換\配線革新\配線革新資料"
    '非特殊配線･特殊配線
    Public Const SPCPath As String = pcadpath & "\SPC.csv"

    Public CoordPath As String                                  '読み込む座標名＋分割名のCSVへのパス
    Public CSVDataPath As String                                '座標名＋分割名に対応した略称、形式、端子種類、器具番号のデータがあるCSVへのパス
    Public AreaCSVPath As String                                'エリアとインデックスの対応表へのパス

    'サーバーパス
    Public Const Nas As String = "\\nas-Server.swg.nissin.co.jp\"

    'E3出力ファイルへのパス
    Public E3Path As String                                                                             'E3出力先フォルダパス
    Public Const DefE3Path As String = Nas & "設計データ\00 工番データ\"      '分野、期、工番、群番で区分けする前までのE3出力ファイルへのパス

    '変換データ出力先パス
    Public Const OutPutWDPath As String = Nas & "結線\工番ﾃﾞｰﾀ\E3\WD\"       'WD(JOBNo).DATの出力ファイルパス
    Public Const OutPutNCPath As String = Nas & "結線\工番ﾃﾞｰﾀ\E3\NC\"       'NC(JOBNo).DAT,DEV(JOBNo).DAT,MARK(JOBNo).DAT,MJ(JOBNo).DAT,HON(JOBNo).DATの出力ファイルパス
    Public Const OutPutLISTPath As String = Nas & "結線\工番ﾃﾞｰﾀ\E3\LIST\"   '･ZWD05FLP.JOBNo ,ZWD05HON.JOBNo ,ZWD05KIK.JOBNo ,ZWD05TBP.JOBNo ,ZWD05TEA.JOBNo ,ZWD05TEB.JOBNo ,ZWD05WAT.JOBNoの出力ファイルパス


    '製作情報_工番_群番.xlsxファイルへのパス
    Public ProdInfo As String

#End Region

#Region "製作情報関係"

    'シート名
    Public Const MarkName As String = "工番群番マーク"                                       '上記ファイルのマークバンドの情報があるシートの名前
    Public Const BanArrName As String = "盤配列"                                             '上記ファイルの隣の盤の情報があるシートの名前

    '製作情報読み込み
    Public Mark(,) As String
    Public SideBan(,) As String

    'Dictionary
    Public NextBanDic As New Dictionary(Of String, String)              '隣の盤がどれかを保存しておくDictionary(Key:箱番)
    Public MarkDic As New Dictionary(Of String, String)                 'マークバンド(Key:工番-群番)

#End Region


#Region " 読み込みDic1"
    Public Dic1 As New Dictionary(Of String, String)    '読み込んだヘッダ行に関するDictionary

    Public Area1() As Integer = {2, 4, 3, 10, 4, 4}
    Public AreaName1() As String = {"布線台番号", "ジョブNo", "インデックス", "工番", "群番", "盤No"}

#End Region


#Region " 読み込みDic2"
    Public Dic2 As New Dictionary(Of String, String)    '読み込んだ布線行に関するDictionary
    Public Area2() As Integer = { _
                                        4, 5, 4, 2, 5, _
                                        10, 7, 8, 2, 2, 2, 3, _
                                        10, 7, 8, 2, 2, 2, 3, _
                                        2, 2, _
                                        1, 3, 4, 3, 7, 2, _
                                        1, 3, 4, 3, 7, 2 _
                                  }
    Public AreaName2() As String = { _
                                        "シリーズNo", "線種", "線サイズ", "色", "線長さ", _
                                        "FT線番号", "FT機器座標", "FT器具端子", "FT色別", "FTマーク", "FT端子種", "FT布線アドレス", _
                                        "TT線番号", "TT機器座標", "TT器具端子", "TT色別", "TTマーク", "TT端子種", "TT布線アドレス", _
                                        "逆文字F", "逆文字T", _
                                        "FW曲方向", "FW曲げ位置", "FWバイパス位置", "FW端末長さ", "FW機器座標", "FW接続本数", _
                                        "TW曲方向", "TW曲げ位置", "TWバイパス位置", "TW端末長さ", "TW機器座標", "TW接続本数" _
                                    }

    'Public AreaRec1() As Integer = {3, 1, 3, 4, 4, 2, 2, 2, 8, 2, 24, 30, 10, 6, 3, 3, 34}
#End Region


#Region " SQコード"

    'E3出力データの布線行の線サイズから
    Public SQCord As New Dictionary(Of String, String) From
        {
            {"0.3", "01"},
            {"0.5", "02"},
            {"0.75", "04"},
            {"0.9", "06"},
            {"1.25", "08"},
            {"2", "10"},
            {"3.5", "20"},
            {"5.5", "40"},
            {"8.0", "80"}
        }

#End Region


#Region " 撚りコード"


Public YoriCord As New Dictionary(Of String, String) From
        {
            {"KIV", "02"}
        }


#End Region


#Region " 色コード"

    'E3出力データの布線行の色からレコード３の色コードへ変換するDictionary

    Public ColorCord As New Dictionary(Of String, String) From
        {
            {"B", "01"},
            {"R", "02"},
            {"D", "04"},
            {"Y", "08"},
            {"G", "10"},
            {"W", "20"}
        }

#End Region

#Region "E3 to 自動布線"
    'E3と自動布線での表記の違いの対応表

    '線種
    Public ConvLnKnd As New Dictionary(Of String, String) From
        {
            {"1", "1C-SW"},
            {"2", "2C-SW"},
            {"F", "FUZOK"},
            {"TW", "TWIST"}
        }

        Public ConvSQ As New Dictionary(Of String, String) From
        {
            {"C", "1.2"}
        }

#End Region


#Region "  D A T "


#Region " 共通ヘッダDic"

    Public ComDic As New Dictionary(Of String, Integer) From
        {
            {"ジョブNo", 3},
            {"レコードNo", 1},
            {"シリーズNo", 3},
            {"空白", 3},
            {"子番号", 4},
            {"盤No", 4}
        }

#End Region


#Region " Record Dic1"

    Public RecDic1 As New Dictionary(Of String, Integer) From
        {
            {"ジョブNo", 3},
            {"レコードNo", 1},
            {"シリーズNo", 3},
            {"空白1", 3},
            {"子番号", 4},
            {"盤No", 4},
            {"前背", 2},
            {"SCT", 2},
            {"ゾーン", 2},
            {"工番", 8},
            {"?", 2},
            {"空白2", 24},
            {"客先名", 30},
            {"客先コード", 10},
            {"空白3", 6},
            {"座標数", 3},
            {"マークバンド", 3},
            {"空白4", 48}
        }

#End Region


#Region " Record Dic2"



    Public SetDic2 As New Dictionary(Of String, Integer) From
    {
                {"座標名", 6},
                {"分割名", 1},
                {"略称", 4},
                {"形式", 12},
                {"フレーム", 5},
                {"空白", 3},
                {"端子種類", 2}
    }

    Public Const Set2Area = 33

    Public Set2() As Integer = { _
                                    SetDic2("座標名"), _
                                    SetDic2("分割名"), _
                                    SetDic2("略称"), _
                                    SetDic2("形式"), _
                                    SetDic2("フレーム"), _
                                    SetDic2("空白"), _
                                    SetDic2("端子種類") _
                                }

    Public RecDic2 As New Dictionary(Of String, Object) From
    {
                {"ジョブNo", 3},
                {"レコードNo", 1},
                {"シリーズNo", 3},
                {"空白1", 3},
                {"子番号", 4},
                {"盤No", 4},
                {"前背", 2},
                {"SCT", 2},
                {"ゾーン", 2},
                {"空白2", 4},
                {"1", Set2},
                {"2", Set2},
                {"3", Set2},
                {"空白3", 31}
    }

#End Region


#Region " Record Dic3"

    Public Rec3Cnt As Integer = 0        '作成したデータ数のカウント
    Public Row3Cnt As Integer = 1        '行数のカウント
    Public Line3 As New StringBuilder   '作成したデータを保存しておくオブジェクト

    Public SetDic3 As New Dictionary(Of String, Integer) From
    {
                {"SQコード", 2},
                {"肉厚コード", 2},
                {"撚りコード", 2},
                {"色コード", 2}
    }

    Public Const Set3Area = 8

    Public Set3() As Integer = { _
                                SetDic3("SQコード"), _
                                SetDic3("肉厚コード"), _
                                SetDic3("撚りコード"), _
                                SetDic3("色コード")
                            }

    'SQコード、肉厚コード、撚りコード、色コードは1レコードにつき10個
    Public RecDic3 As New Dictionary(Of String, Object) From
    {
            {"ジョブNo", 3},
            {"レコードNo", 1},
            {"シリーズNo", 3},
            {"空白1", 3},
            {"子番号", 4},
            {"盤No", 4},
            {"前背", 2},
            {"SCT", 2},
            {"ゾーン", 2},
            {"空白2", 8},
            {"1", Set3},
            {"2", Set3},
            {"3", Set3},
            {"4", Set3},
            {"5", Set3},
            {"6", Set3},
            {"7", Set3},
            {"8", Set3},
            {"9", Set3},
            {"10", Set3},
            {"空白3", 46}
    }


#End Region


#Region " Record Dic4"

    Public Rec4Cnt As Integer = 0       '作成したデータ数のカウント
    Public Row4Cnt As Integer = 1       '行数のカウント
    Public Line4 As New StringBuilder   '作成したデータを保存しておくオブジェクト




    Public State4 As String  'データの入力状態("From" or "To")


    Public SetDic4 As New Dictionary(Of String, Integer) From
    {
                {"仮座標", 3},
                {"カラー", 1},
                {"カール", 1},
                {"空白", 1},
                {"端子番号", 8},
                {"ベンチNo", 2}
    }

    Public Const Set4Area = 16

    Public Set4() As Integer = { _
                                SetDic4("仮座標"), _
                                SetDic4("カラー"), _
                                SetDic4("カール"), _
                                SetDic4("空白"), _
                                SetDic4("端子番号"), _
                                SetDic4("ベンチNo")
                            }

    '仮座標,カラー,カール,空白3,端子番号,ベンチNoは１レコードにつき3*2(FromTo)個
    '線符号は3個
    Public RecDic4 As New Dictionary(Of String, Object) From
    {
            {"ジョブNo", 3},
            {"レコードNo", 1},
            {"シリーズNo", 3},
            {"空白1", 3},
            {"子番号", 4},
            {"盤No", 4},
            {"前背", 2},
            {"SCT", 2},
            {"ゾーン", 2},
            {"空白2", 8},
            {"To1", Set4},
            {"From1", Set4},
            {"線符号1", 10},
            {"To2", Set4},
            {"From2", Set4},
            {"線符号2", 10},
            {"To3", Set4},
            {"From3", Set4},
            {"線符号3", 10}
    }

#End Region


#Region " Record Dic5"

    '器具番号はレコード内に10個
    Public RecDic5 As New Dictionary(Of String, Integer) From
    {
            {"ジョブNo", 3},
            {"レコードNo", 1},
            {"シリーズNo", 3},
            {"空白1", 3},
            {"子番号", 4},
            {"盤No", 4},
            {"前背", 2},
            {"9999", 4},
            {"空白2", 4},
            {"器具番号1", 10},
            {"器具番号2", 10},
            {"器具番号3", 10},
            {"器具番号4", 10},
            {"器具番号5", 10},
            {"器具番号6", 10},
            {"器具番号7", 10},
            {"器具番号8", 10},
            {"器具番号9", 10},
            {"器具番号10", 10},
            {"空白3", 30}
    }

#End Region


#Region " Record Dic6"

    Public Rec6Cnt As Integer = 0       '作成したデータ数のカウント
    Public Row6Cnt As Integer = 1       '行数のカウント
    Public Line6 As New StringBuilder   '作成したデータを保存しておくオブジェクト
    Public State6 As String  'データの入力状態("From" or "To")

    Public SetDic6 As New Dictionary(Of String, Integer) From
    {
                {"仮座標", 3},
                {"**", 2},
                {"端子番号", 8}
    }

    Public Const Set6Area = 13

    Public Set6() As Integer = { _
                                SetDic6("仮座標"), _
                                SetDic6("**"), _
                                SetDic6("端子番号")
                            }


    '仮座標,空白3,端子番号は3*2個
    '線符号、カラーは3個
    '器具番号はレコード内に10個
    Public RecDic6 As New Dictionary(Of String, Object) From
    {
            {"ジョブNo", 3},
            {"レコードNo", 1},
            {"シリーズNo", 3},
            {"空白1", 3},
            {"子番号", 4},
            {"盤No", 4},
            {"前背", 2},
            {"SCT", 2},
            {"ゾーン", 2},
            {"空白2", 7},
            {"電線種", 5},
            {"太さ", 3},
            {"線色", 2},
            {"From1", Set6},
            {"To1", Set6},
            {"線符号1", 10},
            {"カラー1", 2},
            {"From2", Set6},
            {"To2", Set6},
            {"線符号2", 10},
            {"カラー2", 2},
            {"From3", Set6},
            {"To3", Set6},
            {"線符号3", 10},
            {"カラー3", 2},
            {"空白3", 3}
    }

#End Region


#Region " Record Dic7"

    Public Rec7Cnt As Integer = 0       '作成したデータ数のカウント
    Public Row7Cnt As Integer = 1       '行数のカウント
    Public Line7 As New StringBuilder
    Public SetDic7 As New Dictionary(Of String, Integer) From
    {
                {"端子No", 4},
                {"文字板", 10}
    }

    Public Const Set7Area = 14

    Public Set7() As Integer = { _
                                SetDic7("端子No"), _
                                SetDic7("文字板")
                            }

    '端子No、文字板は6個
    Public RecDic7 As New Dictionary(Of String, Object) From
    {
            {"ジョブNo", 3},
            {"レコードNo", 1},
            {"シリーズNo", 3},
            {"空白1", 3},
            {"子番号", 4},
            {"盤No", 4},
            {"前背", 2},
            {"9999", 4},
            {"空白2", 8},
            {"形式", 10},
            {"座標", 4},
            {"1", Set7},
            {"2", Set7},
            {"3", Set7},
            {"4", Set7},
            {"5", Set7},
            {"6", Set7},
            {"空白3", 28}
    }

#End Region


#Region " Record Dic9"

    Public Rec9Cnt As Integer = 1

    Public CountRec9 As New Dictionary(Of String, Integer) From
    {
            {"0.3", 0},
            {"0.5", 0},
            {"1.25", 0},
            {"2.0", 0},
            {"3.5", 0},
            {"5.5", 0},
            {"8", 0},
            {"14", 0},
            {"22", 0},
            {"38", 0},
            {"60", 0},
            {"100", 0},
            {"150", 0},
            {"ETC", 0},
            {"1C", 0},
            {"2C", 0},
            {"FUZO", 0},
            {"TOBI", 0}
    }

    Public RecDic9 As New Dictionary(Of String, Integer) From
    {
            {"ジョブNo", 3},
            {"レコードNo", 1},
            {"シリーズNo", 3},
            {"空白1", 3},
            {"子番号", 4},
            {"盤No", 4},
            {"9詰", 6},
            {"FB", 2},
            {"空白2", 4},
            {"0.3", 5},
            {"0.5", 5},
            {"1.25", 5},
            {"2.0", 5},
            {"3.5", 5},
            {"5.5", 5},
            {"8", 5},
            {"14", 5},
            {"22", 5},
            {"38", 5},
            {"60", 5},
            {"100", 5},
            {"150", 5},
            {"ETC", 5},
            {"1C", 5},
            {"2C", 5},
            {"FUZO", 5},
            {"TOBI", 4},
            {"空白3", 9},
            {"空白4", 30}
    }

#End Region


#End Region


#Region "  Z W D "


#Region " FLP"

    Public Const FLP_Header As String = "     NC DATA LIST      "

    Public FLPCnt As Integer = 0

    Public FLPDic As New Dictionary(Of String, Integer) From
    {
        {"KOOBAN=", 14},
        {"KYAKU=", 30},
        {"DATE:", 29},
        {"従番", 9},
        {"JOB=", 3},
        {"空白1", 10},
        {"シリーズNo", 3},
        {"空白2", 4},
        {"GRP---", 6},
        {"BAN---", 9},
        {"FB---", 4},
        {"SCT---", 5},
        {"ZONE---", 4},
        {"LOCATION---", 4},
        {"空白3", 40}
    }

    Public Border() As String = { _
                             "                                                                                                                                    ", _
                             "                                                                                                                          *         ", _
                             "                                                                                                                         **         ", _
                             "                                                                                                                          *         ", _
                             "                                                                                                                          *         ", _
                             "                                                                                                                          *         ", _
                             "                                                                                                                          *         ", _
                             "                                                                                                                          *         "
                             }

    Public RecFLP As New List(Of String)

#End Region


#Region " HON"

    Public RecHON As New List(Of String)

    Public Const HON_Header1 = "***    H O N S U U     ***    "
    Public Const HON_Header2 As String = "(                              )"
    Public Const HON_Header3 As String = "     GROUP   BAN  ZENHAI   0.3   0.5  1.25     2   3.5   5.5     8    14    22    38    60   100   150   ETC     1C    2C  FUZO TOBI"

    Public HONDic As New Dictionary(Of String, Integer) From
    {
        {"KOOBAN=", 9},
        {"空白1", 9},
        {"DATE:", 29},
        {"従番", 9},
        {"JOB=", 3},
        {"空白2", 5},
        {"GROUP", 8},
        {"BAN", 5},
        {"ZENHAI", 6},
        {"0.3", 6},
        {"0.5", 6},
        {"1.25", 6},
        {"2.0", 6},
        {"3.5", 6},
        {"5.5", 6},
        {"8", 6},
        {"14", 6},
        {"22", 6},
        {"38", 6},
        {"60", 6},
        {"100", 6},
        {"150", 6},
        {"ETC", 6},
        {"1C", 7},
        {"2C", 6},
        {"FUZO", 6},
        {"TOBI", 5}
    }


    Public PartHON As New Dictionary(Of String, Integer) From
    {
            {"0.3", 0},
            {"0.5", 0},
            {"1.25", 0},
            {"2.0", 0},
            {"3.5", 0},
            {"5.5", 0},
            {"8", 0},
            {"14", 0},
            {"22", 0},
            {"38", 0},
            {"60", 0},
            {"100", 0},
            {"150", 0},
            {"ETC", 0},
            {"1C", 0},
            {"2C", 0},
            {"FUZO", 0},
            {"TOBI", 0}
    }

    Public TotalHON As New Dictionary(Of String, Integer) From
    {
            {"0.3", 0},
            {"0.5", 0},
            {"1.25", 0},
            {"2.0", 0},
            {"3.5", 0},
            {"5.5", 0},
            {"8", 0},
            {"14", 0},
            {"22", 0},
            {"38", 0},
            {"60", 0},
            {"100", 0},
            {"150", 0},
            {"ETC", 0},
            {"1C", 0},
            {"2C", 0},
            {"FUZO", 0},
            {"TOBI", 0}
    }
#End Region


#Region " KIK"


    Public RecKIK As New List(Of String)
    Public Const KIKRecLen As Integer = 132
    Public Const KIKHeader1 As String = "******    KIKI LIST    *******                                                  "
    Public Const KIKHeader2 As String = " NC NO.   ZAHYO     ZAHYO BUNKATU       KATASIKI        FRAM  RYAKU   NO.       (HAKO CHUKEI)      KIGU NO.   #          SOU    (T) "

    Public KIKHeaderDic As New Dictionary(Of String, Integer) From
    {
        {"DATE:", 29},
        {"従番", 9},
        {"JOB=", 3},
        {"空白1", 1},
        {"空白2", 5},
        {"KOOBAN=", 13},
        {"GRP=", 8},
        {"BAN=", 9},
        {"FB=", 3},
        {"SCT=", 10},
        {"ZONE=", 10},
        {"空白3", 39}
    }

    Public KIKDic As New Dictionary(Of String, Integer) From
    {
        {"空白1", 1},
        {"仮座標", 3},
        {"空白2", 6},
        {"座標名", 10},
        {"分割名", 2},
        {"端子種類", 2},
        {"空白3", 14},
        {"形式", 16},
        {"フレーム", 6},
        {"略称", 8},
        {"No", 10},
        {"CHUKEI", 19},
        {"KIGU_No", 11},
        {"デバイス名", 11},
        {"SOU", 7},
        {"(T)", 4}
    }
#End Region


#Region " TEA & TEB"

#Region "  Header"

    Public Const HeaderTEA As String = "*** TENKAIZU SETTKEI HONSU ***                                                  "
    Public Const HeaderTEB As String = "*** TENKAI HONSU HYO     ***(RETUBAKO CHUKEI)                                   "

    Public TEHeaderDic As New Dictionary(Of String, Integer) From
    {
        {"DATE:", 18},
        {"PAGE=", 3},
        {"空白1", 3},
        {"従番", 9},
        {"JOB=", 3},
        {"空白2", 3},
        {"空白3", 2},
        {"KOOBAN---", 13},
        {"GRP---", 9},
        {"BAN---", 9},
        {"FB---", 4},
        {"SCT---", 4},
        {"ZONE---", 4},
        {"空白4", 38}
    }

#End Region
    'ファイル書き込み用リスト
    Public RecTEA As New List(Of String)
    Public RecTEB As New List(Of String)

    Public RecTEA_Data As New List(Of String)
    Public RecTEB_Data As New List(Of String)

    Public TEADic As New Dictionary(Of String, Object) From
    {
        {"空白1", 1},
        {"From仮座標", 3},
        {"空白2", 2},
        {"From座標", 6},
        {"From分割名", 2},
        {"-", 4},
        {"1", SetTEA},
        {"2", SetTEA},
        {"3", SetTEA},
        {"4", SetTEA},
        {"5", SetTEA},
        {"6", SetTEA},
        {"7", SetTEA},
        {"8", SetTEA}
    }

    Public SetTEA As New Dictionary(Of String, Integer) From
    {
        {"To座標", 6},
        {"To分割名", 1},
        {"本数", 3}
    }

    Public TEBDic As New Dictionary(Of String, Object) From
    {
        {"空白1", 1},
        {"From仮座標", 3},
        {"空白2", 2},
        {"From座標", 6},
        {"From分割名", 2},
        {"-", 4},
        {"1", SetTEB},
        {"2", SetTEB},
        {"3", SetTEB},
        {"4", SetTEB},
        {"5", SetTEB},
        {"空白3", 8}
    }

    Public SetTEB As New Dictionary(Of String, Integer) From
    {
        {"To座標", 6},
        {"To分割名", 1},
        {"To仮座標", 3},
        {"本数", 3}
    }


    Public Const NextSpace As String = "                   "

    '1行以上の領域をオーバーしてしまった場合にたつフラグ
    Public TEAExtenFlg As Boolean = False
    Public TEBExtenFlg As Boolean = False

    Public Const TEAMAX = 8
    Public TEACount As Integer = 0
    Public TEBCount As Integer = 0

    Public LineTEA As New StringBuilder
    Public LineTEB As New StringBuilder
    Public TotalTEACount As Integer
    Public TotalTEBCount As Integer
    Public Const RecTELen As Integer = 132
    Public PageCnt As Integer = 0

    '集計本数を表示する際に使用する文字列
    Public Const TotalTEA = " TOTAL   "
    Public Const TotalTEB = " TOTAL          "


#End Region


#Region " WAT"
    '線サイズ、色、線符号が違うときに挿入する文字列
    Public Const BorderWAT As String = "------------------------------------------------------------------------------------------------------------------------------------"

    Public Const TitleWAT As String = "** WATARI LIST **   "

    Public WATHeader As New Dictionary(Of String, Integer) From
    {
        {"KOOBAN=", 12},
        {"GRP=", 4},
        {"BAN=", 6},
        {"FB=", 3},
        {"SCT=", 12},
        {"DATE:", 30},
        {"従番", 9},
        {"JOB=", 3}
    }

    Public WATDic As New Dictionary(Of String, Integer) From
    {
        {"座標名", 8},
        {"空白1", 1},
        {"線サイズ", 4},
        {"空白2", 1},
        {"色", 2},
        {"空白3", 5},
        {"線符号", 21}
    }

    Public Const PartLen As Integer = 111                       '線符号までのレコードの長さ
    Public Const NextData As Integer = 11                       ' * 端子番号(8)の桁数
    Public Const RecWATLength As Integer = 132                  '線符号を含めたレコード長
    Public Const NextWAT As String = "                     "    '次の行にいくときに使用
    Public LineWAT As New StringBuilder
    Public RecWAT As New List(Of String)
    Public WATCount As Integer = 0

    Public WATDataDic As New Dictionary(Of String, List(Of String))
#End Region


#End Region


End Module
