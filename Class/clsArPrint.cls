VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'********************************************************************************
'クラス名　：ActiveReport1.5 用クラス(ADOバージョン)
'
'条　件　　：１．参照設定  ActiveReports Viewer Control 1.5J
'　　　　　　　　　　　　　Data Dynamics ActiveReports 1.5J
'　　　　　　　　　　　　　ActiveReports PDF Export Fileter
'　　　　　　　　　　　　　ActiveReports PDF Export Fileter Fixer
'
'           ２．データコントロール名をDataControl1にすること
'
'作成者　　：株式会社 コム・エンジニアリング　渥美
'バージョン：1.00
'作成年月日：２００２／０５／２４
'更新履歴　：
'********************************************************************************

Private Const INI_FILENAME = "ArPrint.ini"
Private m_strIniFileName As String

Private Type PrintInfo
    Name                    As String           'レポート名称
    'objReport               As Object           'ActiveReportオブジェクト
    objReport               As ActiveReport     'ActiveReportオブジェクト
    DatabaseName            As String           'データベース名
    Connection              As String           'データベースコネクション
    SQL                     As String           'SQL句
    Caption                 As String           'プレビューウインドウのキャプション
    TableForExcel           As String           'Excel用ワークテーブル名
    PrnDefaultPrinter       As Boolean          'True:通常使うプリンタ
    PrnDeviceName           As String           'プリンタ名
    PrnPaperSize            As Integer          '用紙サイズ
    PrnOrientation          As Integer          '用紙の向き
    PrnPageTopMargin        As Integer          '余白（上）
    PrnPageBottomMargin     As Integer          '余白（下）
    PrnPageLeftMargin       As Integer          '余白（左）
    PrnPageRightMargin      As Integer          '余白（右）
    NoData                  As Boolean          'データ有無
    PrnPaperSource          As Integer          '給紙方法
End Type
Private m_PrintInfo As PrintInfo

'INIファイルのAPI宣言
Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long

'[ﾌｧｲﾙを開く]と[ﾌｧｲﾙを保存]ﾀﾞｲｱﾛｸﾞﾎﾞｯｸｽの為の構造体宣言
Private Type OPENFILENAME
    lStructSize As Long         '構造体のｻｲｽﾞ
    hwndOwner As Long           'ﾀﾞｲｱﾛｸﾞﾎﾞｯｸｽを所有するｳｨﾝﾄﾞｳのﾊﾝﾄﾞﾙ
    hInstance As Long           'ｱﾌﾟﾘｹｰｼｮﾝ ｲﾝｽﾀﾝｽ
    lpstrFilter As String       'ﾌｨﾙﾀｰ
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long        'ﾃﾞﾌｫﾙﾄのﾌｨﾙﾀ
    lpstrFile As String         '選択されたﾌｧｲﾙ名
    nMaxFile As Long            'ﾌｧｲﾙ名の最大長
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long               'ｵﾌﾟｼｮﾝ
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
'[ﾌｧｲﾙの保存]ﾀﾞｲｱﾛｸﾞﾎﾞｯｸｽを呼び出すAPI
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'OPENFILENAME構造体のflagsで指定する定数
Private Const OFN_READONLY = &H1                '[読み取り専用]ﾁｪｯｸﾎﾞｯｸｽをﾁｪｯｸ
Private Const OFN_OVERWRITEPROMPT = &H2         'ﾌｧｲﾙが存在していた場合、上書きを問い合わせる
Private Const OFN_HIDEREADONLY = &H4            '[読み取り専用]ﾁｪｯｸﾎﾞｯｸｽを非表示
Private Const OFN_SHOWHELP = &H10               '[ﾍﾙﾌﾟ]ﾎﾞﾀﾝの表示
Private Const OFN_ALLOWMULTISELECT = &H200      '複数のﾌｧｲﾙを選択可能に
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800         '存在しないﾊﾟｽ名を入力不可に
Private Const OFN_FILEMUSTEXIST = &H1000        '存在しないﾌｧｲﾙ名を入力不可に
Private Const OFN_CREATEPROMPT = &H2000         'ﾌｧｲﾙが存在しなかった場合、新規作成するかどうか表示
Private Const OFN_EXPLORER = &H80000

'目　的　　：クラス初期化
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub Class_Initialize()

    On Error GoTo Class_Initialize_Err

    'INIファイル名設定
    If right(Trim(App.Path), 1) = "\" Then
        m_strIniFileName = App.Path & INI_FILENAME
    Else
        m_strIniFileName = App.Path & "\" & INI_FILENAME
    End If
    
    'プロパティの初期化
    Set m_PrintInfo.objReport = Nothing
    
    '標準のプロパティ値を設定
    With m_PrintInfo
        .Name = ""
        .DatabaseName = ""
        .Connection = ""
        .SQL = ""
        .Caption = ""
        .TableForExcel = ""
        .PrnDefaultPrinter = True
        .PrnDeviceName = ""
        .PrnPaperSize = vbPRPSA4
        .PrnOrientation = vbPRORLandscape
        .PrnPageTopMargin = 10
        .PrnPageBottomMargin = 10
        .PrnPageLeftMargin = 10
        .PrnPageRightMargin = 10
        .PrnPaperSource = 0
        .NoData = False
    End With
    
    Exit Sub
    
Class_Initialize_Err:
    
   Call MsgBox("クラス初期化エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Class_Initialize_Err")
    
End Sub

'目　的　　：クラス終了時
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Sub Class_Terminate()
    
    '処理なし

End Sub

'目　的　　：ActiveReportの出力処理
'条　件　　：データコントロール名をDataControl1にすること
'結　果　　：
'引　数　　：0:印刷プレビュー 1:印刷 2:PDF出力
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function PrintActiveReport(intDestination As Integer, Optional strPdfFileName As String) As Boolean
    
    On Error GoTo PrintActiveReport_Err
    
    PrintActiveReport = False
    
    'プロパティ更新
    If Refresh() = False Then Exit Function
    
    Select Case intDestination
        Case 0:     '印刷プレビュー
            'マウスポインタを砂時計に変更
            Screen.MousePointer = vbHourglass
        
            '新しいレポートを､ARViewerコントロールに割り当てます
            Set frmARPreview.arv.ReportSource = m_PrintInfo.objReport
            frmARPreview.Caption = m_PrintInfo.Caption
            Set frmARPreview.m_objArPrint = Me
            m_PrintInfo.objReport.Run
            DoEvents
            
            'マウスポインタを元に戻す
            Screen.MousePointer = vbDefault
            If m_PrintInfo.NoData = False Then
                frmARPreview.Show vbModal
            End If
        Case 1:     '印刷ダイアログを表示せずに印刷
            m_PrintInfo.objReport.PrintReport False
        Case 2:     'PDFファイルへ出力
            '出力先指定
            If IsMissing(strPdfFileName) Or strPdfFileName = "" Then
                strPdfFileName = Get_FileName(0)
                If Trim(strPdfFileName) = "" Then
                    Exit Function
                End If
            End If
        
            'マウスポインタを砂時計に変更
            Screen.MousePointer = vbHourglass
        
            Dim objARExport As New ARExportPDF
            Dim objARExportFix As New ARExportPDFFix
            objARExport.FileName = strPdfFileName
            m_PrintInfo.objReport.Run
            m_PrintInfo.objReport.Export objARExport
            'PDFファイルの文字化けを修正
            objARExportFix.ExportFix objARExport.FileName
    
            'マウスポインタを元に戻す
            Screen.MousePointer = vbDefault
    End Select

    PrintActiveReport = True

    Exit Function

PrintActiveReport_Err:

    PrintActiveReport = False
    'マウスポインタを元に戻す
    Screen.MousePointer = vbDefault
    Call MsgBox("ActiveReportの出力処理エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "PrintActiveReport_Err")

End Function

'目　的　　：プロパティ更新
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function Refresh() As Boolean
    
    On Error GoTo Refresh_Err
    
    Screen.MousePointer = vbHourglass
    
    Refresh = False
    
    Call SetIni(m_PrintInfo.Name, "DefaultPrinter", m_PrintInfo.PrnDefaultPrinter)
    If m_PrintInfo.PrnDefaultPrinter = True Then
        m_PrintInfo.PrnDeviceName = ""
    End If
    Call SetIni(m_PrintInfo.Name, "DeviceName", m_PrintInfo.PrnDeviceName)
    Call SetIni(m_PrintInfo.Name, "PaperSize", m_PrintInfo.PrnPaperSize)
    Call SetIni(m_PrintInfo.Name, "Orientation", m_PrintInfo.PrnOrientation)
    Call SetIni(m_PrintInfo.Name, "PageTopMargin", m_PrintInfo.PrnPageTopMargin)
    Call SetIni(m_PrintInfo.Name, "PageBottomMargin", m_PrintInfo.PrnPageBottomMargin)
    Call SetIni(m_PrintInfo.Name, "PageLeftMargin", m_PrintInfo.PrnPageLeftMargin)
    Call SetIni(m_PrintInfo.Name, "PageRightMargin", m_PrintInfo.PrnPageRightMargin)
    Call SetIni(m_PrintInfo.Name, "PrnPaperSource", m_PrintInfo.PrnPaperSource)
    
    With m_PrintInfo.objReport
        'データベース設定
        .DataControl1.ConnectionString = m_PrintInfo.Connection
        DoEvents
        If Trim(m_PrintInfo.SQL) <> "" Then
            .DataControl1.Source = m_PrintInfo.SQL
        End If
        .DataControl1.Refresh
        DoEvents
        'プリンタ設定
        On Error GoTo Printer_Setup
        If Trim(m_PrintInfo.PrnDeviceName) <> "" Then .Printer.DeviceName = m_PrintInfo.PrnDeviceName
        On Error GoTo Refresh_Err
        .Printer.PaperSize = m_PrintInfo.PrnPaperSize
        .Printer.Orientation = m_PrintInfo.PrnOrientation
        If m_PrintInfo.PrnPaperSource <> 0 Then
            .Printer.PaperBin = m_PrintInfo.PrnPaperSource
        End If
        '余白設定
        .PageTopMargin = mm_To_Twips(m_PrintInfo.PrnPageTopMargin)
        .PageBottomMargin = mm_To_Twips(m_PrintInfo.PrnPageBottomMargin)
        .PageLeftMargin = mm_To_Twips(m_PrintInfo.PrnPageLeftMargin)
        .PageRightMargin = mm_To_Twips(m_PrintInfo.PrnPageRightMargin)
        'ドキュメント名
        .documentName = m_PrintInfo.Caption
    End With

    Screen.MousePointer = vbDefault

    Refresh = True

    Exit Function

Printer_Setup:

    Screen.MousePointer = vbDefault
    Call MsgBox("無効なプリンタが設定されています。" & vbCrLf & "プリンタの設定を行ってください。", vbOKOnly + vbCritical, "エラー")
    
    'ページ設定ダイアログの表示
    m_PrintInfo.PrnDeviceName = ""
    Set frmARPage.m_objArPrint = Me
    frmARPage.Show vbModal
    Refresh = Not frmARPage.m_blnCanselFlg
    
    Exit Function

Refresh_Err:

    Refresh = False
    Screen.MousePointer = vbDefault
    Call MsgBox("Refreshエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Refresh_Err")

End Function

'目　的　　：ミリからTwipsへの変換
'条　件　　：内部関数
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Function mm_To_Twips(intmm As Integer) As Integer
    
    Dim intBuff1 As Integer
    
    On Error Resume Next
    
    '※1ｃｍ＝567Twips 1mm＝56.7Twips
    intBuff1 = intmm * 56.7
    mm_To_Twips = Fix(intBuff1)

End Function

'目　的　　：設定読み込み
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Function ReadPrnIni(strName As String) As Boolean
    
    On Error GoTo ReadPrnIni_Err
    
    If GetIni(strName, "DefaultPrinter") <> "" Then
        If GetIni(strName, "DefaultPrinter") = "True" Then
            m_PrintInfo.PrnDefaultPrinter = True
        ElseIf GetIni(strName, "DefaultPrinter") = "False" Then
            m_PrintInfo.PrnDefaultPrinter = False
        Else
            m_PrintInfo.PrnDefaultPrinter = True
        End If
    Else
        m_PrintInfo.PrnDefaultPrinter = True
    End If
    If m_PrintInfo.PrnDefaultPrinter = True Then
        m_PrintInfo.PrnDeviceName = Printer.DeviceName
    Else
        m_PrintInfo.PrnDeviceName = GetIni(strName, "DeviceName")
    End If
    If GetIni(strName, "PaperSize") <> "" Then
        m_PrintInfo.PrnPaperSize = GetIni(strName, "PaperSize")
    End If
    If GetIni(strName, "Orientation") <> "" Then
        m_PrintInfo.PrnOrientation = GetIni(strName, "Orientation")
    End If
    If GetIni(strName, "PageTopMargin") <> "" Then
        m_PrintInfo.PrnPageTopMargin = GetIni(strName, "PageTopMargin")
    End If
    If GetIni(strName, "PageBottomMargin") <> "" Then
        m_PrintInfo.PrnPageBottomMargin = GetIni(strName, "PageBottomMargin")
    End If
    If GetIni(strName, "PageLeftMargin") <> "" Then
        m_PrintInfo.PrnPageLeftMargin = GetIni(strName, "PageLeftMargin")
    End If
    If GetIni(strName, "PageRightMargin") <> "" Then
        m_PrintInfo.PrnPageRightMargin = GetIni(strName, "PageRightMargin")
    End If
    If GetIni(strName, "PrnPaperSource") <> "" Then
        m_PrintInfo.PrnPaperSource = GetIni(strName, "PrnPaperSource")
    End If
    
    ReadPrnIni = True

    Exit Function

ReadPrnIni_Err:

    ReadPrnIni = False
    Call MsgBox("設定読み込みエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "ReadPrnIni_Err")

End Function

'目　的　　：INIファイルからのデータ取得
'条　件　　：
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Function GetIni(ByVal vstrKeyword As String, ByVal vstrIniFile As String) As String
    
    Dim strResult   As String * 1024
    Dim intTemp     As Integer
    Dim intLen      As Integer
    Dim strFileName As String
   
    On Error Resume Next
   
    intTemp = GetPrivateProfileString(vstrKeyword, vstrIniFile, "", strResult, Len(strResult), m_strIniFileName)
    GetIni = left$(strResult, intTemp)

End Function

'目　的　　：INIファイルへデータ更新
'条　件　　：内部関数
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Function SetIni(ByVal strSection As String, ByVal vstrKeyword As String, ByVal vstrKeyVal As String) As Integer
    
    Dim intTemp As Integer

    On Error Resume Next

    intTemp = WritePrivateProfileString(strSection, vstrKeyword, vstrKeyVal, m_strIniFileName)
    SetIni = intTemp

End Function

'目　的　　：保存ダイアログを表示してファイル名を取得する
'条　件　　：内部関数
'結　果　　：
'引　数　　：0:PDFファイル 1:Excelファイル
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Private Function Get_FileName(intFlg As Integer) As String
    
    Dim typOpenFileName As OPENFILENAME
    Dim lngHWnd As Long
    Dim strFilter As String
    Dim strDefDir As String
    
    On Error GoTo Get_FileName_Err
    
    If intFlg = 0 Then
        strFilter = "PDFﾌｧｲﾙ(*.PDF)" & vbNullChar & "*.pdf" & vbNullChar & _
                    "全てのﾌｧｲﾙ(*.*)" & vbNullChar & "*.*" & vbNullChar
    ElseIf intFlg = 1 Then
        strFilter = "Excelﾌｧｲﾙ(*.XLS)" & vbNullChar & "*.xls" & vbNullChar & _
                    "全てのﾌｧｲﾙ(*.*)" & vbNullChar & "*.*" & vbNullChar
    End If
    lngHWnd = 0
    strDefDir = ""
    
    With typOpenFileName
        'ｻｲｽﾞの設定
        .lStructSize = Len(typOpenFileName)
        '親Windowの指定
        .hwndOwner = lngHWnd
        'ｱﾌﾟﾘｹｰｼｮﾝのｲﾝｽﾀﾝｽを指定
        .hInstance = App.hInstance
        'ﾌｨﾙﾀの種類を設定
        .lpstrFilter = strFilter
        'ｱｸﾃｨﾌﾞなﾌｨﾙﾀの番号を設定
        .nFilterIndex = 1
        '[ﾌｧｲﾙ名]ﾎﾞｯｸｽの内容を初期化
        .lpstrFile = String(256, Chr(0))
        '最大ﾌｧｲﾙ長の設定
        .nMaxFile = 256
        'ﾌｧｲﾙのﾀｲﾄﾙを受け取るﾎﾟｲﾝﾀ
        .lpstrFileTitle = String(256, Chr(0))
        '最大ﾌｧｲﾙ長の設定
        .nMaxFileTitle = 256
        'ﾃﾞﾌｫﾙﾄﾃﾞｨﾚｸﾄﾘの設定
        .lpstrInitialDir = strDefDir
        'ﾀﾞｲｱﾛｸﾞのﾀｲﾄﾙを設定
        '.lpstrTitle = "ﾌｧｲﾙを開く"
        'ｵﾌﾟｼｮﾝの設定
        .Flags = OFN_EXPLORER Or OFN_PATHMUSTEXIST _
            Or OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
    End With
    
    '[ﾌｧｲﾙを開く]ﾀﾞｲｱﾛｸﾞの表示
    If GetSaveFileName(typOpenFileName) = 0 Then
        'ｷｬﾝｾﾙされた場合
        Get_FileName = ""
    Else
        'OKﾎﾞﾀﾝを押された場合、ﾌｧｲﾙ名の表示
        'Null文字の除去
        If InStr(typOpenFileName.lpstrFile, ".") > 0 Then
            Get_FileName = left(typOpenFileName.lpstrFile, _
                    InStr(typOpenFileName.lpstrFile, vbNullChar) - 1)
        Else
            '拡張子がない場合
            If intFlg = 0 Then
                Get_FileName = left(typOpenFileName.lpstrFile, _
                        InStr(typOpenFileName.lpstrFile, vbNullChar) - 1) & _
                        ".PDF"
            ElseIf intFlg = 1 Then
                Get_FileName = left(typOpenFileName.lpstrFile, _
                        InStr(typOpenFileName.lpstrFile, vbNullChar) - 1) & _
                        ".XLS"
            End If
        End If
    End If

    Exit Function

Get_FileName_Err:

    Get_FileName = ""
    Call MsgBox("ファイル名取得処理エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "Get_FileName_Err")

End Function

Public Property Get Name() As String

    Name = m_PrintInfo.Name
    
End Property

Public Property Let Name(Param As String)

    m_PrintInfo.Name = Param
    '設定の読み込み
    Call ReadPrnIni(m_PrintInfo.Name)
    
End Property

Public Property Get objReport() As Object

    Set objReport = m_PrintInfo.objReport
    
End Property

Public Property Let objReport(Param As Object)

    Set m_PrintInfo.objReport = Param
    
End Property

Public Property Get DatabaseName() As String

    DatabaseName = m_PrintInfo.DatabaseName
    
End Property

Public Property Let DatabaseName(Param As String)

    m_PrintInfo.DatabaseName = Param
    
End Property

Public Property Get Caption() As String

    Caption = m_PrintInfo.Caption
    
End Property

Public Property Let Caption(Param As String)

    m_PrintInfo.Caption = Param
    
End Property

Public Property Get TableForExcel() As String

    TableForExcel = m_PrintInfo.TableForExcel
    
End Property

Public Property Let TableForExcel(Param As String)

    m_PrintInfo.TableForExcel = Param
    
End Property

Public Property Get PrnDefaultPrinter() As Boolean

    PrnDefaultPrinter = m_PrintInfo.PrnDefaultPrinter
    
End Property

Public Property Let PrnDefaultPrinter(Param As Boolean)

    m_PrintInfo.PrnDefaultPrinter = Param
    
End Property

Public Property Get PrnDeviceName() As String

    PrnDeviceName = m_PrintInfo.PrnDeviceName
    
End Property

Public Property Let PrnDeviceName(Param As String)

    m_PrintInfo.PrnDeviceName = Param
    
End Property

Public Property Get PrnPaperSize() As String

    PrnPaperSize = m_PrintInfo.PrnPaperSize
    
End Property

Public Property Let PrnPaperSize(Param As String)

    m_PrintInfo.PrnPaperSize = Param
    
End Property

Public Property Get PrnOrientation() As String

    PrnOrientation = m_PrintInfo.PrnOrientation
    
End Property

Public Property Let PrnOrientation(Param As String)

    m_PrintInfo.PrnOrientation = Param
    
End Property

Public Property Get PrnPageTopMargin() As String

    PrnPageTopMargin = m_PrintInfo.PrnPageTopMargin
    
End Property

Public Property Let PrnPageTopMargin(Param As String)

    m_PrintInfo.PrnPageTopMargin = Param
    
End Property

Public Property Get PrnPageBottomMargin() As String

    PrnPageBottomMargin = m_PrintInfo.PrnPageBottomMargin
    
End Property

Public Property Let PrnPageBottomMargin(Param As String)

    m_PrintInfo.PrnPageBottomMargin = Param
    
End Property

Public Property Get PrnPageLeftMargin() As String

    PrnPageLeftMargin = m_PrintInfo.PrnPageLeftMargin
    
End Property

Public Property Let PrnPageLeftMargin(Param As String)

    m_PrintInfo.PrnPageLeftMargin = Param
    
End Property

Public Property Get PrnPageRightMargin() As String

    PrnPageRightMargin = m_PrintInfo.PrnPageRightMargin
    
End Property

Public Property Let PrnPageRightMargin(Param As String)

    m_PrintInfo.PrnPageRightMargin = Param
    
End Property

Public Property Get PrnPaperSource() As Integer

    PrnPaperSource = m_PrintInfo.PrnPaperSource
    
End Property

Public Property Let PrnPaperSource(Param As Integer)

    m_PrintInfo.PrnPaperSource = Param
    
End Property

Public Property Get Connection() As String

    Connection = m_PrintInfo.Connection
    
End Property

Public Property Let Connection(Param As String)

    m_PrintInfo.Connection = Param
    
End Property

Public Property Get SQL() As String

    SQL = m_PrintInfo.SQL
    
End Property

Public Property Let SQL(Param As String)

    m_PrintInfo.SQL = Param
    
End Property

Public Property Get NoData() As Boolean

    NoData = m_PrintInfo.NoData
    
End Property

Public Property Let NoData(Param As Boolean)

    m_PrintInfo.NoData = Param
    
End Property


