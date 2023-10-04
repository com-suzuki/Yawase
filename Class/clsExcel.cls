VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExcelOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_varTitleName As Variant
Private m_AdoRecordset As Object

'Windows Api
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

'OPENFILENAME構造体
Private Type OPENFILENAME
    lStructSize As Long             '構造体のサイズ
    hwndOwner As Long               'ウインドウのハンドル
    hInstance As Long               'インスタンスハンドル
    lpstrFilter As String           'フィルター
    lpstrCustomFilter As String     'カスタムフィルター
    nMaxCustFilter As Long          'カスタムフィルターのサイズ
    nFilterIndex As Long            'フィルタのインデックス
    lpstrFile As String             'ファイル名のバッファ
    nMaxFile As Long                'ファイル名のバッファのサイズ
    lpstrFileTitle As String        'フルパス用のバッファ
    nMaxFileTitle As Long           'フルパス用のバッファのサイズ
    lpstrInitialDir As String       'ディレクトリを指定
    lpstrTitle As String            'ダイヤログボックスのタイトル
    flags As Long                   '定数（OFN＿××参照）
    nFileOffset As Integer          'フルパスの中のファイル名までのオフセット
    nFileExtension As Integer       '拡張子までのオフセット
    lpstrDefExt As String           'デフォルトの拡張子
    lCustData As Long               'lpfnHookで渡すデータ
    lpfnHook As Long                'フック関数のポインタ
    lpTemplateName As String        'テンプレート名
End Type

Private Const OFN_ALLOWMULTISELECT = &H200       '複数ファイルを選択可能にする
Private Const OFN_CREATEPROMPT = &H2000          '指定のファイル名が存在しない時にメッセージボックスを表示
Private Const OFN_FILEMUSTEXIST = &H1000         '存在しないファイル名は入力不可
Private Const OFN_HIDEREADONLY = &H4             '読み取り専用のチェックボックスを非表示
Private Const OFN_NOCHANGEDIR = &H8              '他のサブディレクトリから選択不可
Private Const OFN_NOREADONLYRETURN = &H8000      '読み込み専用ファイルと書きこみ禁止ディレクトリの選択不可
Private Const OFN_NOVALIDATE = &H100             'ファイル名の有効性をチェックしない
Private Const OFN_OVERWRITEPROMPT = &H2          '既存のファイル名を指定した時にメッセージを出す
Private Const OFN_PATHMUSTEXIST = &H800          '有効なパスだけをうけつける
Private Const OFN_READONLY = &H1                 '読み取り専用のチェックボックスをチェック
Private Const OFN_SHOWHELP = &H10                'ヘルプボタンを表示

'目　的　　：Excel出力処理(Excelのバージョンに依存しない)
'条　件　　：参照設定でMicrosoft Excel8.0 Object Libraryを参照しない
'結　果　　：
'引　数　　：
'戻り値　　：
'作成者　　：株式会社 コム・エンジニアリング　渥美
'作成年月日：２００２／０５／２４
'更新履歴　：
'
Public Function OutPut() As Boolean

    Dim objExcelApp As Object       'Excel Applicationオブジェクト
    Dim objExcelBook As Object      'Excel Book オブジェクト
    Dim objExcelSheets As Object    'Excel Sheets オブジェクト
    Dim objExcelSheet As Object     'Excel Sheet オブジェクト
    Dim strFilter As String
    Dim strSaveFile As String
    Dim lngRow As Long
    Dim lngCol As Long
    
    On Error GoTo OutPut_Err

    OutPut = False

    '出力先指定
    strFilter = "Excelファイル(*.xls)" & Chr(0) & "*.xls" & Chr(0)
    strFilter = strFilter & "すべてのファイル(*.*)" & Chr(0) & "*.*" & Chr(0)
    strSaveFile = OpenSaveDlg(0, strFilter, "", "xls")
    If Trim(strSaveFile) = "" Then
        Exit Function
    End If
    
    'マウスポインタを砂時計に変更
    Screen.MousePointer = vbHourglass

    'Excelオブジェクト取得
    Set objExcelApp = CreateObject("Excel.Application")
    
    'シートの新規作成
    objExcelApp.Workbooks.Add
    Set objExcelBook = objExcelApp.ActiveWorkbook
    Set objExcelSheets = objExcelBook.Worksheets
    Set objExcelSheet = objExcelBook.Sheets(1)
    objExcelApp.Visible = True
    
    'タイトルの出力
    lngRow = 1
    If Not IsNull(m_varTitleName) Then
        For lngCol = 0 To UBound(m_varTitleName)
            objExcelSheet.Cells(lngRow, lngCol + 1).Value = m_varTitleName(lngCol)
        Next lngCol
        lngRow = lngRow + 1
    End If
    
    'データ出力
    If IsObject(m_AdoRecordset) = True Then
        Do While Not m_AdoRecordset.EOF
            For lngCol = 0 To m_AdoRecordset.Fields.Count - 1
                If Not IsNull(m_AdoRecordset.Fields(lngCol).Value) Then
                    objExcelSheet.Cells(lngRow, lngCol + 1).Value = m_AdoRecordset.Fields(lngCol).Value
                End If
            Next lngCol
            m_AdoRecordset.MoveNext
            lngRow = lngRow + 1
        Loop
    End If
    
    'シートを保存
    objExcelSheet.SaveAs strSaveFile
    
    'マウスポインタを元に戻す
    Screen.MousePointer = vbDefault

    OutPut = True

    Exit Function

OutPut_Err:

    'マウスポインタを元に戻す
    Screen.MousePointer = vbDefault
    Set objExcelSheets = Nothing
    OutPut = False
    Call MsgBox("Excel出力処理エラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "OutPut_Err")
    
End Function

Private Function OpenSaveDlg(lngHandle As Long, strFilter As String, strInitialDir As String, strDefExt As String) As String

    Dim OFN As OPENFILENAME                     'OPENFILENAME構造体
    Dim lngRet As Long                          '戻り値

    On Error GoTo OpenSaveDlg_Err

    With OFN                                    '構造体の設定
        .flags = OFN_PATHMUSTEXIST Or _
                 OFN_FILEMUSTEXIST Or _
                 OFN_HIDEREADONLY Or _
                 OFN_OVERWRITEPROMPT
                 'Or OFN_SHOWHELP
        .hInstance = App.hInstance              'インスタンスハンドルを設定
        .hwndOwner = lngHandle                  'ウインドウハンドルを設定
        '.lpstrTitle = "ファイルを保存する"      'コモンダイアログのタイトルを設定
        .lpstrFilter = strFilter                'フィルターを設定
        .lStructSize = Len(OFN)                 '構造体のサイズを設定
        .nMaxFile = 250                         'ファイル名のバッファサイズを設定
        .lpstrFileTitle = String(250, Chr(0))   'フルパス用のバッファを確保
        .nMaxFileTitle = 250                    'フルパス用のバッファサイズを設定
        .lpstrFile = String(250, Chr(0))        'ファイル名のバッファを確保
        .lpstrInitialDir = strInitialDir        'デフォルトのディレクトリを設定
        .lpstrDefExt = strDefExt                'デフォルトの拡張子を設定
    End With
    
    lngRet = GetSaveFileName(OFN)
    
    If lngRet = 0 Then
         OpenSaveDlg = vbNullString
    Else
         OpenSaveDlg = OFN.lpstrFile             'パスを取得
    End If

    Exit Function

OpenSaveDlg_Err:

    OpenSaveDlg = ""
    Call MsgBox("保存ダイアログエラー！！" _
                & vbCrLf & Error$, vbOKOnly + vbCritical, "OpenSaveDlg_Err")

End Function

Public Property Get TitleName() As Variant

    TitleName = m_varTitleName

End Property

Public Property Let TitleName(ByVal Param As Variant)

    m_varTitleName = Param

End Property

Public Property Get RecordSet() As Object

    Set RecordSet = m_AdoRecordset

End Property

Public Property Let RecordSet(ByVal Param As Object)

    Set m_AdoRecordset = Param

End Property

Private Sub Class_Initialize()

    m_varTitleName = Null

End Sub
