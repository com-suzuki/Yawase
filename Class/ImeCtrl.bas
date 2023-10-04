Attribute VB_Name = "basImectl"
'IMEのｺﾝﾃｷｽﾄを取得するためのAPI
Private Declare Function ImmGetContext Lib "imm32.dll" _
        (ByVal hwnd As Long) As Long

'IMEのｺﾝﾃｷｽﾄを開放するためのAPI
Private Declare Function ImmReleaseContext Lib "imm32.dll" _
        (ByVal hwnd As Long, ByVal himc As Long) As Long
        
'IMEのOn/Off状態を設定するAPI
Private Declare Function ImmSetOpenStatus Lib "imm32.dll" _
        (ByVal himc As Long, ByVal b As Long) As Long

'IMEの初期変換方式や入力ﾓｰﾄﾞを設定するAPI
Private Declare Function ImmSetConversionStatus Lib "imm32.dll" _
    (ByVal himc As Long, ByVal dw1 As Long, ByVal dw2 As Long) As Long

'IMEの初期変換方式や入力ﾓｰﾄﾞを取得するAPI
Private Declare Function ImmGetConversionStatus Lib "imm32.dll" _
        (ByVal himc As Long, lpdw As Long, lpdw2 As Long) As Long

'作成したｺﾝﾃｷｽﾄをｳｨﾝﾄﾞｳを関連付けるAPI
Private Declare Function ImmAssociateContext Lib "imm32.dll" _
        (ByVal hwnd As Long, ByVal himc As Long) As Long

'入力ﾓｰﾄﾞの定数
Private Const IME_CMODE_NATIVE = &H1    '直接入力
Private Const IME_CMODE_KATAKANA = &H2  'ｶﾀｶﾅ
Private Const IME_CMODE_LANGUAGE = &H3  '日本語
Private Const IME_CMODE_FULLSHAPE = &H8 '全角
Private Const IME_CMODE_ROMAN = &H10    'ﾛｰﾏ字
Private Const IME_CMODE_CHARCODE = &H20 'ｺｰﾄﾞ入力

'// 変換モード
Public Enum IME_SMODE_ENUM
   IME_SMODE_NONE = &H0                   '// 無変換
   IME_SMODE_PLAURALCLAUSE = &H1          '// 人名／地名
   IME_SMODE_SINGLECONVERT = &H2          '// シングルキャラクタモード(?)
   IME_SMODE_AUTOMATIC = &H4              '// 自動
   IME_SMODE_PHRASEPREDICT = &H8          '// 次のフレーズを予測(?)
   IME_SMODE_CONVERSATION = &H10          '// 話し言葉優先
End Enum

Dim lngTargetWindow As Long

'目  的    ：IMEの入力ﾓｰﾄﾞを設定
'条  件    ：
'結  果    ：
'引  数    ：
'               0:なし
'               1:オン
'               2:オフ
'               3:オフ固定
'               4:全角ひらがな
'               5:全角カタカナ
'               6:半角入力
'               7:全角英数
'               8:半角英数
'戻り値    ：
'作成者    ：株式会社 コム・エンジニアリング 渥美
'作成年月日：１９９８／３／１３
'更新履歴  ：
'
Public Sub SetImeMode(lngNewTarget As Long, lngNewMode As Long)

    Dim lngAPIReVal As Long
    Dim lngInputMode As Long
    Dim lngConvertMode As Long
    Dim lngIMEHandle As Long

    On Error GoTo SetImeMode_Err

    'IMEのｺﾝﾃｷｽﾄを開放する
    lngAPIReVal = ImmReleaseContext(lngTargetWindow, lngIMEHandle)
    'IMEのｺﾝﾃｷｽﾄを取得
    lngIMEHandle = ImmGetContext(lngNewTarget)
    '現在のコントロールのハンドルを格納
    lngTargetWindow = lngNewTarget
    
    'IMEの初期変換方式と入力ﾓｰﾄﾞを取得
    lngAPIReVal = ImmGetConversionStatus(lngIMEHandle, lngInputMode, lngConvertMode)

    '入力ﾓｰﾄﾞの設定
    Select Case lngNewMode
        Case 0
            'なし
        Case 1
            'オン
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 1)
            Exit Sub
        Case 2
            'オフ
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 0)
            Exit Sub
        Case 3
            'オフ固定
            
        Case 4
            '全角モード
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 1)
            '全角ひらがな
            lngInputMode = lngInputMode And Not IME_CMODE_KATAKANA
            lngInputMode = lngInputMode Or IME_CMODE_FULLSHAPE Or IME_CMODE_NATIVE
            'lngConvertMode = IME_SMODE_ENUM.IME_SMODE_PHRASEPREDICT
        Case 5
            '全角モード
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 1)
            '全角カタカナの場合
            lngInputMode = lngInputMode Or IME_CMODE_LANGUAGE Or _
                           IME_CMODE_FULLSHAPE Or IME_CMODE_KATAKANA
        Case 6
            '半角モード
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 0)
            '半角ｶﾀｶﾅの場合
            lngInputMode = lngInputMode And Not IME_CMODE_FULLSHAPE
            lngInputMode = lngInputMode Or IME_CMODE_LANGUAGE Or IME_CMODE_KATAKANA
        Case 7
            '半角モード
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 0)
            '全角英数の場合
            lngInputMode = lngInputMode And Not IME_CMODE_LANGUAGE
            lngInputMode = lngInputMode Or IME_CMODE_FULLSHAPE
        Case 8
            '半角モード
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 0)
            '半角英数の場合
            lngInputMode = lngInputMode And Not IME_CMODE_FULLSHAPE
            lngInputMode = lngInputMode And Not IME_CMODE_LANGUAGE
        Case 9
            '全角モード
            lngAPIReVal = ImmSetOpenStatus(lngIMEHandle, 1)
            '半角ｶﾀｶﾅの場合
            lngInputMode = lngInputMode And Not IME_CMODE_FULLSHAPE
            lngInputMode = lngInputMode Or IME_CMODE_LANGUAGE Or IME_CMODE_KATAKANA
    End Select

    'IMEの初期変換方式と入力ﾓｰﾄﾞを設定
    lngAPIReVal = ImmSetConversionStatus(lngIMEHandle, lngInputMode, lngConvertMode)

    Exit Sub

SetImeMode_Err:

    MsgBox Error$

End Sub
