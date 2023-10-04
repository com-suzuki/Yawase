Attribute VB_Name = "basType"
Option Explicit

Public Const SYSTEM_NAME = "植木市場精算システム"
Public Const PROGRAM_VERSION = "1.00"

'フォーカスセット時の色
Public Const FOCUS_STOP_COLOR = &H80FF80
Public Const FOCUS_NO_COLOR = &HFFFFFF

'処理区分ボタンクリック時の色
Public Const BUTTON_ON = &H80C0FF
Public Const BUTTON_OFF = &H8000000F

'レジストリキー
Public Const REG_KEY = "Software\Com\Yawase"

'端数処理
Public Const FRACTION_OMISSION = 1          '切り捨て
Public Const FRACTION_ROUNDOFF = 2          '四捨五入
Public Const FRACTION_ROUNDUP = 3           '切り上げ

'税区分
Public Const TAX_SOTOZEI = 1                '外税
Public Const TAX_UCHIZEI = 2                '内税
Public Const TAX_HIKAZEI = 3                '非課税

'領収書区分
Public Const RECEIPT_OFF = 0                '出力しない
Public Const RECEIPT_ON = 1                 '出力する

'業務区分
Public Const BUSINESS_DIV_EXHIBITION = 1    '出品
Public Const BUSINESS_DIV_BUYER = 2         '買主
Public Const BUSINESS_DIV_ALL = 3           '両方

'確認表出力区分
Public Const CHECK_REPORT_OFF = 0           '未出力
Public Const CHECK_REPORT_ON = 1            '出力済み

'出品伝票区分
Public Const EXHIBITION_REPORT_OFF = 0      '未出力
Public Const EXHIBITION_REPORT_ON = 1       '出力済み

'買主伝票区分
Public Const BUYER_REPORT_OFF = 0           '未出力
Public Const BUYER_REPORT_ON = 1            '出力済み

'入金区分
Public Const PAYMENT_OFF = 0                '未入金
Public Const PAYMENT_ON = 1                 '入金済み

'入金種別
Public Const PAYMENT_DIV_CASH = 1           '現金
Public Const PAYMENT_DIV_CHECK = 2          '小切手
Public Const PAYMENT_DIV_TRANSFER = 3       '銀行振込

'支払区分
Public Const SHIHARAI_OFF = 0               '未払い
Public Const SHIHARAI_ON = 1                '支払済み

'支払種別
Public Const SHIHARAI_DIV_CASH = 1          '現金
Public Const SHIHARAI_DIV_CHECK = 2         '小切手
Public Const SHIHARAI_DIV_TRANSFER = 3      '銀行振込

'入力済み区分
Public Const INPUT_OFF = 0                  '未入力
Public Const INPUT_ON = 1                   '入力済み

'売上形式
Public Const URIAGE_KEISIKI_KYOUBAI = 1     '競売
Public Const URIAGE_KEISIKI_KOMUKAI = 2     '小向

'再発行区分
Public Const REPRINT_OFF = 0                'しない
Public Const REPRINT_ON = 1                 'する

'鉢物区分
Public Const HATIMONO_DIV_OFF = 0           '鉢物でない
Public Const HATIMONO_DIV_ON = 1            '鉢物

'市内市外区分
Public Const TIKU_DIV_OFF = 0               '市外
Public Const TIKU_DIV_ON = 1                '市内

'競売不成立区分
Public Const AUCTION_ON = 0                 '成立
Public Const AUCTION_OFF = 1                '不成立

'開催日
Public Function HOLDING_DATE() As Variant
    '１・８・１５・２３
    HOLDING_DATE = Array("1", "8", "15", "23")
End Function

'月別の開催日
Public Function MONTH_HOLDING_DATE(intMonth As Integer, intKaisu As Integer) As Variant

    Select Case intMonth
        Case 1:
            MONTH_HOLDING_DATE = Array("15", "23")
            intKaisu = 2
        Case 7:
            MONTH_HOLDING_DATE = Array("1", "8", "15")
            intKaisu = 3
        Case 8:
            MONTH_HOLDING_DATE = Null
            intKaisu = 0
        Case Else
            MONTH_HOLDING_DATE = Array("1", "8", "15", "23")
            intKaisu = 4
    End Select
    
End Function
