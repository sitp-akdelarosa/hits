<%
'========================================================================
' ExcelCreator Ver3.0 定数一覧
'========================================================================

' Excelバージョン
Public Const ver95 = 95                        ' Excel95
Public Const ver97 = 97                        ' Excel97
Public Const ver2000 = 2000                    ' Excel2000
Public Const ver2002 = 2002                    ' Excel2002

' カラーパレット
Public Const xcDefault = 0                     ' カラーパレット 自動
Public Const xcBlack = 8                       '                黒
Public Const xcWhite = 9                       '                白
Public Const xcRed = 10                        '                赤
Public Const xcGreen = 11                      '                緑
Public Const xcBlue = 12                       '                青
Public Const xcYellow = 13                     '                黄
Public Const xcPink = 14                       '                ピンク
Public Const xcCyan = 15                       '                水
' フォントスタイル
Public Const fsDefault = &H0                   ' フォントスタイル ノーマル
Public Const fsNormal = &H0                    '                  ノーマル
Public Const fsBold = &H1                      '                  ボールド
Public Const fsItalic = &H2                    '                  イタリック
Public Const fsStrike = &H4                    '                  取消し線
Public Const fsUp = &H8                        '                  上付き
Public Const fsDown = &H10                     '                  下付き
' フォント下線
Public Const fuDefault = 0                     ' フォント下線 なし
Public Const fuNone = 0                        '              なし
Public Const fuNormal = 1                      '              下線
Public Const fuDouble = 2                      '              二重線
' 罫線の線種
Public Const lsNone = &H0                      ' 罫線 なし
Public Const lsDefault = &H0                   '      なし
Public Const lsNormal = &H1000000              '      実線
Public Const lsThick = &H2000000               '      太線
Public Const lsBroken = &H3000000              '      破線
Public Const lsDot = &H4000000                 '      点線
Public Const lsThick2 = &H5000000              '      極太線
Public Const lsDouble = &H6000000              '      二重線
Public Const lsSlender = &H7000000             '      細実線
Public Const lsMidBroken = &H8000000           '      中破線
Public Const lsSlnChain1 = &H9000000           '      細一点鎖線
Public Const lsMidChain1 = &HA000000           '      中一点鎖線
Public Const lsSlnChain2 = &HB000000           '      細二点鎖線
Public Const lsMidChain2 = &HC000000           '      中二点鎖線
Public Const lsMidChains = &HD000000           '      中一点斜鎖線
' パターン番号
Public Const pn01 = &H1000000                  ' パターン番号 １
Public Const pn02 = &H2000000                  ' パターン番号 ２
Public Const pn03 = &H3000000                  ' パターン番号 ３
Public Const pn04 = &H4000000                  ' パターン番号 ４
Public Const pn05 = &H5000000                  ' パターン番号 ５
Public Const pn06 = &H6000000                  ' パターン番号 ６
Public Const pn07 = &H7000000                  ' パターン番号 ７
Public Const pn08 = &H8000000                  ' パターン番号 ８
Public Const pn09 = &H9000000                  ' パターン番号 ９
Public Const pn10 = &HA000000                  ' パターン番号 10
Public Const pn11 = &HB000000                  ' パターン番号 11
Public Const pn12 = &HC000000                  ' パターン番号 12
Public Const pn13 = &HD000000                  ' パターン番号 13
Public Const pn14 = &HE000000                  ' パターン番号 14
Public Const pn15 = &HF000000                  ' パターン番号 15
Public Const pn16 = &H10000000                 ' パターン番号 16
Public Const pn17 = &H11000000                 ' パターン番号 17
Public Const pn18 = &H12000000                 ' パターン番号 18
' ボックスの形状
Public Const btBox = 0                         ' ボックスの形状 箱線
Public Const btLtc = 1                         '                格子線
Public Const btOver = 2                        '                上横線
Public Const btUnder = 3                       '                下横線
Public Const btLeft = 4                        '                左縦線
Public Const btRight = 5                       '                右縦線
' 水平位置
Public Const phNormal = 20                     ' 水平位置 標準
Public Const phLeft = 21                       '          左詰
Public Const phCenter = 22                     '          中央揃え
Public Const phRight = 23                      '          右詰
Public Const phLoop = 24                       '          繰り返し
Public Const phBothEven = 25                   '          両端揃え
Public Const phSelCenter = 26                  '          選択範囲内で中央
Public Const phEven = 27                       '          均等
' 垂直位置
Public Const pvUp = 0                          ' 垂直位置 上詰め
Public Const pvCenter = 1                      '          中央揃え
Public Const pvDown = 2                        '          下詰め
Public Const pvBothEven = 3                    '          両端揃え
Public Const pvEven = 4                        '          均等
' 文字方向
Public Const ptHorz = 100                      ' 文字方向 横
Public Const ptVert = 101                      '          縦
Public Const ptRvTurn90 = 102                  '          半時計回りに９０°
Public Const ptTurn90 = 103                    '          時計回りに９０°
' 用紙サイズ
Public Const psLetter = 1                      ' 用紙サイズ レター
Public Const psRejoer = 3                      '            レジャー
Public Const psRegal = 5                       '            リーガル
Public Const psExecutive = 7                   '            エグゼクティブ
Public Const psA3 = 8                          '            Ａ３
Public Const psA4 = 9                          '            Ａ４
Public Const psA5 = 11                         '            Ａ５
Public Const psB4 = 12                         '            Ｂ４
Public Const psB5 = 13                         '            Ｂ５
' 用紙方向
Public Const orLandscape = 0                   ' 用紙方向 横
Public Const orPortrait = 2                    '          縦
' オープンモード
Public Const opOpen = 0                        ' オープンモード オープン
Public Const opCreate = 1                      '                クリエイト
' ページ中央
Public Const pcNone = 0                        ' 指定なし
Public Const pcHorz = 1                        ' 水平方向
Public Const pcVert = 2                        ' 垂直方向
' ページ方向
Public Const poRightFromLeft = 0               ' 左から右へ
Public Const poBottomFromTop = 1               ' 上から下へ
' 変数名の差し込みモード
Public Const vmValidAll = 0                    ' 全ての変数名を対象
Public Const vmInvalidAnswer = 1               ' 関数結果の変数名は対象外
Public Const vmInvalidTextBox = 2              ' テキストボックスの変数名は対象外
Public Const vmInvalidCellText = 4             ' セルに直接設定した変数名は対象外

Public Const vtStr = 0                         ' 文字列
Public Const vtDouble = 1                      ' 実数値
Public Const vtInt = 2                         ' 整数
Public Const vtFunc = 3                        ' 計算式
Public Const vtNull = 4                        ' データなし

Public Const errDontOpen = 1                   ' ファイルがオープンできない
Public Const errFileHandle = 2                 ' 無効なファイルハンドル
Public Const errNotOpen = 3                    ' 未オープンエラー
Public Const errSheetNo = 4                    ' シート番号エラー
Public Const errWrite = 5                      ' 書き込み時のエラー
Public Const errAction = 6                     ' 動作モードエラー
Public Const errDataCat = 7                    ' データ種別エラー
Public Const errClear = 20                     ' クリアエラー
Public Const errCopy = 21                      ' コピーエラー
Public Const errPaste = 22                     ' 貼り付けエラー
Public Const errInsert = 23                    ' 挿入エラー
Public Const errDelete = 24                    ' 削除エラー
Public Const errLength = 31                    ' 長さエラー
Public Const errLocate = 32                    ' 座標エラー
Public Const errAttr = 33                      ' 属性番号エラー
Public Const errParame = 34                    ' パラメータエラー
Public Const errNoData = 35                    ' データが無い
Public Const errEndOfData = 36                 ' データの終わり
Public Const errVerPoint = 37                  ' 指定した変数が存在しない
Public Const errBreakCount = 38                ' 改ページ数が制限を超えたた
Public Const errMemory1 = 40                   ' メモリー不足エラー１
Public Const errMemory2 = 41                   ' メモリー不足エラー２
Public Const errOther = 50                     ' 内部エラー
Public Const errFunction = 100                 ' 計算式形式エラー
%>