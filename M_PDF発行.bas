Attribute VB_Name = "M_PDF発行"
Option Explicit
'==============================================================================
' モジュール名 : M_PDF発行
' 目的         : 「☆-内訳書用」シートを、選んだ対象者ごとに PDF で保存する
'                （印刷実行マクロ btn印刷_Click の PDF 版）
' 呼び出し元   : UserForm_印刷選択
'                  ・UserForm_Initialize … 対象者一覧を取得
'                  ・btnPDF発行_Click    … PDF発行を実行
'------------------------------------------------------------------------------
' ■仕様（仕様を変えたら、ここと「変更履歴」も必ず更新すること）
'  1. ファイル名    : eemmdd_内訳書_コード_フリガナ.pdf
'                     例）081117_内訳書_70107_ウラベコウスケ.pdf
'                     eemmdd … マクロを実行した日の和暦（令和8年11月17日 → 081117）
'                     内訳書 … 書類の名前（定数「書類名」）
'                     コード … 「☆-内訳書用」I2 の値
'                               0・空欄・エラーのときは「未登録」にする
'  2. 保存先        : 「設定」シートの項目「既定の保存先」
'                     空欄・存在しないフォルダ → この PC のデスクトップ（確認画面に表示）
'                     確認画面で［いいえ］→ その回だけ別のフォルダを選べる（設定は変えない）
'  3. 同じ名前      : 確認せずに上書き。PDF を開いていて上書きできない人はスキップして報告
'  4. 見直し        : N5 が「見直し」（またはエラー）の人は出力せず報告
'  5. 後始末        : J2 は実行前の値に戻す（エラーで止まったときも）
'  6. 対象者一覧    : 「掛_月別-☆」の見出し「カナ」の列のうち、
'                     見出し「合計」の列が 0 以外の人（印刷・PDF 共通の一覧）
'  7. 設定シート    : なければ最初の実行時に自動で作る（項目｜値｜説明 の 3 列）
'  8. 完了後        : 完了メッセージで「保存先フォルダを開きますか？」を選ぶ
'------------------------------------------------------------------------------
' ■変更履歴
'  v1.0  2026/11/17  新規作成
'  v1.1  2026/11/17  ファイル名に書類名「内訳書」を追加（日付とコードの間）
'------------------------------------------------------------------------------
' ■よくある変更の直し方
'  ・シート名が変わった        → 下の「シート名」の定数を直す
'  ・J2/I2/N5 の場所が変わった → 下の「セル番地」の定数を直す
'  ・見出しの文字が変わった    → 下の「見出し」の定数を直す
'  ・ファイル名の「内訳書」を変えたい → 下の定数「書類名」を直す
'  ・保存先を変えたい          → コードではなく「設定」シートの値を直す
'==============================================================================


'==============================================================================
' 定数（仕様変更のときは、原則ここだけ直す）
'==============================================================================

'--- シート名 ---
Private Const SHEET_内訳書 As String = "☆-内訳書用"     ' PDF にするシート
Private Const SHEET_一覧 As String = "掛_月別-☆"       ' 対象者一覧の元になるシート
Private Const SHEET_設定 As String = "設定"             ' 設定シート（なければ自動で作る）

'--- 「☆-内訳書用」のセル番地 ---
Private Const CELL_フリガナ As String = "J2"   ' ここに名前を入れると内訳書が切り替わる
Private Const CELL_コード As String = "I2"     ' 顧客コード（XLOOKUP の数式）
Private Const CELL_見直し As String = "N5"     ' 検算結果（合計が合わないと「見直し」）

'--- 「掛_月別-☆」の見出し（列は見出しの文字で探すので、列を足しても動く）---
Private Const 見出し_カナ As String = "カナ"
Private Const 見出し_合計 As String = "合計"
Private Const 見出しを探す最終行 As Long = 20   ' 1 行目からこの行までで見出しを探す

'--- 設定シート ---
Private Const 設定_見出し行 As Long = 1                       ' 設定シートの見出しの行
Private Const 設定見出し_項目 As String = "項目"
Private Const 設定見出し_値 As String = "値"
Private Const 設定見出し_説明 As String = "説明"
Private Const 項目_保存先 As String = "既定の保存先"
Private Const 保存先の初期値 As String = "C:\Users\nakamura\Desktop"
Private Const 保存先の説明 As String = "PDF の保存先フォルダ。空欄または存在しないフォルダのときは、この PC のデスクトップを使います。"

'--- ファイル名 ---
Private Const 日付の書式 As String = "eemmdd"      ' 和暦の年 2 桁＋月 2 桁＋日 2 桁
Private Const 令和の基準年 As Long = 2018         ' 予備計算用（西暦 - 2018 = 令和の年）
Private Const 書類名 As String = "内訳書"        ' ファイル名に入れる書類の名前
Private Const 区切り文字 As String = "_"
Private Const 未登録の表示 As String = "未登録"
Private Const 拡張子 As String = ".pdf"
Private Const 使えない文字 As String = "\/:*?""<>|"   ' ファイル名に使えない文字
Private Const 置き換える文字 As String = "_"

'--- 判定に使う文字 ---
Private Const 見直しの表示 As String = "見直し"

'--- 画面表示 ---
Private Const 画面タイトル As String = "PDF発行"
Private Const 一覧のタイトル As String = "対象者の選択"
Private Const 表示する最大人数 As Long = 20     ' メッセージに名前を並べる上限（超えたら「ほか○名」）

'--- 設定シートの列（列を入れ替えたら、ここの番号を直す）---
Private Enum 設定列
    項目列 = 1
    値列 = 2
    説明列 = 3
End Enum

'--- 設定シートの列幅 ---
Private Const 列幅_項目 As Double = 18
Private Const 列幅_値 As Double = 45
Private Const 列幅_説明 As Double = 80


'==============================================================================
' 公開する処理 (1) 対象者一覧を取得
'   「掛_月別-☆」から、合計が 0 以外の人のフリガナを集めて返す。
'   フォームの UserForm_Initialize から呼ばれる（印刷・PDF 共通の一覧）。
'   うまくいかないときはメッセージを出して Nothing を返す。
'==============================================================================
Public Function 対象者一覧を取得() As Collection
    Dim ws As Worksheet
    Dim 結果 As Collection
    Dim 見出し行 As Long
    Dim カナ列 As Long
    Dim 合計列 As Long
    Dim 最終行 As Long
    Dim r As Long
    Dim カナ As String
    Dim 合計 As Variant
    Dim 処理段階 As String

    On Error GoTo ErrHandler

    '--- シートがあるか確認
    処理段階 = "シート「" & SHEET_一覧 & "」の確認"
    Set ws = シートを取得(SHEET_一覧)
    If ws Is Nothing Then Exit Function     ' メッセージは表示済み

    '--- 見出し「カナ」と「合計」の位置を探す
    処理段階 = "見出しの検索"
    If Not 見出しの位置を探す(ws, 見出し_カナ, 見出し行, カナ列) Then
        MsgBox "シート「" & SHEET_一覧 & "」の 1～" & 見出しを探す最終行 & " 行目に、" & _
               "見出し「" & 見出し_カナ & "」が見つかりません。" & vbCrLf & _
               "見出しの文字が変わっていないか確認してください。", vbCritical, 一覧のタイトル
        Exit Function
    End If

    合計列 = 行の中で見出しを探す(ws, 見出し行, 見出し_合計)
    If 合計列 = 0 Then
        MsgBox "シート「" & SHEET_一覧 & "」の " & 見出し行 & " 行目に、" & _
               "見出し「" & 見出し_合計 & "」が見つかりません。" & vbCrLf & _
               "見出しの文字が変わっていないか確認してください。", vbCritical, 一覧のタイトル
        Exit Function
    End If

    '--- 見出しの下から最終行まで、合計が 0 以外の人を集める
    処理段階 = "対象者の読み込み"
    Set 結果 = New Collection
    最終行 = ws.Cells(ws.Rows.Count, カナ列).End(xlUp).Row

    For r = 見出し行 + 1 To 最終行
        カナ = Trim$(セルの文字(ws.Cells(r, カナ列)))
        合計 = ws.Cells(r, 合計列).Value

        If カナ <> "" Then
            If 金額が0以外(合計) Then 結果.Add カナ
        End If
    Next r

    Set 対象者一覧を取得 = 結果
    Exit Function

ErrHandler:
    MsgBox "対象者一覧を作る途中でエラーが発生しました。" & vbCrLf & vbCrLf & _
           "処理段階：" & 処理段階 & vbCrLf & _
           "エラー番号：" & Err.Number & vbCrLf & _
           "内容：" & Err.Description, vbCritical, 一覧のタイトル
    Set 対象者一覧を取得 = Nothing
End Function


'==============================================================================
' 公開する処理 (2) PDF発行を実行
'   引数 対象者 … フォームで選ばれたフリガナの配列
'   戻り値     … 最後まで処理したら True（フォームを閉じる）
'                 キャンセル・エラーのときは False（フォームは開いたまま）
'==============================================================================
Public Function PDF発行を実行(ByRef 対象者() As String) As Boolean
    Dim ws As Worksheet
    Dim 人数 As Long
    Dim 発行日 As String
    Dim 既定フォルダ As String
    Dim 保存先の注記 As String
    Dim 保存先 As String
    Dim 回答 As VbMsgBoxResult

    Dim 元のJ2 As Variant          ' 実行前の J2（最後に戻す）
    Dim J2を変えた As Boolean
    Dim 元の画面更新 As Boolean
    Dim 画面更新を止めた As Boolean

    Dim i As Long
    Dim 名前 As String
    Dim コード As String
    Dim ファイル名 As String
    Dim フルパス As String
    Dim エラー内容 As String
    Dim 出力数 As Long

    Dim 未登録者 As Collection      ' コード未登録で出力した人
    Dim 見直し者 As Collection      ' 見直しのため出力しなかった人
    Dim 上書き不可者 As Collection  ' PDF を開いていて上書きできなかった人
    Dim エラー者 As Collection      ' そのほかのエラーで出力できなかった人

    Dim 処理段階 As String
    Dim 最後まで完了 As Boolean

    On Error GoTo ErrHandler
    PDF発行を実行 = False

    '--------------------------------------------------------------------------
    ' 0. 事前チェック
    '--------------------------------------------------------------------------
    処理段階 = "事前チェック"
    人数 = 配列の件数(対象者)
    If 人数 = 0 Then
        MsgBox "PDF を発行する対象者が選ばれていません。", vbExclamation, 画面タイトル
        Exit Function
    End If

    Set ws = シートを取得(SHEET_内訳書)
    If ws Is Nothing Then Exit Function     ' メッセージは表示済み

    発行日 = 和暦の日付文字(Date)

    '--------------------------------------------------------------------------
    ' 1. 保存先を決める（設定シート → 確認画面）
    '--------------------------------------------------------------------------
    処理段階 = "保存先の読み込み（設定シート）"
    既定フォルダ = 既定の保存先を取得(保存先の注記)

    処理段階 = "実行の確認"
    回答 = MsgBox(確認メッセージを作る(対象者, 人数, 発行日, 既定フォルダ, 保存先の注記), _
                  vbYesNoCancel + vbQuestion, 画面タイトル & "の確認")

    Select Case 回答
        Case vbYes
            保存先 = 既定フォルダ                 ' いつもの保存先
        Case vbNo
            保存先 = フォルダを選ぶ(既定フォルダ)  ' その回だけ別のフォルダ
            If 保存先 = "" Then Exit Function     ' フォルダ選択をキャンセル → 何もしない
        Case Else
            Exit Function                         ' キャンセル → 何もしない
    End Select

    '--------------------------------------------------------------------------
    ' 2. 準備
    '--------------------------------------------------------------------------
    処理段階 = "準備"
    Set 未登録者 = New Collection
    Set 見直し者 = New Collection
    Set 上書き不可者 = New Collection
    Set エラー者 = New Collection

    元のJ2 = ws.Range(CELL_フリガナ).Formula      ' 値でも数式でもそのまま戻せるよう Formula で保存

    元の画面更新 = Application.ScreenUpdating
    Application.ScreenUpdating = False            ' 画面のちらつきを止めて速くする
    画面更新を止めた = True

    ページ設定をそろえる ws                       ' ページ設定は最初に 1 回だけ（速くするため）

    '--------------------------------------------------------------------------
    ' 3. 1 人ずつ PDF にする
    '--------------------------------------------------------------------------
    For i = LBound(対象者) To UBound(対象者)
        名前 = 対象者(i)
        処理段階 = "PDF 出力（" & 名前 & "）"
        Application.StatusBar = "PDF出力中 " & (i - LBound(対象者) + 1) & "/" & 人数 & "　" & 名前

        '--- 3-1. J2 に名前を入れて、内訳書をその人に切り替える
        ws.Range(CELL_フリガナ).Value = 名前
        J2を変えた = True
        Application.Calculate

        '--- 3-2. 見直しチェック（合わない人は出力しない）
        If 見直しが出ている(ws.Range(CELL_見直し).Value) Then
            見直し者.Add 名前
        Else
            '--- 3-3. ファイル名を作る（コードがなければ「未登録」）
            コード = コードの文字(ws.Range(CELL_コード).Value)

            ファイル名 = 発行日 & 区切り文字 & 書類名 & 区切り文字 & _
                         コード & 区切り文字 & ファイル名に使える形(名前) & 拡張子
            フルパス = パスをつなぐ(保存先, ファイル名)

            '--- 3-4. PDF を書き出す（同じ名前は上書き）
            If ファイルが使用中(フルパス) Then
                上書き不可者.Add 名前
            ElseIf PDFを書き出す(ws, フルパス, エラー内容) Then
                出力数 = 出力数 + 1
                If コード = 未登録の表示 Then 未登録者.Add 名前   ' 出力できた未登録の人だけ報告
            Else
                エラー者.Add 名前 & "（" & エラー内容 & "）"
            End If
        End If
    Next i

    最後まで完了 = True

Cleanup:
    '--------------------------------------------------------------------------
    ' 4. 後始末（エラーで止まったときもここを通る）
    '--------------------------------------------------------------------------
    On Error Resume Next
    If J2を変えた Then
        ws.Range(CELL_フリガナ).Formula = 元のJ2      ' J2 を実行前の値に戻す
        Application.Calculate
    End If
    Application.StatusBar = False
    Application.PrintCommunication = True
    If 画面更新を止めた Then Application.ScreenUpdating = 元の画面更新
    On Error GoTo 0

    '--------------------------------------------------------------------------
    ' 5. 完了メッセージ（最後まで終わったときだけ）
    '--------------------------------------------------------------------------
    If 最後まで完了 Then
        完了メッセージを出す 人数, 出力数, 保存先, 未登録者, 見直し者, 上書き不可者, エラー者
        PDF発行を実行 = True
    End If
    Exit Function

ErrHandler:
    MsgBox "PDF発行の途中でエラーが発生したため、処理を中止しました。" & vbCrLf & vbCrLf & _
           "処理段階：" & 処理段階 & vbCrLf & _
           "エラー番号：" & Err.Number & vbCrLf & _
           "内容：" & Err.Description & vbCrLf & vbCrLf & _
           "※ここまでに保存した PDF はそのまま残っています。" & _
           IIf(J2を変えた, vbCrLf & "※" & CELL_フリガナ & " は実行前の値に戻します。", ""), _
           vbCritical, 画面タイトル & "（エラー）"
    Resume Cleanup
End Function


'==============================================================================
' ここから下は内部で使う部品（Private）
'==============================================================================

'------------------------------------------------------------------------------
' シートを取得：名前のシートを返す。なければメッセージを出して Nothing
'------------------------------------------------------------------------------
Private Function シートを取得(ByVal シート名 As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(シート名)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "シート「" & シート名 & "」が見つかりません。" & vbCrLf & _
               "シート名が変わっていないか確認してください。" & vbCrLf & _
               "（変わった場合は、モジュール M_PDF発行 の冒頭にある" & vbCrLf & _
               "　シート名の定数を直すと動くようになります）", vbCritical, 画面タイトル
    End If

    Set シートを取得 = ws
End Function

'------------------------------------------------------------------------------
' 見出しの位置を探す：1 行目～「見出しを探す最終行」で、見出し名と一致するセルを探す
'   見つかったら 行・列 に位置を入れて True を返す
'------------------------------------------------------------------------------
Private Function 見出しの位置を探す(ByVal ws As Worksheet, ByVal 見出し名 As String, _
                                    ByRef 行 As Long, ByRef 列 As Long) As Boolean
    Dim r As Long
    Dim 探す最終行 As Long

    探す最終行 = 見出しを探す最終行
    If 探す最終行 > 使用範囲の最終行(ws) Then 探す最終行 = 使用範囲の最終行(ws)

    For r = 1 To 探す最終行
        列 = 行の中で見出しを探す(ws, r, 見出し名)
        If 列 > 0 Then
            行 = r
            見出しの位置を探す = True
            Exit Function
        End If
    Next r

    見出しの位置を探す = False
End Function

'------------------------------------------------------------------------------
' 行の中で見出しを探す：指定した行で、見出し名と完全に一致するセルの列番号を返す
'   見つからなければ 0（「合計額」のような似た文字には一致しない）
'------------------------------------------------------------------------------
Private Function 行の中で見出しを探す(ByVal ws As Worksheet, ByVal 行 As Long, _
                                      ByVal 見出し名 As String) As Long
    Dim c As Long
    Dim 最終列 As Long

    最終列 = 使用範囲の最終列(ws)
    For c = 1 To 最終列
        If Trim$(セルの文字(ws.Cells(行, c))) = 見出し名 Then
            行の中で見出しを探す = c
            Exit Function
        End If
    Next c

    行の中で見出しを探す = 0
End Function

'------------------------------------------------------------------------------
' 使用範囲の最終行・最終列：シートで使われている範囲の端を返す
'------------------------------------------------------------------------------
Private Function 使用範囲の最終行(ByVal ws As Worksheet) As Long
    With ws.UsedRange
        使用範囲の最終行 = .Row + .Rows.Count - 1
    End With
End Function

Private Function 使用範囲の最終列(ByVal ws As Worksheet) As Long
    With ws.UsedRange
        使用範囲の最終列 = .Column + .Columns.Count - 1
    End With
End Function

'------------------------------------------------------------------------------
' セルの文字：セルの値を文字で返す（エラー値のときは空文字）
'------------------------------------------------------------------------------
Private Function セルの文字(ByVal セル As Range) As String
    If IsError(セル.Value) Then
        セルの文字 = ""
    Else
        セルの文字 = CStr(セル.Value)
    End If
End Function

'------------------------------------------------------------------------------
' 金額が0以外：数値で 0 以外なら True（空欄・文字・エラーは False）
'------------------------------------------------------------------------------
Private Function 金額が0以外(ByVal 金額 As Variant) As Boolean
    金額が0以外 = False
    If IsError(金額) Then Exit Function
    If IsEmpty(金額) Then Exit Function
    If Not IsNumeric(金額) Then Exit Function
    金額が0以外 = (CDbl(金額) <> 0)
End Function

'------------------------------------------------------------------------------
' 見直しが出ている：N5 が「見直し」またはエラーなら True
'------------------------------------------------------------------------------
Private Function 見直しが出ている(ByVal 値 As Variant) As Boolean
    If IsError(値) Then
        見直しが出ている = True
    Else
        見直しが出ている = (InStr(1, CStr(値), 見直しの表示) > 0)
    End If
End Function

'------------------------------------------------------------------------------
' コードの文字：I2 の値をファイル名用の文字にする
'   0・空欄・エラー → 「未登録」
'------------------------------------------------------------------------------
Private Function コードの文字(ByVal 値 As Variant) As String
    Dim s As String

    If IsError(値) Then
        コードの文字 = 未登録の表示
        Exit Function
    End If

    s = Trim$(CStr(値))
    If s = "" Or s = "0" Then
        コードの文字 = 未登録の表示
    Else
        コードの文字 = ファイル名に使える形(s)
    End If
End Function

'------------------------------------------------------------------------------
' 和暦の日付文字：日付を eemmdd（例 081117）にする
'   日本語以外の地域設定で和暦にならないときは、令和として計算する（予備）
'------------------------------------------------------------------------------
Private Function 和暦の日付文字(ByVal 日付 As Date) As String
    Dim s As String

    s = Format$(日付, 日付の書式)
    If Len(s) = 6 And IsNumeric(s) Then
        和暦の日付文字 = s
    Else
        和暦の日付文字 = Format$(Year(日付) - 令和の基準年, "00") & Format$(日付, "mmdd")
    End If
End Function

'------------------------------------------------------------------------------
' ファイル名に使える形：使えない文字を「_」に置き換え、前後の空白を取る
'------------------------------------------------------------------------------
Private Function ファイル名に使える形(ByVal 文字 As String) As String
    Dim i As Long
    Dim s As String

    s = 文字
    For i = 1 To Len(使えない文字)
        s = Replace(s, Mid$(使えない文字, i, 1), 置き換える文字)
    Next i
    ファイル名に使える形 = Trim$(s)
End Function

'------------------------------------------------------------------------------
' パスをつなぐ：フォルダとファイル名を「\」でつなぐ（「\」が二重にならないように）
'------------------------------------------------------------------------------
Private Function パスをつなぐ(ByVal フォルダ As String, ByVal ファイル名 As String) As String
    If Right$(フォルダ, 1) = "\" Then
        パスをつなぐ = フォルダ & ファイル名
    Else
        パスをつなぐ = フォルダ & "\" & ファイル名
    End If
End Function

'------------------------------------------------------------------------------
' 配列の件数：配列の要素数を返す（空の配列なら 0）
'------------------------------------------------------------------------------
Private Function 配列の件数(ByRef 配列() As String) As Long
    On Error GoTo EmptyArray
    配列の件数 = UBound(配列) - LBound(配列) + 1
    Exit Function
EmptyArray:
    配列の件数 = 0
End Function

'------------------------------------------------------------------------------
' 既定の保存先を取得：設定シートの「既定の保存先」を返す
'   空欄・存在しないフォルダなら、この PC のデスクトップを返し、注記に理由を入れる
'------------------------------------------------------------------------------
Private Function 既定の保存先を取得(ByRef 注記 As String) As String
    Dim 設定値 As String

    設定値 = Trim$(設定値を読む(項目_保存先, 保存先の初期値, 保存先の説明))
    注記 = ""

    If 設定値 <> "" Then
        If フォルダがある(設定値) Then
            既定の保存先を取得 = 設定値
            Exit Function
        End If
        注記 = "（設定シートの保存先「" & 設定値 & "」が見つからないため、" & _
               "この PC のデスクトップを使います）"
    Else
        注記 = "（設定シートの保存先が空欄のため、この PC のデスクトップを使います）"
    End If

    既定の保存先を取得 = デスクトップのパス()
End Function

'------------------------------------------------------------------------------
' 設定値を読む：設定シートで 項目名 の行を探し、値を返す
'   設定シートがなければ作る。項目がなければ初期値で行を足す
'------------------------------------------------------------------------------
Private Function 設定値を読む(ByVal 項目名 As String, ByVal 初期値 As String, _
                              ByVal 説明文 As String) As String
    Dim ws As Worksheet
    Dim 最終行 As Long
    Dim r As Long

    Set ws = 設定シートを用意()

    最終行 = ws.Cells(ws.Rows.Count, 設定列.項目列).End(xlUp).Row
    For r = 設定_見出し行 + 1 To 最終行
        If Trim$(セルの文字(ws.Cells(r, 設定列.項目列))) = 項目名 Then
            設定値を読む = セルの文字(ws.Cells(r, 設定列.値列))
            Exit Function
        End If
    Next r

    ' 項目が見つからない → 末尾に初期値で追加する
    If 最終行 < 設定_見出し行 Then 最終行 = 設定_見出し行
    r = 最終行 + 1
    ws.Cells(r, 設定列.項目列).Value = 項目名
    ws.Cells(r, 設定列.値列).NumberFormat = "@"          ' 文字として保存
    ws.Cells(r, 設定列.値列).Value = 初期値
    ws.Cells(r, 設定列.説明列).Value = 説明文
    設定値を読む = 初期値
End Function

'------------------------------------------------------------------------------
' 設定シートを用意：設定シートを返す。なければ見出しつきで新しく作る
'------------------------------------------------------------------------------
Private Function 設定シートを用意() As Worksheet
    Dim ws As Worksheet
    Dim 元のシート As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_設定)
    On Error GoTo 0

    If ws Is Nothing Then
        Set 元のシート = ThisWorkbook.ActiveSheet

        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SHEET_設定

        ws.Cells(設定_見出し行, 設定列.項目列).Value = 設定見出し_項目
        ws.Cells(設定_見出し行, 設定列.値列).Value = 設定見出し_値
        ws.Cells(設定_見出し行, 設定列.説明列).Value = 設定見出し_説明
        ws.Rows(設定_見出し行).Font.Bold = True
        ws.Columns(設定列.項目列).ColumnWidth = 列幅_項目
        ws.Columns(設定列.値列).ColumnWidth = 列幅_値
        ws.Columns(設定列.説明列).ColumnWidth = 列幅_説明

        ' 元のシートに戻る（戻せなくても処理は続ける）
        On Error Resume Next
        If Not 元のシート Is Nothing Then 元のシート.Activate
        On Error GoTo 0
    End If

    Set 設定シートを用意 = ws
End Function

'------------------------------------------------------------------------------
' フォルダがある：フォルダが存在すれば True
'------------------------------------------------------------------------------
Private Function フォルダがある(ByVal パス As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    フォルダがある = fso.FolderExists(パス)
End Function

'------------------------------------------------------------------------------
' デスクトップのパス：この PC のデスクトップ（OneDrive に移っていても正しく取れる）
'------------------------------------------------------------------------------
Private Function デスクトップのパス() As String
    Dim p As String

    On Error Resume Next
    p = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    On Error GoTo 0

    If p = "" Then p = Environ$("USERPROFILE") & "\Desktop"
    デスクトップのパス = p
End Function

'------------------------------------------------------------------------------
' フォルダを選ぶ：フォルダ選択の画面を出す。キャンセルなら空文字
'------------------------------------------------------------------------------
Private Function フォルダを選ぶ(ByVal 最初のフォルダ As String) As String
    フォルダを選ぶ = ""
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "PDF の保存先フォルダを選んでください"
        .InitialFileName = パスをつなぐ(最初のフォルダ, "")
        .AllowMultiSelect = False
        If .Show = -1 Then フォルダを選ぶ = .SelectedItems(1)
    End With
End Function

'------------------------------------------------------------------------------
' ページ設定をそろえる：印刷実行と同じ体裁（横・1×1 ページ）にする
'   プリンター通信を止めてまとめて設定するので速い
'   プリンターがなく設定できない場合も、シートに保存済みの設定で出力を続ける
'------------------------------------------------------------------------------
Private Sub ページ設定をそろえる(ByVal ws As Worksheet)
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintArea = ""
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' ファイルが使用中：同じ名前の PDF が開かれていて上書きできなければ True
'------------------------------------------------------------------------------
Private Function ファイルが使用中(ByVal パス As String) As Boolean
    Dim fso As Object
    Dim 番号 As Integer

    ファイルが使用中 = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(パス) Then Exit Function     ' まだない → 使用中ではない

    番号 = FreeFile
    On Error Resume Next
    Open パス For Binary Access Read Write Lock Read Write As #番号
    If Err.Number <> 0 Then ファイルが使用中 = True
    Close #番号
    Err.Clear
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' PDFを書き出す：シートを PDF で保存する。失敗したら False とエラー内容を返す
'------------------------------------------------------------------------------
Private Function PDFを書き出す(ByVal ws As Worksheet, ByVal パス As String, _
                               ByRef エラー内容 As String) As Boolean
    On Error GoTo ErrHandler

    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                           Filename:=パス, _
                           Quality:=xlQualityStandard, _
                           IncludeDocProperties:=False, _
                           IgnorePrintAreas:=False, _
                           OpenAfterPublish:=False
    エラー内容 = ""
    PDFを書き出す = True
    Exit Function

ErrHandler:
    エラー内容 = Err.Description
    PDFを書き出す = False
End Function

'------------------------------------------------------------------------------
' 確認メッセージを作る：実行前の確認画面の文章
'------------------------------------------------------------------------------
Private Function 確認メッセージを作る(ByRef 対象者() As String, ByVal 人数 As Long, _
                                      ByVal 発行日 As String, ByVal 保存先 As String, _
                                      ByVal 注記 As String) As String
    Dim m As String
    Dim i As Long
    Dim 表示数 As Long

    m = "以下 " & 人数 & " 名の PDF を発行します。" & vbCrLf & vbCrLf

    For i = LBound(対象者) To UBound(対象者)
        If 表示数 >= 表示する最大人数 Then
            m = m & "　ほか " & (人数 - 表示数) & " 名" & vbCrLf
            Exit For
        End If
        m = m & "　・" & 対象者(i) & vbCrLf
        表示数 = 表示数 + 1
    Next i

    m = m & vbCrLf
    m = m & "ファイル名：" & 発行日 & 区切り文字 & 書類名 & 区切り文字 & _
            "コード" & 区切り文字 & "フリガナ" & 拡張子 & vbCrLf
    m = m & "保存先：" & 保存先 & vbCrLf
    If 注記 <> "" Then m = m & 注記 & vbCrLf
    m = m & vbCrLf
    m = m & "［はい］　　　この保存先で実行" & vbCrLf
    m = m & "［いいえ］　　保存先を選んで実行" & vbCrLf
    m = m & "［キャンセル］中止"

    確認メッセージを作る = m
End Function

'------------------------------------------------------------------------------
' 完了メッセージを出す：結果をまとめて表示し、保存先フォルダを開くか聞く
'------------------------------------------------------------------------------
Private Sub 完了メッセージを出す(ByVal 選択人数 As Long, ByVal 出力数 As Long, _
                                 ByVal 保存先 As String, _
                                 ByVal 未登録者 As Collection, ByVal 見直し者 As Collection, _
                                 ByVal 上書き不可者 As Collection, ByVal エラー者 As Collection)
    Dim m As String
    Dim アイコン As VbMsgBoxStyle

    m = "PDF発行が完了しました。" & vbCrLf & vbCrLf
    m = m & "出力：" & 出力数 & " 名（選択 " & 選択人数 & " 名）" & vbCrLf
    m = m & "保存先：" & 保存先 & vbCrLf

    m = m & 名前の一覧(未登録者, "■コード未登録のため「" & 未登録の表示 & "」で出力")
    m = m & 名前の一覧(見直し者, "■「" & 見直しの表示 & "」が出ているため未出力（内訳を確認してください）")
    m = m & 名前の一覧(上書き不可者, "■同じ名前の PDF が開かれていて上書きできず未出力（閉じてから再実行してください）")
    m = m & 名前の一覧(エラー者, "■エラーのため未出力")

    m = m & vbCrLf & "保存先フォルダを開きますか？"

    ' 出力できなかった人がいれば注意のアイコンにする
    If 見直し者.Count + 上書き不可者.Count + エラー者.Count > 0 Then
        アイコン = vbExclamation
    Else
        アイコン = vbInformation
    End If

    If MsgBox(m, vbYesNo + アイコン + vbDefaultButton2, 画面タイトル & "（完了）") = vbYes Then
        On Error Resume Next                       ' フォルダが開けなくても PDF は保存済み
        Shell "explorer.exe """ & 保存先 & """", vbNormalFocus
        On Error GoTo 0
    End If
End Sub

'------------------------------------------------------------------------------
' 名前の一覧：完了メッセージに載せる名前の一覧（0 名なら何も出さない）
'------------------------------------------------------------------------------
Private Function 名前の一覧(ByVal 名前たち As Collection, ByVal 見出し As String) As String
    Dim m As String
    Dim i As Long

    If 名前たち.Count = 0 Then
        名前の一覧 = ""
        Exit Function
    End If

    m = vbCrLf & 見出し & "：" & 名前たち.Count & " 名" & vbCrLf
    For i = 1 To 名前たち.Count
        If i > 表示する最大人数 Then
            m = m & "　ほか " & (名前たち.Count - 表示する最大人数) & " 名" & vbCrLf
            Exit For
        End If
        m = m & "　・" & 名前たち(i) & vbCrLf
    Next i

    名前の一覧 = m
End Function
