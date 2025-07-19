
’***************月分のシートを生成する******************

'休日をカウントしない
Function IsHoliday(checkDate As Date) As Boolean
    Dim rng As Range
    Dim cell As Range
    Set rng = Worksheets("休日シート").Range("A2:A100") ' 祝日が入っている範囲　A列100行まで

    For Each cell In rng
        If cell.Value = checkDate Then
            IsHoliday = True
            Exit Function
        End If
    Next
    IsHoliday = False
End Function

'今月の月～金シート自動生成
Sub CreateWeekdaySheetsForThisMonth()  
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.StatusBar = "処理実行中..."
    
    'コピー元シートの定義*********
    Dim TemplateSheet As Worksheet
    Set TemplateSheet = ThisWorkbook.Sheets("テンプレートシート") ’コピー元となるシート名
    
    '当月の開始日と最終日を求める*********
    Dim startDate As Date, endDate As Date
    startDate = DateSerial(Year(Date), Month(Date), 1)
                 'DateSerial関数で指定した日付を作成。例：（Year（求めたい日））,...
                 '1 今月の「1日」
                 'Date関数　現在（システム）の日付。返り値は日付型の値
    endDate = DateSerial(Year(Date), Month(Date) + 1, 0)
                'Month+1で来月、0 を日付に指定すると、前月の最終日となる。＝当月最終日
               
    
    '自動生成するシート***********
    
    Dim newSheets As Worksheet  '複製するシートたち
    Dim day_cnt As Date         '日付型のカウンタ

    For day_cnt = startDate To endDate  '1日～最終日まで繰り返す
        If Weekday(day_cnt, vbMonday) <= 5 And Not IsHoliday(day_cnt) Then   ' vbMonday : Mon=1 … Fri=5　月～金曜日の場合　かつ祝日を除く
           
             'テンプレートシートをコピーして、ブック内の最後のシートの後ろに新しいシートを追加
            TemplateSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            Set newSheets = ActiveSheet  'アクティブなシート
            newSheets.Name = Format(day_cnt, "yyyy年mm月dd日")              
            
            ' 日付をセルH1に書き込む
            newSheets.Range("H1").Value = day_cnt
            newSheets.Range("H1").NumberFormat = "yyyy年mm月dd日"
            
            End If
        End If
    Next day_cnt
    
    Application.Calculation = xlCalculationManual
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    MsgBox "シート生成完了"
    
End Sub
