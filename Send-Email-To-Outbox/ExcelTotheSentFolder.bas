
Sub ScheduleEmailToSendBox()
    Dim OutlookApp As Object
    Dim MailItem As Object
    Dim MailBody As String
    Dim SendTime As Date

    ' Outlookアプリケーションの起動
    Set OutlookApp = CreateObject("Outlook.Application")
    Set MailItem = OutlookApp.CreateItem(0)

    ' メール本文の作成
    MailBody = Range("C1").Value & vbCrLf & Range("D1").Value

    ' 今日の16時を送信時間として設定
    SendTime = Date + TimeValue("10:15:00")

    ' もし現在が16時を過ぎていたら、翌日の16時に設定
    If Now > SendTime Then
        SendTime = SendTime + 1
    End If
    With MailItem
        .To = Range("A1").Value
        .Subject = Range("B1").Value
        .Body = MailBody
        .DeferredDeliveryTime = SendTime ' 予約送信時間を設定
        .Send ' ← 送信トレイに入れて予約送信
    End With

    ' オブジェクトの解放
    Set MailItem = Nothing
    Set OutlookApp = Nothing

    MsgBox "メールを" & Format(SendTime, "yyyy/mm/dd hh:mm") & "に送信予約しました", vbInformation
End Sub
