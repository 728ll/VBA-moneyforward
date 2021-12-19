Sub chrome()


  Dim Driver As New Selenium.WebDriver

  Driver.AddArgument ("user-data-dir=" & "USER_PATH") 'ユーザ情報を読み込む

  Driver.Start "Chrome" 'Chromeを開く
  Driver.Get ("https://moneyforward.com/")　'指定したアドレスに飛ぶ


  Cells(x, y) = Driver.FindElementByCss("USER_CLASS").Text '指定のセルに書き込む

  Driver.Close '閉じる

End Sub
