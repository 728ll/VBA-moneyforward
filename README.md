Sub chrome()


  Dim Driver As New Selenium.WebDriver

  Driver.AddArgument ("user-data-dir=" & "USER_PATH") 'ユーザ情報を読み込む

  Driver.Start "Chrome" 'Chromeを開く
  Driver.Get ("https://moneyforward.com/")　'指定したアドレスに飛ぶ


  Cells(1, 1) = Driver.FindElementByCss("#registered-accounts > ul > li:nth-child(11) > ul.amount").Text '指定のセルに書き込む
  Cells(2, 1) = Driver.FindElementByCss("#registered-accounts > ul > li:nth-child(12) > ul.amount").Text
  Cells(3, 1) = Driver.FindElementByCss("#registered-accounts > ul > li:nth-child(13) > ul.amount").Text
  Cells(4, 1) = Driver.FindElementByCss("#registered-accounts > ul > li:nth-child(14) > ul.amount").Text

  Driver.Close '閉じる

End Sub
