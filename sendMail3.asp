<%
Set objcdo = Server.CreateObject("cdo.Message")
with objcdo

//外部SMTPを使用する設定
.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp01.intra.oki.co.jp"
.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
.Configuration.Fields.Update

//テキストボックスからの値を取得
    Dim fromAddress, toAddress, subject, textBody, signature, hplink
    fromAddress = Request.Form("from")
    toAddress = Request.Form("to")
    subject = Request.Form("subject")
    hplink = Request.Form("hplink")
    textBody = Request.Form("textbody")
    signature = Request.Form("signature")

    //Const mojilink = "<a href="https://www.google.com">年始動画再生</a>"


//改行コードを認識する形式に置き換え
    textBody = Replace(textBody, vbCrLf, "<br>")
    signature = Replace(signature, vbCrLf, "<br>")

//テキストボックスからの取得した値を挿入
    .From = fromAddress
    .To = toAddress
    .Subject = subject

//HTML形式で本文を設定
    .HTMLBody = Request.Form("textbody") & "<br>" & <a href="https://www.google.com">年始動画再生</a> & "<br>" & Request.Form("signature")

//送信
.send
end with
Set objcdo = Nothing

//送信完了メッセージ
Response.Write("Clear!!")

%>