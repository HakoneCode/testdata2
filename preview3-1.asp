<%@ LANGUAGE="VBScript" CODEPAGE=65001 %>
<% Session.CodePage=65001 %>
<%
On Error Resume Next
Set objBP = Server.CreateObject("basp21")
%>

<html lang="ja">
<head>
	<meta charset="UTF-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">

    <style>
      
      .table{
        margin: auto;
      }

      th {
        white-space:nowrap;
      }

      .right{
        text-align: right;
      }

    </style>
</head>

<body>
  <p>年賀状送信のために以下の設定をしてください。</p>

    <table class="table" cellspacing="10">

  <tr><th></th><td class="right"><button id="prevButton">&larr;</button>
    <button id="nextButton">&rarr;</button></td></tr>

    <tr><th>送信リスト読込 ：</th>
  <td><input type="file" id="file1" />
<button id="displayButton">表示</button>

    <pre id="pre1">
		<script>
// 現在の行の初期設定。ヘッダー行はスキップするため、1から始める
let currentRow = 1;
// CSVファイルの全行を格納する配列
let rows = [];

// 「表示」ボタンのクリックイベントリスナー
document.getElementById('displayButton').addEventListener('click', function() {
    // ファイル入力要素を取得
    const f = document.getElementById('file1');
    // ファイルが選択されていない場合はアラートを表示
    if (f.files.length == 0) {
        alert('ファイルを選択してください。');
        return;
    }
    
    // 選択された最初のファイルを取得
    const file = f.files[0];
    const reader = new FileReader();
    // ファイルの読み込みが完了したら実行されるイベント
    reader.onload = () => {
        // 読み込んだデータを行ごとに分割
        const data = reader.result;
        rows = data.split('\n');
        // 初期行を表示
        displayRow(currentRow);
    };
    // ファイルをテキストとして読み込む
    reader.readAsText(file);
});

// 「次へ」ボタンのクリックイベントリスナー
document.getElementById('nextButton').addEventListener('click', function() {
    // 現在の行がファイルの総行数より少ない場合、次の行に移動
    if (currentRow < rows.length - 1) {
        currentRow++;
        displayRow(currentRow);
    }
});

// 「前へ」ボタンのクリックイベントリスナー
document.getElementById('prevButton').addEventListener('click', function() {
    // 現在の行が1より大きい場合、前の行に移動
    if (currentRow > 1) {
        currentRow--;
        displayRow(currentRow);
    }
});

// 指定された行のデータをテキストボックスに表示する関数
function displayRow(row) {
    if (rows.length > row) {
        // 行をカンマで分割してデータを取得
        const columns = rows[row].split(',');
        const no = columns[0]; // 番号
		const company = columns[1]; // 会社名
        const name = columns[2]; // 名前
        const email = columns[3]; // メールアドレス
        // それぞれのデータをテキストボックスに表示
        document.getElementById('no').value = `${no}`;
        document.getElementById('to').value = `${email}`;
		document.getElementById('textbody').value = `${company}\n${name}`;
    }
}
</script></td></tr>

	<form action="sendMail3.asp" method="post" name="preview3">

    
  <tr>
    <th>番　　　　　号 ：</th>
    <td><input type="text" id="no" name="number" size="3" maxlength="3"></td>
  </tr><BR>
  <tr>
    <th>送信元アドレス ：</th>
    <td><input type="text" name="from" size="70" maxlength="254"></td>
  </tr><BR>
  <tr>
    <th>送信先アドレス ：</th>
    <td><input type="text" id="to" name="to" size="70" maxlength="254"></td>
  </tr><BR>
  <tr>
    <th>件　　　　　名 ：</th>
    <td><input type="text" name="subject" size="50" maxlength="70"></td>
  </tr><BR>
    <tr>
        <th><label for="textbody">本　　　　　文 ：</label></th>
        <td><textarea  id="textbody" name="textbody" rows="9" cols="70"></textarea></td>
    </tr><BR>
  <tr>
  

    <tr><th>サムネイル画像 ：</th><td><input type="text" name="snail"><% Request.Form("snail") %></td></tr><BR>
    <tr><th>ハイパーリンク ：</th><td><input type="text" name="hplink"><% Request.Form("hplink") %></td></tr><BR>
    

  <tr>
    <th><label for="textbody">署　　　　　名 ：</th>
    <td></label><textarea id="signature" name="signature" rows="15" cols="70"></textarea></td>
</tr><BR>
  <tr><th></th><td class="right"><input type="submit" value="Send"></td></tr>
  

</div>
</table>
</form>
</body>
</html>

