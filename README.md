# gas-gmail-api-create-label

### サービスで Gmail API を追加します

```javascript
function createLabelsForGmailAccount() {
  var sheet = SpreadsheetApp.getActiveSheet(); // 現在アクティブなスプレッドシートのシートを取得
  var labelNames = sheet.getRange("A:A").getValues().flat(); // シートから1列目の値を取得し、配列に変換

  for (var i = 1; i <= labelNames.length; i++) { // ラベル名の配列をループして、Gmailにラベルを作成
    var labelName = labelNames[i-1];
    if (labelName && !labelName.match(/^\s*$/)) { // ラベル名が空でない場合のみ処理を実行
      try {
        Gmail.Users.Labels.create(
            {
              name: labelName,
              labelListVisibility: 'labelShow',
              messageListVisibility: 'show',
            },
            'me'
          );
        }
      catch(e) {
        console.log(e);
      }
    }
  }
}
```
