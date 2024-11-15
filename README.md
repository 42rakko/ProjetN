##Projet N  
ボット
  
python3 でうごく  
  
#実行に必要な作業  
ブラウザで、  
Google Cloud Console にアクセスして  
projectを作成してください  
作成したプロジェクトの「APIとサービス」から  
「認証情報」を選択して、  
「+認証情報を作成」で  
「サービスアカウント」を作成します  
作成したサービスアカウントを開いて、  
「キー」タブを開いて、  
「鍵を追加」します  
「新しい鍵を作成」して  
キーのタイプは「JSON」を選択します  
作成すると、jsonファイルがダウンロードされるので、  
適当な場所に置きます  
これは、.envでGOOGLE_KEY_PATHで設定しています  

spreadsheet_idを設定してください  
これは、ブラウザでspreadsheetを開いたときのURIから  
/d/ と /edit?  
の間にある文字列です  
  
これでdiscordbot.pyが動作するはずです  
動作しない場合は、
discorcbotにspreadsheetへのアクセス権限を与えてください  

