function doPost(e) {
  var verificationToken = e.parameter.token;
  var from_name = e.parameter.user_name;
  if (verificationToken != '***Token***') { 
    throw new Error('Invalid token');
  }

  var errflg = false;
  var command = e.parameter.text;
  if (command.indexOf('：') == -1) {
    var text = '【名前：メッセージ】の形式で記入してください！　※「：」が「:」になっていないか確認してください！'
    errflg = true;
  }  
  //「：」の前と後をわける
  var result = command.split( '：' );
  var name = result[0];
  var contents = result[1];
  //改行をなくす
  contents = contents.replace( /\n/g , "" ) ;
  var len = contents.length;
  if(len >= 81)
  {
    var text = 'すみません！メッセージは80文字以内で書いてください！'
    errflg = true;
  }
  if(errflg == true)
  {
    var response = { text: text};
    return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
  }
  
  var date = new Date();
  var year = date.getFullYear();
  var month = date.getMonth() + 1;
  if(month < 10)
  {
    month = '0' + month
  }
  var sheet_name = year + '年' + month　+ '月'
  
  //スプレッドシート
  var spreadsheet = SpreadsheetApp.openById('***spreadsheet***');
  var sheet = spreadsheet.getSheetByName(sheet_name);
  
  //サンクスカードの現在の枚数取得
  var countSheet = sheet.getRange("B3").getValues();
  var count = Number(countSheet) + 1;
  var count_place = 0;
  //6枚以下の場合
  if(count < 6)
  {
    var upper_left_num = 6;
    var bottom_right_num = 13;
  }
  else
  {
    var upper_left_num = 15;
    var bottom_right_num = 22;
    //6枚以上の場合は、5で割る。余りの数によりカードの位置を取得
    count_place = count % 5;
  }
  if(count == 1 || count_place == 1)
  {
    var upper_left = "B";
    var bottom_right = "E";
    var bottom_place = "D";
  }
  else if(count == 2 || count_place == 2)
  {
    var upper_left = "G";
    var bottom_right = "J";
    var bottom_place = "I";
  }
  else if(count == 3 || count_place == 3)
  {
    var upper_left = "L";
    var bottom_right = "O";
    var bottom_place = "N";
  }
  else if(count == 4 || count_place == 4)
  {
    var upper_left = "Q";
    var bottom_right = "T";
    var bottom_place = "S";
  }
  else
  {
    var upper_left = "V";
    var bottom_right = "Y";
    var bottom_place = "X";
  }
  
  var flg = 0;
  while(flg == 0) {
    //スプレッドシートに追加するサンクスカードの位置を取得
    var upper_left_place = upper_left + upper_left_num;
    var bottom_right_place = bottom_right + bottom_right_num;
    var get_place = upper_left_place + ":" + bottom_right_place
    
    //計算結果で取得したカード左上の位置に値が無ければカード投稿。値があれば位置を計算し直す。
    var get_place_value = sheet.getRange(upper_left_place).getValues();
    if(get_place_value == '' || get_place_value == ' ' ||get_place_value == null || count < 6)
    {
      flg = 1;
    }
    else
    {
      upper_left_num = upper_left_num + 9;
      bottom_right_num = bottom_right_num + 9;
    }
  }
  var place = get_place;
  
  //カードを投稿する箇所の背景色を変更
  sheet.getRange(place).setBackground('#ffefd5');
  //呼び捨てにならないようにTOの名前に「さん」をつける
  if ( name.indexOf('さん') != -1 || name.indexOf('くん') != -1 || name.indexOf('君') != -1 || name.indexOf('ちゃん') != -1 || name.indexOf('様') != -1 || name.indexOf('殿') != -1) 
  {
    var to_name_cell = sheet.getRange(upper_left_place).setValue(name + ' へ');
  }
  else
  {
    var to_name_cell = sheet.getRange(upper_left_place).setValue(name + ' さんへ');
  }
  
  //名前の文字の大きさ変更
  to_name_cell.setFontSize('14')
  
  //1行20文字までになるように20行で行を分ける
  var val_start = 0
  var contents_val = contents.substr( val_start, 20 );
  upper_left_num = upper_left_num + 2
  upper_left_place = upper_left + upper_left_num;
  while(val_start <= 60){
    sheet.getRange(upper_left_place).setValue(contents_val);
    val_start += 20;
    contents_val = contents.substr( val_start, 20 );
    upper_left_num = upper_left_num + 1
    upper_left_place = upper_left + upper_left_num;
  }
  
  //FROMの値を入力。文字の大きさ変更
  var bottom_place_cell = bottom_place + bottom_right_num;
  var from_name_cell = sheet.getRange(bottom_place_cell).setValue(from_name);
  from_name_cell.setFontSize('14')
  
  //サンクスカードに枠をつける。
  var rng = sheet.getRange(place);
  rng.setBorder(true, true, true, true, false, false,'#ffdab9',SpreadsheetApp.BorderStyle.SOLID);
  
  //ここまで成功した場合は現在サンクスカード枚数に+1する
  var text = '正常に投稿されました！'
  sheet.getRange("B3").setValue(count);

  var response = { text: text};
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}

