try {
    var excelApp = WScript.CreateObject("Excel.Application");
    //エクセルブックのパス、及びシート名を指定する。
    var book = excelApp.Workbooks.Open("C:\\xampp\\htdocs\\damaz_hp\\damaz-hp\\developTools\\labelResourceGen\\Table.xlsx");
    var sheet = book.WorkSheets("Sheet1");
    
    var share_code_start = '';  //共通コード(開始)
    var share_code_end = '';  //共通コード(終了)
    var share_code_forDisp = '';  //共通コード(表示ソース用)
    var ja_code = '';      //日本語用コード
    var en_code = '';      //英語用コード
    var disp_code = '';    //表示用のコード
    
    //ラベル生成ソース共通の処理を書き込む1(始まり）
    share_code_start = '__Localizer.strings = {';
    //ラベル生成ソース共通の処理を書き込む3(終わり)
    share_code_end = '}\;';
    
    //ラベル表示(呼び出し)ソース共通の処理を書き込む1(php宣言)
    share_code_forDisp += '<?php\n\n';
    //ラベル表示(呼び出し)ソース共通の処理を書き込む2(言語選択) ※言語を増やす場合はここを編集する。
    share_code_forDisp += '$lang = setResourceLang(filter_input(INPUT_SERVER, \'HTTP_ACCEPT_LANGUAGE\'));\n\n'
            + 'if($lang == \'ja\'){\n    require_once(ROOT_PATH.\'php/ja_label_server.php\');\n'
            + '}else if($lang == \'en\'){\n    require_once(ROOT_PATH.\'php/en_label_server.php\');\n'
            + '}else{\n    require_once(ROOT_PATH.\'php/en_label_server.php\');\n}\n'
            + '$l1 = new Localizer_server($lang);\n\n';
    
    //繰返し用の変数
    var r = 0;

    //エクセルの内容を読み込む
    while (1) {
        //エラーが起こった個所を見つけるためにidが出るようにしている
        WScript.Echo(sheet.Cells(3 + r, 2).value);

        if (sheet.Cells(3 + r, 2).value === 'END')
            break;
        if (sheet.Cells(3 + r, 2).value.substr(0, 1) === '#') {
            //コメントアウトの回避策
            r++;
            continue;
        }
        //var test = "$script_i18n[\'" + sheet.Cells(3 + r, 3) + "\'][\'ja\'] = \'" + sheet.Cells(3 + r, 4) + "\'";
        ja_code += '\'' + sheet.Cells(3 + r, 3) + '\' : \'' + sheet.Cells(3 + r, 4) + '\',\n';
        en_code += '\'' + sheet.Cells(3 + r, 3) + '\' : \'' + sheet.Cells(3 + r, 5) + '\',\n';
       //disp_code += '$' + sheet.Cells(3 + r, 3) + ' = $l1->g(\'' + sheet.Cells(3 + r, 3) + '\');\n';

        //ja_Str += test + ";\n";
        r++;
    }
} catch (e) {
    //例外発生時はここを通る。
    WScript.Echo('error');
} finally {
    book.Close();
    excelApp.Quit();
    excelApp = null;
}

//最終的に読み込んだデータをファイルに書き出す
saveToFile("output/js/ja.js", share_code_start + ja_code + share_code_end);
saveToFile("output/js/en.js", share_code_start + en_code + share_code_end);
saveToFile("../dest/js/lang/locale/ja.js", share_code_start + ja_code + share_code_end);
saveToFile("../dest/js/lang/locale/en.js", share_code_start + en_code + share_code_end);
saveToFile("../../js/lang/locale/ja.js", share_code_start + ja_code + share_code_end);
saveToFile("../../js/lang/locale/en.js", share_code_start + en_code + share_code_end);
//saveToFile("output/displayLabelList.php", share_code_forDisp + disp_code);

/**
 * UTF-8 BOMなしでファイルを作成する
 * @param {String} fname 出力先ファイル名
 * @param {String} text  出力文字列
 */
function saveToFile(fname, text) {
    // ADODB.Streamのモード
    var adTypeBinary = 1;
    var adTypeText = 2;
    // ADODB.Streamを作成
    var pre = new ActiveXObject("ADODB.Stream");
    // 最初はテキストモードでUTF-8で書き込む
    pre.Type = adTypeText;
    pre.Charset = 'UTF-8';
    pre.Open();
    pre.WriteText(text);
    // バイナリモードにするためにPositionを一度0に戻す
    // Readするためにはバイナリタイプでないといけない
    pre.Position = 0;
    pre.Type = adTypeBinary;
    // Positionを3にしてから読み込むことで最初の3バイトをスキップする
    // つまりBOMをスキップする
    pre.Position = 3;
    var bin = pre.Read();
    pre.Close();

    // 読み込んだバイナリデータをバイナリデータとしてファイルに出力する
    var stm = new ActiveXObject("ADODB.Stream");
    stm.Type = adTypeBinary;
    stm.Open();
    stm.Write(bin);
    stm.SaveToFile(fname, 2); // force overwrite
    stm.Close();
}