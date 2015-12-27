/* 定数定義フィールド始まり */

/** 読み込むエクセルファイルの絶対パス */
var EXCEL_FILE_PATH = "D:\\rabel\\labelCreate\\Table.xlsx";
/** 改行コード */
var LINE_CODE = "\r\n";

/* 定数定義フィールド終わり*/

/* 実処理 */
writeFile("Sheet1");

/* 以降は関数定義 */

/**
 * ファイル出力処理
 * @param {String} sheetName シート名
 */
function writeFile(sheetName) {
try {
    var excelApp = WScript.CreateObject("Excel.Application");
    //エクセルブックのパス、及びシート名を指定する。
    var book = excelApp.Workbooks.Open(EXCEL_FILE_PATH);
    var sheet = book.WorkSheets(sheetName);
    
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
        ja_code += '\'' + sheet.Cells(3 + r, 3) + '\' : \'' + sheet.Cells(3 + r, 4) + '\',' + LINE_CODE;
        en_code += '\'' + sheet.Cells(3 + r, 3) + '\' : \'' + sheet.Cells(3 + r, 5) + '\',' + LINE_CODE;

        r++;
    }
} catch (e) {
    //例外発生時はここを通る。
    WScript.Echo(e);
} finally {
    book.Close();
    excelApp.Quit();
    excelApp = null;
}

//最終的に読み込んだデータをファイルに書き出す
saveToFile("output/test.txt", share_code_start + ja_code + share_code_end);
//saveToFile("output/js/en.js", share_code_start + en_code + share_code_end);
//saveToFile("../dest/js/lang/locale/ja.js", share_code_start + ja_code + share_code_end);
//saveToFile("../dest/js/lang/locale/en.js", share_code_start + en_code + share_code_end);
//saveToFile("../../js/lang/locale/ja.js", share_code_start + ja_code + share_code_end);
//saveToFile("../../js/lang/locale/en.js", share_code_start + en_code + share_code_end);
//saveToFile("output/displayLabelList.php", share_code_forDisp + disp_code);

}

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