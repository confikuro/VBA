@if(0)==(0) echo off
cscript.exe //nologo //E:JScript "%~f0" %*
goto :EOF
@end

//ショートカット用作成用クラス
var ShortcutCreater = function() {
    this.wshObj = openWsh();
    //JavaScriptでバックスラッシュを表現するためには、"\"を付け（2つにして）エスケープする必要あり（ショートカット名だけ）
    this.shortcutfile = 'C:\\Users\\Confi\\Documents\\cre\\  test.lnk';
    this.link = 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe';
    this.icon = 'C:\Users\Confi\Documents\obake.ico';
    this.create = createShortcut;
    this.cleanup = closeWsh;
    this.toString = createrToString;
}

// メイン処理
function main() {
    var shortcut = null;

    try {
        // ショートカット作成
        shortcut = new ShortcutCreater();
        shortcut.create();

        // 作成したショートカットをコンソールに出力
        Console.println("ショートカットを作成しました");
        Console.println(shortcut);

    } catch (e) {

        // エラー要因をコンソールに出力
        Console.println("[エラー]: " + e.description);

        // 異常終了でコマンドを返す
        Console.back(e.number);

    } finally {

        // WSHオブジェクト片付け
        if (shortcut !== null)
            shortcut.cleanup();
    }

    // 正常終了でコマンドを返す
    Console.back(0);
}

// Windows Script Host実行用クラス
var Console = ((function() {
    var constructor = function() {}
    constructor.println = echoConsole;
    constructor.back = exitScript;
    return constructor;
})())

// 関数一覧

function createShortcut() {
    var lnkFile = this.wshObj.CreateShortcut(this.shortcutfile);
    lnkFile.TargetPath = this.link;
    lnkFile.IconLocation = this.icon;    
    lnkFile.Save();
}

function createrToString() {
    return "file=\"" + this.shortcutfile + "\", linkTo=\"" + this.link + "\"";
}

function openWsh() {
    return WScript.CreateObject("WScript.Shell");
}

function closeWsh() {
    this.wshObj = null;
}

function echoConsole(msg) {
    WScript.echo(msg);
}

function exitScript(errNum) {
    WScript.Quit(errNum);
}

// メイン処理呼び出し
main();