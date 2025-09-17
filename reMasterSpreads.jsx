//選択したページに選択したマスターページを適合させるスクリプト
// Script to adapt the master page to the selected page

app.doScript(main, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "親ページ適用");

function main() {
    if (!app.documents.length) return;
    var doc = app.activeDocument;

    // ダイアログ作成
    var dlg = new Window("dialog", "親ページ適用");

    // 親ページ選択
    dlg.add("statictext", undefined, "親ページ選択:");
    var masterList = [];
    for (var i = 0; i < doc.masterSpreads.length; i++) {
        masterList.push(doc.masterSpreads[i].name);
    }
    var masterDropdown = dlg.add("dropdownlist", undefined, masterList);
    masterDropdown.selection = 0;

    // 対象ページ選択
    dlg.add("statictext", undefined, "対象ページ:");
    var modeGroup = dlg.add("group");
    modeGroup.orientation = "row";

    var oddBtn = modeGroup.add("radiobutton", undefined, "奇数");
    var evenBtn = modeGroup.add("radiobutton", undefined, "偶数");
    var customBtn = modeGroup.add("radiobutton", undefined, "任意指定");
    oddBtn.value = true; // デフォルトは奇数

    // 入力エリア
    var inputGroup = dlg.add("group");
    inputGroup.add("statictext", undefined, "任意指定:");
    var pageInput = inputGroup.add("edittext", undefined, "");
    pageInput.characters = 20;
    pageInput.enabled = false; // デフォルトは無効

    // ラジオボタン切替イベント
    oddBtn.onClick = function () { pageInput.enabled = false; };
    evenBtn.onClick = function () { pageInput.enabled = false; };
    customBtn.onClick = function () { pageInput.enabled = true; };

    // ボタン
    var btnGroup = dlg.add("group");
    btnGroup.alignment = "right";
    var okBtn = btnGroup.add("button", undefined, "OK");
    var cancelBtn = btnGroup.add("button", undefined, "キャンセル", {name: "cancel"});

    var result = dlg.show();
    if (result != 1) {
        return; // キャンセル時は閉じて終了
    }

    var masterName = masterDropdown.selection.text;
    var masterSpread = doc.masterSpreads.item(masterName);

    // ドキュメントの論理ページ番号の最初と最後を取得
    var firstPageNum = parseInt(doc.pages[0].name, 10);
    var lastPageNum  = parseInt(doc.pages[-1].name, 10);

    // ページ指定解析
    var pages = [];
    if (oddBtn.value) {
        for (var i = firstPageNum; i <= lastPageNum; i++) {
            if (i % 2 !== 0) pages.push(i);
        }
    } else if (evenBtn.value) {
        for (var i = firstPageNum; i <= lastPageNum; i++) {
            if (i % 2 === 0) pages.push(i);
        }
    } else if (customBtn.value) {
        pages = parsePages(pageInput.text, firstPageNum, lastPageNum);
    }

    // 親ページ適用
    for (var i = 0; i < pages.length; i++) {
        var p = doc.pages.itemByName(pages[i].toString());
        if (p.isValid) {
            p.appliedMaster = masterSpread;
        }
    }

    alert("親ページ「" + masterName + "」を " + pages.length + " ページに適用しました。");
}

// ページ指定文字列を解析する関数（論理ページ番号対応）
function parsePages(input, firstPageNum, lastPageNum) {
    var result = [];
    input = input.replace(/\s+/g, "");

    var parts = input.split(",");
    for (var i = 0; i < parts.length; i++) {
        if (parts[i].match(/^\d+$/)) {
            var n = parseInt(parts[i], 10);
            if (n >= firstPageNum && n <= lastPageNum) result.push(n);
        } else if (parts[i].match(/^(\d+)-(\d+)$/)) {
            var m = parts[i].match(/^(\d+)-(\d+)$/);
            var start = parseInt(m[1], 10);
            var end = parseInt(m[2], 10);
            if (start > end) { var tmp = start; start = end; end = tmp; }
            for (var j = start; j <= end; j++) {
                if (j >= firstPageNum && j <= lastPageNum) result.push(j);
            }
        }
    }
    return result;
}
