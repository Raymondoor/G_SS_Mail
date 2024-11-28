// こいつら↓変えないと使えないよ。
// ['項目名', その項目名のある行番号]　（安全のために項目名も入れてもらってるよ。面倒だけど許してね）
const confirmSend = ['相手に通知する', 3];
// [質問者の列番号, 回答してほしい人の列番号,　質問内容の列番号]　（例：Aであれば1と入力）
const messageColumns = [3, 4, 5];
// "https://スプシのURL"
const spsurl = ""

// こっから先は勝手にやってくれるから無視でOK
var sheet = SpreadsheetApp.getActiveSheet(); //Open Sheet
var currentCell;

function submit(e) {
    fetchCellCoord(e);
}

function fetchCellCoord(e) {
  // var sheet = SpreadsheetApp.getActiveSheet(); //Open Sheet
  currentCell = e.range; // Use e.range to get the edited cell
  var value = currentCell.getValue(); // Get Cell's state, true or false
  var row = currentCell.getRow(); // Get Cell's Coord in Numbers
  var column = currentCell.getColumn(); // Get Cell's Coord in Numbers
  var name = sheet.getRange(confirmSend[1], column).getValue();
  // If it is a checkbox and has header value of $name, return that coord, otherwise ignore
  if (value === true && name === confirmSend[0]) {
    var coordArray = [value, row, column];
    Logger.log(coordArray);
    Logger.log(currentCell);
    // showWarning('debug', coordArray.join(', '));
    getContent(coordArray);
  } else {
    return null;
  }
}

function getContent (cell) {
  var mailBuffer = [];
  mailBuffer[0] = sheet.getRange(cell[1], messageColumns[0]).getValue(); // 質問者
  mailBuffer[1] = sheet.getRange(cell[1], messageColumns[1]).getValue(); // 回答者
  mailBuffer[2] = sheet.getRange(cell[1], messageColumns[2]).getValue(); // 内容
  Logger.log(mailBuffer.join(", "));
  if (mailBuffer[0] && mailBuffer[2]) {
    if (mailBuffer[1] !== null) {
      if (sendMail(mailBuffer)) {
        makeCheckboxUneditable();
        showWarning('Success!', 'メールを相手に送信しました！回答を待ちましょう！')
      }
    } else {
      showWarning('Warning:', '回答者の指名が無い場合通知を送れません、ごめんね。');
      return null;
    }
  } else {
    showWarning('Warning:', '質問者と質問内容を埋めてね。');
    return null;
  }
}

function sendMail(mailContent) {
  try {
    MailApp.sendEmail({
      name: mailContent[0],
      to: mailContent[1],
      subject: "地球技大会質問箱に「" + mailContent[0] + "」さんから質問が届いています。",
      htmlBody: '<h4>' + mailContent[0] + "さんからの質問：</h4><p>" + mailContent[2] + "</p>リンクはこちらから→\n" + spsurl,
    });
    return true;
  } catch (error) {
    showWarning('Error:', 'プログラムが作動しませんでした。こっち側の問題です、すみません。Error: ' + error.toString());
    Logger.log(error.toString());
    return false;
  }
}

function showWarning(warnTitle, warnMessage) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(warnTitle, warnMessage, ui.ButtonSet.OK);
}

function makeCheckboxUneditable() {
  if (currentCell) {
    // Align text to center
    currentCell.setHorizontalAlignment("center");
    // Remove checkbox (data validation)
    currentCell.clearDataValidations();
    currentCell.setValue("送信済み");

    // Protect the cell to make it fully uneditable
    var protection = currentCell.protect().setDescription("Read Only Cell");

    // Remove all editors (including the current user)
    var editors = protection.getEditors();
    editors.forEach(function (editor) {
      protection.removeEditor(editor);
    });

    // Prevent domain-wide editing if applicable
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }

    // Change the cell appearance to indicate it's locked
    currentCell.setBackground("lightgrey"); // Optional: Use any color to indicate locking
  } else {
    Logger.log("No cell is selected or provided.");
  }
}

