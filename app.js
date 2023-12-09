// LINE Messenging API Token
const CHANNEL_ACCESS_TOKEN = "";

// 1 - 6 班 15 員 
// 7 - 9 班 14 員
const classes = ['one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine'];
const member = [15, 15, 15, 15, 15, 15, 14, 14, 14];
const offset = [0, 15, 30, 45, 60, 75, 90, 104, 118] // list(accumulate(member, initial=0))

// 2 hour 
const cache_time = 2 * 60 * 60 * 1000;

function doPost(e) { 
    // 以 JSON 格式解析 User 端傳來的 e 資料
    var msg = JSON.parse(e.postData.contents);
  
    /* 
    * LINE API JSON 解析資訊
    *
    * replyToken : 一次性回覆 token
    * user_id : 使用者 user id，查詢 username 用
    * userMessage : 使用者訊息，用於判斷是否為預約關鍵字
    * event_type : 訊息事件類型
    */
    const replyToken = msg.events[0].replyToken;
    const user_id = msg.events[0].source.userId;
    const userMessage = msg.events[0].message.text;
    const event_type = msg.events[0].source.type;
  
    // 回傳訊息給line 並傳送給使用者
    function send_to_line(reply_message) {
        var url = 'https://api.line.me/v2/bot/message/reply';
        UrlFetchApp.fetch(url, {
            'headers': {
                'Content-Type': 'application/json; charset=UTF-8',
                'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
            },
            'method': 'post',
            'payload': JSON.stringify({
                'replyToken': replyToken,
                'messages': format_text_message(reply_message),
            }),
        });
    }

    // 將輸入值 word 轉為 LINE 文字訊息格式之 JSON
    function format_text_message(word) {
        let text_json = [{
            "type": "text",
            "text": word
        }]

        return text_json;
    }

    if (typeof replyToken === 'undefined') {
        return;
    };

    const regex1 = /^[1-9]-[0-9][0-9][0-9].*/;
    const regex2 = /^[1-9]班休假回報/;
    if (regex1.test(userMessage)) {
        let team = parseInt(userMessage[0])
        let id = parseInt(userMessage.slice(2,5)) - offset[team - 1];
        let doing = userMessage.slice(9); // 若存成數字會錯(arr[1].trim()) => 需強型別轉換成string
        let thing = [[doing, Date.now()]];
        let data = get_data(classes[team - 1]);
        data.getRange(id, 2, 1, 2).setValues(thing);
        check_all(data, team);
    } else if (regex2.test(userMessage)) {
        let team = parseInt(userMessage[0])
        let data = get_data(classes[team - 1]);
        send_to_line(get_report(data, team));
    }
}

function get_data(sheet_name) {
    /*
    * Google Sheet 資料表資訊設定
    *
    * 將 sheet_url 改成你的 Google sheet 網址
    * 將 sheet_name 改成你的工作表名稱
    */
    const sheet_url = '';
    const SpreadSheet = SpreadsheetApp.openByUrl(sheet_url);
    const data = SpreadSheet.getSheetByName(sheet_name);
    //Logger.log(data.getSheetValues(1,1,member[team - 1],3));
    return data;
}


function get_report(data, team) {
    let report = get_title(team) + '\n';
    let tb = data.getSheetValues(1, 1, member[team - 1], 3);
    let expire = Date.now() - cache_time;
    for (let i = 0; i < member[team - 1]; i++) {
        let arr = tb[i];
        arr[1] = arr[1].trim();
        if (arr[2] < expire) {
            arr[1] = '?';
        }
        report += (arr[0] + ' ' + arr[1] + (i < member[team - 1] - 1 ? '\n':''));
    }
    //Logger.log(report);
    return report;
}

function get_title(team) {
    // 1班 11/11 11:00 休假回報
    // 取得執行時的當下時間
    let time = Utilities.formatDate(new Date(), "Asia/Taipei", "M/d H").split(' '); 
    time[1] = Math.abs(parseInt(time[1]) - 11) < Math.abs(parseInt(time[1]) - 18) ? 11 : 18;
    let title = team + '班 ' + time[0] + ' ' + time[1] + ':00 休假回報';
        
    //Logger.log(title);
    return title;
}

function check_all(data, team) {
    let tb = data.getSheetValues(1, 3, member[team - 1], 1);
    let expire = Date.now() - cache_time;
    for (let i = 0; i < member[team - 1]; i++) {
        data.getRange(i + 1, 4).setValue(tb[i][0] < expire ? '\u274C': '\u2705');
    }
    /*
    if (count == member) {
        send_to_line(classes + " 班 " + member + " 員 已回報完畢");
    }
    */
}
