// slackUtil
const props = PropertiesService.getScriptProperties();
const slackToken = props.getProperty('SLACK_BOT_TOKEN');
const slackNotifyChannel = props.getProperty('SLACK_NOTIFY_CHANNEL');

const slackAPIURL = 'https://slack.com/api/';
type RequestMethods = 'get' | 'delete' | 'patch' | 'post' | 'put';

// slackUser
interface User {
  id: string;
  team_id: string;
  name: string;
  deleted: boolean;
  color: string;
  real_name: string;
  tz: string;
  tz_label: string;
  tz_offset: number;
  profile: UserProfile;
  is_admin: boolean;
  is_owner: boolean;
  is_primary_owner: boolean;
  is_restricted: boolean;
  is_ultra_restricted: boolean;
  is_bot: boolean;
  updated: number;
  is_app_user: boolean;
  has_2fa: boolean;
}

interface UserProfile {
  avatar_hash: string;
  status_text: string;
  status_emoji: string;
  real_name: string;
  display_name: string;
  real_name_normalized: string;
  display_name_normalized: string;
  email: string;
  image_24: string;
  image_32: string;
  image_48: string;
  image_72: string;
  image_192: string;
  image_512: string;
  team: string;
}

interface UserListResponse {
  ok: boolean;
  members: User[];
  cache_ts: number;
  response_metadata: any;
}

function _getUsers(): User[] | null {
  const resourceURL = 'users.list';

  const reqURL = slackAPIURL + resourceURL;
  const method: RequestMethods = 'get';
  const reqParams = {
    method: method,
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      token: slackToken
    }
  };
  const result = UrlFetchApp.fetch(reqURL, reqParams);
  const content = JSON.parse(result.getContentText()) as UserListResponse;
  if (content.ok) {
    return content.members;
  }
  return null;
}
// slackEmoji
type Emoji = { [name: string]: string };
interface EmojiListResponse {
  ok: boolean;
  emoji: Emoji;
}

function _getEmojis(): Emoji | null {
  const resourceURL = 'emoji.list';

  const reqURL = slackAPIURL + resourceURL;
  const method: RequestMethods = 'get';
  const reqParams = {
    method: method,
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      token: slackToken
    }
  };
  const result = UrlFetchApp.fetch(reqURL, reqParams);
  const content = JSON.parse(result.getContentText()) as EmojiListResponse;
  if (content.ok) {
    return content.emoji;
  }
  return null;
}

// slackNotify
function _notify(channelName: string, text: string) {
  const resourceURL = 'chat.postMessage';

  const reqURL = slackAPIURL + resourceURL;
  const method: RequestMethods = 'post';
  const reqParams = {
    method: method,
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      token: slackToken,
      channel: channelName,
      text: text
    }
  };
  const result = UrlFetchApp.fetch(reqURL, reqParams);
  Logger.log(result);
}

// spreadsheet
const _sheetMemo: { [key: string]: any[][] } = {};

function _getSheetData(sheetname: string, reload = false): any[][] {
  if (reload || !_sheetMemo.hasOwnProperty(sheetname)) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const ss = sheet.getSheetByName(sheetname);
    _sheetMemo[sheetname] = ss.getDataRange().getValues();
  }
  return _sheetMemo[sheetname];
}

// spreadsheetUser
const DataHeader = [
  'id',
  'icon',
  'image_512',
  'display_name',
  'name',
  'updated',
  'updateDate',
  'updatedTime',
  'status_emoji',
  'status_emoji_image',
  'status_text'
];

function _userToRowArray(user: User) {
  const date = new Date(user.updated * 1000);
  const genFormatDatetime = (d: Date) => {
    const format = (i: number) => {
      return ('0' + i).slice(-2);
    };
    const year = d.getFullYear();
    const month = d.getMonth() + 1;
    const date = d.getDate();
    const hour = d.getHours();
    const min = d.getMinutes();
    const sec = d.getSeconds();
    return `${year}/${format(month)}/${format(date)}:${format(hour)}:${format(
      min
    )}:${format(sec)}`;
  };

  return [
    user.id,
    '=IMAGE(INDIRECT("RC[1]", false))',
    user.profile.image_512,
    user.profile.display_name,
    user.name,
    user.updated,
    '=LEFT(INDIRECT("RC[1]", false), (SEARCH(":", INDIRECT("RC[1]", false))-1))',
    genFormatDatetime(date),
    user.profile.status_emoji,
    // emoji_text=>emoji_image
    '=IF(ISNA(VLOOKUP(SUBSTITUTE(INDIRECT("RC[-1]", false), ":", ""), emoji!A:C, 3, FALSE)),INDIRECT("RC[-1]", false), IF(ISURL(VLOOKUP(SUBSTITUTE(INDIRECT("RC[-1]", false), ":", ""), emoji!A:C, 3, FALSE)), IMAGE(VLOOKUP(SUBSTITUTE(INDIRECT("RC[-1]", false), ":", ""), emoji!A:C, 3, FALSE)), VLOOKUP(SUBSTITUTE(INDIRECT("RC[-1]", false), ":", ""), emoji!A:C, 3, FALSE)))',
    user.profile.status_text
  ];
}

// spreadsheetEmoji
const EmojiHeader = ['name', 'path', 'url'];

function _emojiToSheetData(emoji: Emoji): any[][] {
  const data = [];
  for (const name of Object.keys(emoji)) {
    data.push([
      name,
      emoji[name],
      // alias=>URL
      '=IF(ISNA( IF(LEFT(INDIRECT("RC[-1]", false), 5) = "alias",VLOOKUP(RIGHT(INDIRECT("RC[-1]", false), (LEN(INDIRECT("RC[-1]", false))-6)), A:B, 2, FALSE) , INDIRECT("RC[-1]", false))), INDIRECT("RC[-2]", false),  IF(LEFT(INDIRECT("RC[-1]", false), 5) = "alias",VLOOKUP(RIGHT(INDIRECT("RC[-1]", false), (LEN(INDIRECT("RC[-1]", false))-6)), A:B, 2, FALSE) , INDIRECT("RC[-1]", false)))'
    ]);
  }
  return data;
}

// main
const logSheetname = 'log';
const emojiSheetname = 'emoji';

function _userFilter(users: User[]): User[] {
  return users
    .filter(u => u.deleted == false)
    .filter(u => u.is_bot == false)
    .filter(u => u.is_restricted == false)
    .filter(u => u.is_ultra_restricted == false);
}

function _initLog(sheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {
  if (!sheet.getSheetByName(logSheetname)) {
    sheet.insertSheet(logSheetname);
  }
  const users = _getUsers();
  if (users) {
    const logSheet = sheet.getSheetByName(logSheetname);
    logSheet.appendRow(DataHeader);
    for (const user of _userFilter(users)) {
      logSheet.appendRow(_userToRowArray(user));
    }
  }
}

function _initEmoji(sheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {
  if (!sheet.getSheetByName(emojiSheetname)) {
    sheet.insertSheet(emojiSheetname);
  }
  const emoji = _getEmojis();
  if (emoji) {
    const emojiSheet = sheet.getSheetByName(emojiSheetname);
    const emojiData = _emojiToSheetData(emoji);
    emojiSheet
      .getRange(1, 1, emojiData.length, emojiData[0].length)
      .setValues(_emojiToSheetData(emoji));
  }
}

function init() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  _initLog(sheet);
  _initEmoji(sheet);
}

function update() {
  const data = _getSheetData('log');
  const users = _getUsers();
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = sheet.getSheetByName(logSheetname);
  if (users) {
    for (const u of _userFilter(users)) {
      Logger.log({ user: u.name, isUpdate: _isUpdate(data, u) });
      if (_isUpdate(data, u)) {
        logSheet.appendRow(_userToRowArray(u));
        _notify(
          slackNotifyChannel,
          `${u.profile.display_name} update status to ${u.profile.status_emoji}`
        );
      }
    }
  }
}

function _isUpdate(data: any[][], user: User): boolean {
  const idCol = _headerNameToNum(data[0], 'id');
  const updatedCol = _headerNameToNum(data[0], 'updated');
  const userRows = data.filter(row => row[idCol] === user.id);
  if (!userRows) {
    return true;
  }
  const latest = userRows[userRows.length - 1];
  return latest && latest[updatedCol] !== user.updated;
}

function updateEmoji() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const emojiSheet = sheet.getSheetByName(emojiSheetname);
  emojiSheet.clear();
  _initEmoji(sheet);
}

function _headerNameToNum(
  headerRow: string[],
  headerName: string
): number | null {
  let result: number | null = null;
  headerRow.forEach((r, i) => {
    if (r === headerName) {
      result = i;
    }
  });
  return result !== null ? result : null;
}

//
function testNotify() {
  _notify('bot-test', 'testMessage');
}

function testGetEmoji() {
  Logger.log(_getEmojis());
}
