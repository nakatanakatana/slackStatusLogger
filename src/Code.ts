// slackUtil
const props = PropertiesService.getScriptProperties();
const slackToken = props.getProperty('SLACK_BOT_TOKEN');

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
function _getSheetData(): any[][] {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  return data;
}

// main
const DataHeader = [
  'id',
  'icon',
  'display_name',
  'name',
  'updated',
  'updatedTime',
  'status_emoji',
  'status_text',
  'image_512'
];

function _userToRowArray(user: User) {
  const date = new Date(user.updated * 1000);
  return [
    user.id,
    '=IMAGE(INDIRECT("RC[7]", false))',
    user.profile.display_name,
    user.name,
    user.updated,
    `${date.getFullYear()}/${date.getMonth() +
      1}/${date.getDate()}:${date.getHours()}:${date.getMinutes()}:${date.getSeconds()}`,
    user.profile.status_emoji,
    user.profile.status_text,
    user.profile.image_512
  ];
}

function _userFilter(users: User[]): User[] {
  return users
    .filter(u => u.deleted == false)
    .filter(u => u.is_bot == false)
    .filter(u => u.is_restricted == false)
    .filter(u => u.is_ultra_restricted == false);
}

function init() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.appendRow(DataHeader);
  const users = _getUsers();
  if (users) {
    for (const u of _userFilter(users)) {
      sheet.appendRow(_userToRowArray(u));
    }
  }
}

function update() {
  const data = _getSheetData();
  const users = _getUsers();
  const sheet = SpreadsheetApp.getActiveSheet();
  if (users) {
    for (const u of _userFilter(users)) {
      Logger.log({ user: u.name, isUpdate: _isUpdate(data, u) });
      if (_isUpdate(data, u)) {
        sheet.appendRow(_userToRowArray(u));
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

function testNotify() {
  _notify('bot-test', 'testMessage');
}
