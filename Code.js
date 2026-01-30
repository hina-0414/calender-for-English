const CALENDAR_ID = "1f92d6ad5d14992609bd132d3fb14337e04b9d566ce8cc31c7d2c8d84e85f1d2@group.calendar.google.com";
const SS = SpreadsheetApp.getActiveSpreadsheet();
const MEMBER_SHEET = SS.getSheetByName('member_db');
const RESERVE_SHEET = SS.getSheetByName('reservations');

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('英語科資料室予約システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ログイン処理（学生の予約が授業に上書きされていないかチェックする機能を追加）
function login(id, pw, role) {
  const data = MEMBER_SHEET.getDataRange().getValues();
  let user = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id) && String(data[i][1]) === String(pw) && data[i][2] == role) {
      user = { success: true, name: data[i][3], role: data[i][2] };
      break;
    }
  }

  if (user && role === "学生") {
    user.conflictMessage = checkAndCleanupConflicts(id);
  }

  return user || { success: false };
}

// 学生の予約が授業と重複していないかチェックし、重複していれば削除してメッセージを返す
function checkAndCleanupConflicts(userId) {
  const data = RESERVE_SHEET.getDataRange().getValues();
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  let conflicts = [];
  
  // シートを逆順にスキャン（削除を考慮）
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === String(userId) && data[i][6] === "通常") {
      const dateStr = Utilities.formatDate(new Date(data[i][2]), "JST", "yyyy-MM-dd");
      const startTime = formatToHHmm(data[i][3]);
      const startDateTime = new Date(dateStr + "T" + startTime + ":00");
      const endDateTime = new Date(dateStr + "T" + formatToHHmm(data[i][4]) + ":00");

      // その時間に「授業」のイベントがあるか確認
      const events = calendar.getEvents(startDateTime, endDateTime);
      const hasClass = events.some(e => e.getTitle().includes("[授業]"));

      if (hasClass) {
        conflicts.push(`${dateStr} ${startTime}の予約`);
        RESERVE_SHEET.deleteRow(i + 1);
      }
    }
  }

  if (conflicts.length > 0) {
    return "【重要】教員が後日授業を登録したため、以下の予約は授業が優先され、取り消されました。日時を変更してください：\n" + conflicts.join("\n");
  }
  return null;
}

// 現在資料室が使用中かチェック
function isRoomBusy() {
  try {
    const now = new Date();
    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) return false;
    return calendar.getEventsForDay(now).some(e => e.getStartTime() <= now && e.getEndTime() >= now);
  } catch (e) {
    return false;
  }
}

// 共通の時間フォーマット関数
function formatToHHmm(val) {
  if (val instanceof Date) return Utilities.formatDate(val, "JST", "HH:mm");
  const s = String(val);
  if (s.includes(':')) {
    const parts = s.split(':');
    return parts[0].padStart(2, '0') + ':' + parts[1].substring(0, 2).padStart(2, '0');
  }
  return s;
}

// 自分の予約一覧を取得
function getMyReservations(userId) {
  try {
    const data = RESERVE_SHEET.getDataRange().getValues();
    const myRes = [];
    const now = new Date();
    now.setHours(0,0,0,0);

    let userRole = "学生";
    const memberData = MEMBER_SHEET.getDataRange().getValues();
    for(let i=1; i<memberData.length; i++){
      if(String(memberData[i][0]) === String(userId)) {
        userRole = memberData[i][2];
        break;
      }
    }

    for (let i = 1; i < data.length; i++) {
      if (!data[i][2]) continue;
      let resDate = new Date(data[i][2]);
      let checkDate = new Date(resDate.getTime());
      checkDate.setHours(0,0,0,0);

      if (String(data[i][0]) === String(userId) && checkDate >= now) {
        myRes.push({
          date: Utilities.formatDate(resDate, "JST", "yyyy-MM-dd"),
          start: formatToHHmm(data[i][3]),
          end: formatToHHmm(data[i][4]),
          note: String(data[i][5]),
          type: data[i][6],
          role: userRole
        });
      }
    }
    return myRes;
  } catch (e) {
    throw new Error("予約データの取得に失敗しました");
  }
}

// 登録済みの授業グループを取得（一括削除用）
function getMyClassGroups(userId) {
  const data = RESERVE_SHEET.getDataRange().getValues();
  const groups = {};
  const dayNames = ["日", "月", "火", "水", "木", "金", "土"];
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId) && data[i][6] === "授業") {
      const className = data[i][5];
      const resDate = new Date(data[i][2]);
      const dayLabel = dayNames[resDate.getDay()] + "曜";
      const startTime = formatToHHmm(data[i][3]);
      
      const key = `${className}_${dayLabel}_${startTime}`;
      
      if (!groups[key]) {
        groups[key] = { name: className, day: dayLabel, start: startTime, count: 0 };
      }
      groups[key].count++;
    }
  }
  return Object.values(groups);
}

// 新規予約
function addReservation(res) {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const start = new Date(res.date + "T" + res.start + ":00");
  const end = new Date(res.date + "T" + res.end + ":00");
  
  const now = new Date();
  if (start < now) throw new Error("過去の日付・時刻には予約できません");
  if (calendar.getEvents(start, end).length > 0) throw new Error("すでに予約が埋まっています");

  let userRole = "学生";
  const memberData = MEMBER_SHEET.getDataRange().getValues();
  for(let i=1; i<memberData.length; i++){
    if(String(memberData[i][0]) === String(res.id)) {
      userRole = memberData[i][2];
      break;
    }
  }

  const title = `【${userRole}】${res.name}`;
  calendar.createEvent(title, start, end, {description: res.note});
  RESERVE_SHEET.appendRow([res.id, res.name, res.date, res.start, res.end, res.note, "通常"]);
  return "予約が完了しました";
}

// 授業一括登録
function registerClass(res) {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const periods = {
    "1": ["08:50", "10:20"], "2": ["10:30", "12:00"], "3": ["13:10", "14:40"],
    "4": ["14:50", "16:20"], "5": ["16:30", "18:00"]
  };
  const times = periods[res.period];
  
  const now = new Date();
  const currentYear = now.getFullYear();
  let startDate, endDate;

  if (res.term === "前期") {
    startDate = new Date(currentYear, 3, 1);
    endDate = new Date(currentYear, 6, 31);
  } else {
    if (now.getMonth() <= 2) {
      startDate = new Date(currentYear - 1, 8, 1);
      endDate = new Date(currentYear, 0, 31);
    } else {
      startDate = new Date(currentYear, 8, 1);
      endDate = new Date(currentYear + 1, 0, 31);
    }
  }

  let checkCursor = new Date(startDate);
  if (checkCursor < now) checkCursor = new Date(now.getTime());
  checkCursor.setHours(0, 0, 0, 0);

  let count = 0;
  while (checkCursor <= endDate) {
    if (checkCursor.getDay() == res.dayOfWeek) {
      let dStr = Utilities.formatDate(checkCursor, "JST", "yyyy-MM-dd");
      let startTime = new Date(dStr + "T" + times[0] + ":00");
      let endTime = new Date(dStr + "T" + times[1] + ":00");
      
      if (startTime > now) {
        // もしその時間に学生の「通常」予約があれば、カレンダーから削除（授業優先）
        const existingEvents = calendar.getEvents(startTime, endTime);
        existingEvents.forEach(e => {
          if (!e.getTitle().includes("[授業]")) e.deleteEvent();
        });

        const title = "[授業] " + res.name;
        calendar.createEvent(title, startTime, endTime, {description: res.note});
        RESERVE_SHEET.appendRow([res.id, res.name, dStr, times[0], times[1], res.note, "授業"]);
        count++;
      }
    }
    checkCursor.setDate(checkCursor.getDate() + 1);
  }
  return count > 0 ? count + "件の授業を登録しました（重複する学生予約は自動解除されます）" : "登録可能な日程がありませんでした";
}

// 授業一括削除
function deleteClassGroup(userId, className, startTime) {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const data = RESERVE_SHEET.getDataRange().getValues();
  let count = 0;

  for (let i = data.length - 1; i >= 1; i--) {
    const sheetStartTime = formatToHHmm(data[i][3]);
    
    if (String(data[i][0]) === String(userId) && 
        data[i][6] === "授業" && 
        data[i][5] === className && 
        sheetStartTime === startTime) {
      
      const resDate = new Date(data[i][2]);
      const events = calendar.getEventsForDay(resDate);
      events.forEach(e => {
        if (Utilities.formatDate(e.getStartTime(), "JST", "HH:mm") === sheetStartTime) {
          e.deleteEvent();
        }
      });

      RESERVE_SHEET.deleteRow(i + 1);
      count++;
    }
  }
  return `${className} (${startTime}〜) の予約を ${count} 件削除しました。`;
}

// 予約取り消し実行
function deleteReservation(id, date, start) {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const events = calendar.getEventsForDay(new Date(date.replace(/-/g, "/")));
  
  let deleted = false;
  events.forEach(e => {
    let eventStart = Utilities.formatDate(e.getStartTime(), "JST", "HH:mm");
    if (eventStart === start) {
      e.deleteEvent();
      deleted = true;
    }
  });

  if (deleted) {
    const data = RESERVE_SHEET.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      let d = Utilities.formatDate(new Date(data[i][2]), "JST", "yyyy-MM-dd");
      let sheetStart = formatToHHmm(data[i][3]);
      if (String(data[i][0]) === String(id) && d === date && sheetStart === start) {
        RESERVE_SHEET.deleteRow(i + 1);
        break;
      }
    }
    return "削除しました";
  }
  throw new Error("予約が見つかりませんでした");
}