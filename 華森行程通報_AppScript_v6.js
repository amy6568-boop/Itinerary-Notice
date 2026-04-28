// ============================================================
//  華森行程通報 — Google Apps Script  v5
//
//  欄位對照（行程通知分頁，第3列起為資料）：
//  A=日期  B=類型  C=行程主題  D=開始時間  E=結束時間
//  F=地點  G=參與人員(名)  H=通知群組(名)  I=備註
//  J=固定通知(是/否)  K=固定日期(月/日)
//  L=啟用7天前(是/否)  M=啟用1天前(是/否)  N=啟用當天(是/否)
//  O=7天前已發(是/否)  P=1天前已發(是/否)  Q=當天已發(是/否)
//  R=距今天數(公式)  S=通知狀態(公式)
// ============================================================

const CONFIG = {
  SHEET_NAME:   '行程通知',
  ID_SHEET:     '人員ID對照',
  TOKEN:        'iuXvq2wyXMqpYntDy1WIaH08AmkXzjx67Wk/6KRlDyOq/9P99eEeC9GnUJrOpILvHbjma4/vXEot5Cx9mfzUAFaKeBw8grfbynVDJ6rnMQlcn8w/uJHVaFORZ6xifOrTGIhRmo1BEcAPmcVruaAY0AdB04t89/1O/w1cDnyilFU=',
  FIXED_NOTIFY: [
    'U1b19067c31898048c56de0943859d2b6',  // 鄭惠方
    'U6c58e101842b4bbce5b5c43c410406f6',  // 林香君
  ],
};

function doGet(e) {
  const action = (e.parameter.action || '').trim();
  if (action === 'ping')       return out('pong');
  if (action === 'add')        return handleAdd(e.parameter);
  if (action === 'getMembers') return handleGetMembers();
  return out(JSON.stringify({ error: 'unknown action' }));
}

function handleGetMembers() {
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.ID_SHEET)?.getDataRange().getValues();
  if (!data) return out(JSON.stringify({ ok:true, people:[], groups:[], types:[] }));

  const people = [], groups = [], types = [];
  let mode = null;

  for (let i = 0; i < data.length; i++) {
    const cell = String(data[i][0]||'').trim();
    if (cell === '【個人 userId】')  { mode='people'; continue; }
    if (cell === '【群組 groupId】') { mode='groups'; continue; }
    if (!cell || cell.startsWith('💡') || cell==='姓名/暱稱' || cell==='群組名稱') continue;
    if (mode==='people') {
      const userId = String(data[i][1]||'').trim().replace(/[\r\n\s]/g,'');
      people.push({ name:cell, userId, title:String(data[i][3]||'').trim() });
    } else if (mode==='groups') {
      const groupId = String(data[i][1]||'').trim().replace(/[\r\n\s]/g,'');
      groups.push({ name:cell, groupId });
    }
  }

  // 讀取類型清單
  const tData = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('類型清單')?.getDataRange().getValues();
  if (tData) {
    for (let i=2; i<tData.length; i++) {
      const t = String(tData[i][0]||'').trim();
      if (t) types.push(t);
    }
  }

  return out(JSON.stringify({ ok:true, people, groups, types }));
}

function handleAdd(p) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  sheet.appendRow([
    p.date||'', p.type||'', p.title||'',
    p.timeStart||'', p.timeEnd||'', p.location||'',
    p.members||'', p.groups||'', p.note||'',
    p.fixed||'否', p.fixedDate||'',
    p.notify7||'否', p.notify1||'否', p.notifyD||'否',
    '否','否','否',
  ]);
  const msg = buildImmediateMsg(p);
  CONFIG.FIXED_NOTIFY.filter(Boolean).forEach(uid => sendLine(uid, msg));
  resolveUserIds(p.members||'').forEach(uid => {
    if (!CONFIG.FIXED_NOTIFY.includes(uid)) sendLine(uid, msg);
  });
  resolveGroupIds(p.groups||'').forEach(gid => sendLine(gid, msg));
  return out(JSON.stringify({ ok:true }));
}

function checkAndNotify() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const data  = sheet.getDataRange().getValues();
  const today = new Date(); today.setHours(0,0,0,0);

  for (let i=2; i<data.length; i++) {
    const row = data[i];
    const isFixed = String(row[9]).trim()==='是';
    if (!row[0] && !isFixed) continue;

    const p = {
      date:row[0], type:row[1], title:row[2],
      timeStart:row[3], timeEnd:row[4], location:row[5],
      members:row[6], groups:row[7], note:row[8],
      fixed:row[9], fixedDate:row[10],
    };

    if (isFixed && row[10]) {
      processFixedNotify(sheet, i, row, p, today);
      continue;
    }
    if (!row[0]) continue;

    const ed = new Date(row[0]); ed.setHours(0,0,0,0);
    const diff = Math.round((ed-today)/86400000);
    const en7=String(row[11]).trim()==='是', en1=String(row[12]).trim()==='是', enD=String(row[13]).trim()==='是';
    const s7 =String(row[14]).trim()==='是', s1 =String(row[15]).trim()==='是', sD =String(row[16]).trim()==='是';

    let label=null, sentCol=null;
    if      (diff===7&&en7&&!s7){label='📅 提前 7 天提醒';sentCol=15;}
    else if (diff===1&&en1&&!s1){label='⏰ 明天提醒';      sentCol=16;}
    else if (diff===0&&enD&&!sD){label='🔔 今日行程';      sentCol=17;}
    if (!label) continue;

    sendToAll(buildScheduledMsg(label,p), p);
    sheet.getRange(i+1,sentCol).setValue('是');
  }
}

function processFixedNotify(sheet, i, row, p, today) {
  const parts = String(row[10]).trim().split('/');
  if (parts.length!==2) return;
  const thisYear = today.getFullYear();
  const target   = new Date(thisYear, parseInt(parts[0])-1, parseInt(parts[1]));
  target.setHours(0,0,0,0);
  const diff = Math.round((target-today)/86400000);

  if (today.getMonth()===0 && today.getDate()===1) {
    sheet.getRange(i+1,15).setValue('否');
    sheet.getRange(i+1,16).setValue('否');
    sheet.getRange(i+1,17).setValue('否');
  }

  const en7=String(row[11]).trim()==='是', en1=String(row[12]).trim()==='是', enD=String(row[13]).trim()==='是';
  const s7 =String(row[14]).trim()==='是', s1 =String(row[15]).trim()==='是', sD =String(row[16]).trim()==='是';

  let label=null, sentCol=null;
  if      (diff===7&&en7&&!s7){label='📅【固定】提前 7 天提醒';sentCol=15;}
  else if (diff===1&&en1&&!s1){label='⏰【固定】明天提醒';      sentCol=16;}
  else if (diff===0&&enD&&!sD){label='🔔【固定】今日行程';      sentCol=17;}
  if (!label) return;

  const displayDate = `${thisYear}/${parts[0]}/${parts[1]}`;
  sendToAll(buildScheduledMsg(label,{...p,date:displayDate}), p);
  sheet.getRange(i+1,sentCol).setValue('是');
}

function sendToAll(msg, p) {
  CONFIG.FIXED_NOTIFY.filter(Boolean).forEach(uid => sendLine(uid,msg));
  resolveUserIds(String(p.members||'')).forEach(uid => {
    if (!CONFIG.FIXED_NOTIFY.includes(uid)) sendLine(uid,msg);
  });
  resolveGroupIds(String(p.groups||'')).forEach(gid => sendLine(gid,msg));
}

function resolveUserIds(namesStr) {
  if (!namesStr.trim()) return [];
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.ID_SHEET)?.getDataRange().getValues();
  if (!data) return [];
  const map = {}; let in_ = false;
  for (let i=0;i<data.length;i++) {
    const c=String(data[i][0]||'').trim();
    if (c==='【個人 userId】'){in_=true;continue;}
    if (c==='【群組 groupId】') break;
    if (!in_||!c||c==='姓名/暱稱') continue;
    const uid=String(data[i][1]||'').trim().replace(/[\r\n\s]/g,'');
    if (uid) map[c]=uid;
  }
  return namesStr.split(',').map(n=>map[n.trim()]).filter(Boolean);
}

function resolveGroupIds(namesStr) {
  if (!namesStr.trim()) return [];
  const data = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(CONFIG.ID_SHEET)?.getDataRange().getValues();
  if (!data) return [];
  const map = {}; let in_ = false;
  for (let i=0;i<data.length;i++) {
    const c=String(data[i][0]||'').trim();
    if (c==='【群組 groupId】'){in_=true;continue;}
    if (!in_||!c||c==='群組名稱'||c.startsWith('💡')) continue;
    const gid=String(data[i][1]||'').trim().replace(/[\r\n\s]/g,'');
    if (gid) map[c]=gid;
  }
  return namesStr.split(',').map(n=>map[n.trim()]).filter(Boolean);
}

function buildImmediateMsg(p) {
  const notify=[];
  if (p.fixed==='是')   notify.push(`🔁 固定（每年${p.fixedDate}）`);
  if (p.notify7==='是') notify.push('7天前');
  if (p.notify1==='是') notify.push('1天前');
  if (p.notifyD==='是') notify.push('當天');
  const t=(p.timeStart&&p.timeEnd)?`${p.timeStart}–${p.timeEnd}`:p.timeStart||p.timeEnd||'待確認';
  return ['📋 新行程通報','━━━━━━━━━━━━━',
    `📌 類型：${p.type}`,`📝 主題：${p.title}`,`📆 日期：${p.date}`,
    `🕐 時間：${t}`,`📍 地點：${p.location||'待確認'}`,
    `👥 人員：${p.members||'無'}`,`🔔 通知時機：${notify.join('、')||'無'}`,
    p.note?`💬 備註：${p.note}`:'',
  ].filter(Boolean).join('\n');
}

function buildScheduledMsg(label, p) {
  const ds=p.date instanceof Date?Utilities.formatDate(p.date,'Asia/Taipei','yyyy/MM/dd'):String(p.date);
  const ts=p.timeStart instanceof Date?Utilities.formatDate(p.timeStart,'Asia/Taipei','HH:mm'):String(p.timeStart||'');
  const te=p.timeEnd   instanceof Date?Utilities.formatDate(p.timeEnd,  'Asia/Taipei','HH:mm'):String(p.timeEnd  ||'');
  const t =ts&&te?`${ts}–${te}`:ts||te||'待確認';
  return [label,'━━━━━━━━━━━━━',
    `📌 類型：${p.type}`,`📝 主題：${p.title}`,`📆 日期：${ds}`,
    `🕐 時間：${t}`,`📍 地點：${p.location||'待確認'}`,
    `👥 人員：${p.members||'無'}`,
    p.note?`💬 備註：${p.note}`:'',
  ].filter(Boolean).join('\n');
}

function sendLine(to, text) {
  if (!to||!text) return;
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push',{
      method:'post', contentType:'application/json',
      headers:{'Authorization':'Bearer '+CONFIG.TOKEN},
      payload:JSON.stringify({to,messages:[{type:'text',text}]}),
      muteHttpExceptions:true,
    });
  } catch(e){Logger.log('LINE失敗：'+e.message);}
}

function out(text) {
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function testPing() {
  Logger.log('✓ 連線正常');
  const members = JSON.parse(handleGetMembers().getContent());
  Logger.log('人員：'+members.people.map(p=>p.name).join(', '));
  Logger.log('群組：'+members.groups.map(g=>g.name).join(', '));
  sendLine(CONFIG.FIXED_NOTIFY[0],'🔌 華森行程通報 v5 連線測試成功！');
}
