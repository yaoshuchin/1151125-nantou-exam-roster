const XLSX = require('xlsx');
const fs = require('fs');

// Read data.js indirectly or use data definitions
const wb = XLSX.utils.book_new();

const examDuties = [
  // 國文
  { subject: "國文", grade: "一年級", exam1_setter: "鄭淑芬", exam1_reviewer: "林姵辰", exam2_setter: "林姵辰(作文)", exam2_reviewer: "黃全帆", exam3_setter: "黃全帆", exam3_reviewer: "鄭淑芬", makeup_setter: "林志恩" },
  { subject: "國文", grade: "二年級", exam1_setter: "黃惠君(作文)", exam1_reviewer: "連英式", exam2_setter: "連英式(作文)", exam2_reviewer: "簡雨緹", exam3_setter: "簡雨緹", exam3_reviewer: "黃惠君", makeup_setter: "林志恩" },
  { subject: "國文", grade: "三年級", exam1_setter: "黃珮姿", exam1_reviewer: "蔡依秀", exam2_setter: "蔡依秀", exam2_reviewer: "施惠婷", exam3_setter: "施惠婷", exam3_reviewer: "黃珮姿", makeup_setter: "林志恩" },
  // 英語
  { subject: "英語", grade: "一年級", exam1_setter: "簡鈴燕", exam1_reviewer: "周學琪", exam2_setter: "周學琪", exam2_reviewer: "簡蘚讌", exam3_setter: "簡蘚讌", exam3_reviewer: "簡鈴燕", makeup_setter: "王政茹" },
  { subject: "英語", grade: "二年級", exam1_setter: "許瓊芳", exam1_reviewer: "黃瓊慧", exam2_setter: "黃瓊慧", exam2_reviewer: "許招玉", exam3_setter: "許招玉", exam3_reviewer: "許瓊芳", makeup_setter: "王政茹" },
  { subject: "英語", grade: "三年級", exam1_setter: "蔡宜家", exam1_reviewer: "曾佳珍", exam2_setter: "曾佳珍", exam2_reviewer: "古士傑", exam3_setter: "古士傑", exam3_reviewer: "蔡宜家", makeup_setter: "王政茹" },
  // 數學
  { subject: "數學", grade: "一年級", exam1_setter: "黃香喻", exam1_reviewer: "劉珈妤", exam2_setter: "劉珈妤", exam2_reviewer: "林美貞", exam3_setter: "林美貞", exam3_reviewer: "黃香喻", makeup_setter: "林儀婷" },
  { subject: "數學", grade: "二年級", exam1_setter: "林富銓", exam1_reviewer: "劉先雲", exam2_setter: "劉先雲", exam2_reviewer: "林孟嶸", exam3_setter: "林孟嶸", exam3_reviewer: "林富銓", makeup_setter: "林儀婷" },
  { subject: "數學", grade: "三年級", exam1_setter: "陳秉綸", exam1_reviewer: "黃瓊英", exam2_setter: "黃瓊英", exam2_reviewer: "劉彩宜", exam3_setter: "劉彩宜", exam3_reviewer: "陳秉綸", makeup_setter: "林儀婷" },
  // 自然/理化/地科
  { subject: "自然", grade: "一年級", exam1_setter: "許昭雅", exam1_reviewer: "伍子善", exam2_setter: "伍子善", exam2_reviewer: "蔡惠玲", exam3_setter: "許昭雅", exam3_reviewer: "伍子善", makeup_setter: "伍子善" },
  { subject: "自然", grade: "二年級", exam1_setter: "許正儀", exam1_reviewer: "林鈺展", exam2_setter: "林鈺展", exam2_reviewer: "白能宗", exam3_setter: "蔡惠玲", exam3_reviewer: "許正儀", makeup_setter: "蔡惠玲" },
  { subject: "理化", grade: "三年級", exam1_setter: "林筱琦", exam1_reviewer: "羅丞書", exam2_setter: "羅丞書", exam2_reviewer: "羅丞書", exam3_setter: "白能宗", exam3_reviewer: "林筱琦", makeup_setter: "林筱琦" },
  // 歷史
  { subject: "歷史", grade: "一年級", exam1_setter: "黃琇琴", exam1_reviewer: "藍承莉", exam2_setter: "黃琇琴", exam2_reviewer: "藍承莉", exam3_setter: "葉春楓", exam3_reviewer: "藍承莉", makeup_setter: "胡怡妙" },
  { subject: "歷史", grade: "二年級", exam1_setter: "藍承莉", exam1_reviewer: "黃琇琴", exam2_setter: "藍承莉", exam2_reviewer: "黃琇琴", exam3_setter: "藍承莉", exam3_reviewer: "藍承莉", makeup_setter: "林曉雯" },
  { subject: "歷史", grade: "三年級", exam1_setter: "黃琇琴", exam1_reviewer: "藍承莉", exam2_setter: "藍承莉", exam2_reviewer: "黃琇琴", exam3_setter: "藍承莉", exam3_reviewer: "黃琇琴", makeup_setter: "李佩儒" },
  // 地理
  { subject: "地理", grade: "一年級", exam1_setter: "林曉雯", exam1_reviewer: "胡怡妙", exam2_setter: "林曉雯", exam2_reviewer: "胡怡妙", exam3_setter: "林曉雯", exam3_reviewer: "胡怡妙", makeup_setter: "胡怡妙" },
  { subject: "地理", grade: "二年級", exam1_setter: "孫雅惠", exam1_reviewer: "孫雅惠", exam2_setter: "孫雅惠", exam2_reviewer: "孫雅惠", exam3_setter: "孫雅惠", exam3_reviewer: "孫雅惠", makeup_setter: "李佩儒" },
  { subject: "地理", grade: "三年級", exam1_setter: "胡怡妙", exam1_reviewer: "林曉雯", exam2_setter: "胡怡妙", exam2_reviewer: "林曉雯", exam3_setter: "胡怡妙", exam3_reviewer: "林曉雯", makeup_setter: "林曉雯" },
  // 公民
  { subject: "公民", grade: "一年級", exam1_setter: "張君端", exam1_reviewer: "李協融", exam2_setter: "李協融", exam2_reviewer: "張君端", exam3_setter: "張君端", exam3_reviewer: "李協融", makeup_setter: "林曉雯" },
  { subject: "公民", grade: "二年級", exam1_setter: "李佩儒", exam1_reviewer: "簡雨緹", exam2_setter: "李協融", exam2_reviewer: "簡雨緹", exam3_setter: "簡雨緹", exam3_reviewer: "李佩儒", makeup_setter: "李佩儒" },
  { subject: "公民", grade: "三年級", exam1_setter: "李協融", exam1_reviewer: "吳庭萱", exam2_setter: "吳庭萱", exam2_reviewer: "李協融", exam3_setter: "李協融", exam3_reviewer: "吳庭萱", makeup_setter: "李佩儒" },
];

// Sheet 1: 整合總表
const s1Data = [
  ["南投國中 115學年度第1學期 各科評量命題與審校輪值整合表"],
  ["科目", "年級", "第一次段考(命題)", "第一次段考(審校)", "第二次段考(命題)", "第二次段考(審校)", "第三次段考(命題)", "第三次段考(審校)", "補考命題"]
];
examDuties.forEach(r => {
  s1Data.push([r.subject, r.grade, r.exam1_setter, r.exam1_reviewer, r.exam2_setter, r.exam2_reviewer, r.exam3_setter, r.exam3_reviewer, r.makeup_setter]);
});
const ws1 = XLSX.utils.aoa_to_sheet(s1Data);
XLSX.utils.book_append_sheet(wb, ws1, "輪值整合總表");

// Sheet 2: 命題一覽表
const s2Data = [
  ["南投國中 115 學年度第 1 學期各科評量命題教師一覽表"],
  ["科目", "年級", "第一次段考", "第二次段考", "第三次段考", "補考"]
];
examDuties.forEach(r => {
  s2Data.push([r.subject, r.grade, r.exam1_setter, r.exam2_setter, r.exam3_setter, r.makeup_setter]);
});
const ws2 = XLSX.utils.aoa_to_sheet(s2Data);
XLSX.utils.book_append_sheet(wb, ws2, "命題教師一覽表");

// Sheet 3: 審校一覽表
const s3Data = [
  ["南投國中 115 學年第 1 學期各科評量試卷審校教師一覽表"],
  ["科目", "年級", "第一次段考", "第二次段考", "第三次段考"]
];
examDuties.forEach(r => {
  s3Data.push([r.subject, r.grade, r.exam1_reviewer, r.exam2_reviewer, r.exam3_reviewer]);
});
const ws3 = XLSX.utils.aoa_to_sheet(s3Data);
XLSX.utils.book_append_sheet(wb, ws3, "審校教師一覽表");

// Sheet 4: 教師個人輪值表
const teacherMap = {};
examDuties.forEach(row => {
  const process = (rawName, role, exam) => {
    const isComposition = rawName.includes('作文');
    const name = rawName.replace(/\(作文\)/g, '').trim();
    if (!name) return;
    if (!teacherMap[name]) {
      teacherMap[name] = { name, setterCount: 0, reviewerCount: 0, makeupCount: 0, total: 0, duties: [] };
    }
    teacherMap[name].duties.push(`${exam}${row.subject}${row.grade}(${role}${isComposition ? '-作文' : ''})`);
    if (role === '命題') teacherMap[name].setterCount++;
    else if (role === '審校') teacherMap[name].reviewerCount++;
    else if (role === '補考') teacherMap[name].makeupCount++;
    teacherMap[name].total++;
  };
  process(row.exam1_setter, '命題', '第1次');
  process(row.exam1_reviewer, '審校', '第1次');
  process(row.exam2_setter, '命題', '第2次');
  process(row.exam2_reviewer, '審校', '第2次');
  process(row.exam3_setter, '命題', '第3次');
  process(row.exam3_reviewer, '審校', '第3次');
  process(row.makeup_setter, '補考', '補考');
});

const s4Data = [
  ["南投國中 115學年度第1學期 教師個人評量輪值清單與任務總覽"],
  ["教師姓名", "任務總量", "命題次數", "審校次數", "補考次數", "詳細任務項目"]
];
Object.values(teacherMap).sort((a,b) => b.total - a.total).forEach(t => {
  s4Data.push([t.name, t.total, t.setterCount, t.reviewerCount, t.makeupCount, t.duties.join("； ")]);
});
const ws4 = XLSX.utils.aoa_to_sheet(s4Data);
XLSX.utils.book_append_sheet(wb, ws4, "教師個人任務清單");

// Write out to file
XLSX.writeFile(wb, "115南投國中各科評量教師輪值表.xlsx");
console.log("Generated 115南投國中各科評量教師輪值表.xlsx successfully!");
