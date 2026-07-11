const TEMPLATE_PATH = "template.docx?v=20260711-fix2";
const TEXT_NODE_MAP = {
  rocCenturyPrefix: 21,
  rocYearLastDigit: 22,
  month: 24,
  day: 26,
  hour: 28,
  minuteTens: 30,
  minuteOnes: 31,
  supervisor: 11,
  unitName: 33,
  unitSuffix: 34
};
const SUMMARY_NODES = [35, 46, 54, 69, 79, 80, 82];
const SUMMARY_CLEAR_NODES = [
  36, 37, 38, 39, 40, 41, 42, 43, 44, 45,
  47, 48, 49, 50, 51, 52, 53,
  55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68,
  70, 71, 72, 73, 74, 75, 76, 77, 78,
  81, 83, 84, 85
];
const HANDLING_NODES = [86, 88, 89, 90, 91];
const SUMMARY_POOL = [
  "查值班台環境、常用簿冊及辦公物品擺放情形，現場維持整潔，物品均能分類收納。",
  "查勤務紀錄、工作紀錄及出入登記情形，各項資料均能依勤務執行情形記載。",
  "查廳舍公共空間及備勤區域環境，動線維持暢通，未見明顯雜物堆置或影響安全情形。",
  "查車輛使用、保養及基本檢查紀錄，相關項目均能按日常管理流程辦理。",
  "查車上器材、耗材及必要裝備清點情形，數量及擺放位置大致符合勤務需求。",
  "查個人基本裝備及共用裝備管理情形，同仁均能依規定放置並維持可用狀態。",
  "查通訊設備、電力及基本測試情形，各項設備功能尚稱正常，可支應一般勤務使用。",
  "查本日訓練、演練或常態測試安排情形，參與人員均能配合執行並記錄相關成果。",
  "查業務承辦進度及待辦事項追蹤情形，各項工作均已按優先順序管制辦理。",
  "查人員備勤、休息及基本安全管理情形，現場秩序良好，未見明顯異常情事。",
  "查防颱、防汛及災害整備情形，救災器材、油料、照明及後勤物資均已完成基本檢視。",
  "查轄內易淹水地區、低窪地區及災害潛勢資料掌握情形，相關清冊均有持續更新及備查。",
  "查抽水機、發電機、鏈鋸及照明設備功能測試情形，設備可正常啟動並維持待命狀態。",
  "查水域救援裝備保養及清點情形，救生裝備、繩索及相關器材均有依規定整理備用。",
  "查救護訓練辦理情形，訓練內容包含基本流程、團隊分工及器材操作，參訓人員均能配合演練。",
  "查救護車內部整潔與器耗材管理情形，常用耗材均已分類放置，保存及補充情形尚稱良好。",
  "查消防查察及水源查察勤務排定情形，相關案件均有依期程追蹤，尚未完成者已列管續辦。",
  "查自衛消防編組、集合住宅宣導或防火宣導案件辦理情形，承辦進度已依業務需求持續追蹤。",
  "查勤務表編排及人力調度情形，各班別安排尚能兼顧勤務需求及人員休息。",
  "查外出、外宿、支援或公假人員登記情形，相關紀錄均能與勤務安排相互勾稽。",
  "查值宿及備勤人員服勤情形，相關人員均能依規定於指定地點待命，未見擅離情形。",
  "查隊員＿＿＿辦理指定業務情形，相關資料已依規定建檔，後續進度仍請持續追蹤。",
  "查小隊長＿＿＿帶領同仁執行訓練或勤務整備情形，現場操作流程清楚，並能提醒安全注意事項。",
  "查役男＿＿＿協助值班、環境維護或行政庶務情形，均能依指示辦理，服務態度尚佳。",
  "查分隊長＿＿＿督導勤業務及人員管理情形，對於重點事項均能適時提醒並列入追蹤。",
  "查無線電、手提台及相關通訊備品管理情形，測試結果尚稱正常，故障或異常項目已列管處理。",
  "查車輛里程、油料、水箱水及基本電路測試情形，相關項目均能配合日常檢查辦理。",
  "查空氣呼吸器、氣瓶壓力及搶救器材整備情形，檢查結果尚符合基本出勤需求。",
  "查化學品資料查詢、危害辨識及相關資訊平台熟悉情形，受測人員均能依程序完成查詢。",
  "查同仁自主健康紀錄及安全衛生管理情形，相關量測或紀錄資料均有依規定保存。",
  "查廳舍牆面、漏水、照明、門窗及設備修繕需求，已請承辦人員列管並依程序陳報處理。",
  "查婦宣、義消或民力協勤訓練聯繫情形，相關活動均有依計畫安排並留存紀錄。",
  "查節能管理情形，未使用區域燈光及電源均能適時關閉，水電設備未見明顯異常。",
  "查消防車、救護車及勤務車輛內外部清潔情形，車體及車內器材均維持可出勤狀態。",
  "查責任區業務、待確認案件及系統通報事項辦理情形，承辦人員已依期程追蹤並回報進度。"
];

const form = document.querySelector("#reportForm");
const statusEl = document.querySelector("#status");
const generateBtn = document.querySelector("#generateBtn");
const dateInput = document.querySelector("#reportDate");
const timeInput = document.querySelector("#reportTime");

const now = new Date();
dateInput.value = toDateInputValue(now);
timeInput.value = toTimeInputValue(now);

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  setBusy(true, "\u7522\u751f\u4e2d...");

  try {
    const data = readForm();
    const template = await fetchTemplate();
    const output = buildDocx(template, data);
    downloadBlob(output, makeFileName(data));
    setBusy(false, "\u5df2\u7522\u751f Word \u6a94\u3002");
  } catch (error) {
    console.error(error);
    setBusy(false, `\u7522\u751f\u5931\u6557\uFF1A${error.message}`);
  }
});

function readForm() {
  const date = new Date(`${dateInput.value}T00:00:00`);
  const [hour, minute] = timeInput.value.split(":").map(Number);
  const rocYear = date.getFullYear() - 1911;
  const supervisor = document.querySelector("#supervisor").value.trim();
  const unit = document.querySelector("#unit").value.trim();
  const firefighterCount = document.querySelector("#firefighterCount").value.trim();
  const conscriptCount = document.querySelector("#conscriptCount").value.trim();

  return {
    rocYear: String(rocYear),
    month: String(date.getMonth() + 1),
    day: String(date.getDate()),
    hour: String(hour),
    minute: String(minute).padStart(2, "0"),
    supervisor,
    unit,
    firefighterCount,
    conscriptCount
  };
}

async function fetchTemplate() {
  const response = await fetch(TEMPLATE_PATH, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`template.docx HTTP ${response.status}`);
  }
  return await response.arrayBuffer();
}

function buildDocx(templateBuffer, data) {
  const zip = new PizZip(templateBuffer);
  const xmlPath = "word/document.xml";
  const xmlFile = zip.file(xmlPath);
  if (!xmlFile) {
    throw new Error("word/document.xml not found");
  }

  const unitParts = splitUnit(data.unit);
  const summaryItems = buildSummaryItems(data);
  let xml = xmlFile.asText();
  const replacements = {
    [TEXT_NODE_MAP.rocCenturyPrefix]: data.rocYear.slice(0, -1) || "0",
    [TEXT_NODE_MAP.rocYearLastDigit]: data.rocYear.slice(-1),
    [TEXT_NODE_MAP.month]: data.month,
    [TEXT_NODE_MAP.day]: data.day,
    [TEXT_NODE_MAP.hour]: data.hour,
    [TEXT_NODE_MAP.minuteTens]: data.minute[0],
    [TEXT_NODE_MAP.minuteOnes]: data.minute[1],
    [TEXT_NODE_MAP.supervisor]: data.supervisor,
    [TEXT_NODE_MAP.unitName]: unitParts.name,
    [TEXT_NODE_MAP.unitSuffix]: unitParts.suffix
  };

  for (const node of SUMMARY_CLEAR_NODES) {
    replacements[node] = "";
  }
  for (let i = 0; i < SUMMARY_NODES.length; i += 1) {
    replacements[SUMMARY_NODES[i]] = summaryItems[i] || "";
  }
  replacements[HANDLING_NODES[0]] = "\u64ec\uff1a\u6587\u5b58\u67e5\u3002";
  for (const node of HANDLING_NODES.slice(1)) {
    replacements[node] = "";
  }

  xml = replaceTextNodes(xml, replacements);

  zip.file(xmlPath, xml);
  return zip.generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });
}

function buildSummaryItems(data) {
  const picked = shuffle(SUMMARY_POOL).slice(0, 5);
  const dutyRange = getDutyRange(Number(data.hour));
  const first = `\u67e5${dutyRange}\u503c\u73ed\u4eba\u54e1\uff3f\uff3f\uff3f\uff0c\u4f9d\u898f\u65bc\u503c\u52e4\u53f0\u53d7\u7406\u76f8\u95dc\u6848\u4ef6\uff0c\u8b66\u89ba\u6027\u53ca\u670d\u52d9\u614b\u5ea6\u826f\u597d\uff1b\u672c\u65e5\u4e3b\u7ba1\u70ba\u5206\u968a\u9577\uff3f\uff3f\uff3f\u3002`;
  const second = `\u67e5\u7763\u5c0e\u671f\u9593\u4eba\u54e1\u5728\u968a\u60c5\u5f62\uff0c\u52e4\u52d9\u8868\u986f\u793a\u4e0a\u73ed\u8b66\u6d88${data.firefighterCount}\u540d\u3001\u5f79\u7537${data.conscriptCount}\u540d\uff0c\u7763\u5c0e\u7576\u4e0b\u5747\u5728\u968a\uff0c\u672a\u6709\u9055\u53cd\u52e4\u4f11\u6216\u98f2\u9152\u60c5\u4e8b\u3002`;
  return [first, second, ...picked];
}

function getDutyRange(hour) {
  const start = Math.floor(hour / 2) * 2;
  const end = (start + 2) % 24;
  return end === 0 ? `${start}\u6642\u81f3\u7fcc\u65e50\u6642` : `${start}\u6642\u81f3${end}\u6642`;
}

function shuffle(items) {
  const copy = [...items];
  for (let i = copy.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy;
}

function replaceTextNodes(xml, replacements) {
  let index = 0;
  const seen = new Set();
  const textNodePattern = /<((?:\w+:)?t)\b([^>]*?)\/>|<((?:\w+:)?t)\b([^>]*?)>([\s\S]*?)<\/\3>/g;
  const nextXml = xml.replace(textNodePattern, (match, selfTag, selfAttrs, tag, attrs, text) => {
    if (!Object.prototype.hasOwnProperty.call(replacements, index)) {
      index += 1;
      return match;
    }

    const value = escapeXml(String(replacements[index]));
    seen.add(index);
    index += 1;
    const nodeName = selfTag || tag;
    const nodeAttrs = selfAttrs || attrs || "";
    return `<${nodeName}${nodeAttrs}>${value}</${nodeName}>`;
  });

  for (const key of Object.keys(replacements)) {
    if (!seen.has(Number(key))) {
      throw new Error(`template text node ${key} not found`);
    }
  }

  return nextXml;
}

function escapeXml(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function splitUnit(unit) {
  const suffix = "\u5206\u968a";
  if (unit.endsWith(suffix)) {
    return { name: unit.slice(0, -suffix.length), suffix };
  }
  return { name: unit, suffix: "" };
}

function makeFileName(data) {
  const mm = data.month.padStart(2, "0");
  const dd = data.day.padStart(2, "0");
  return `${data.unit}${mm}${dd}${data.supervisor}.docx`;
}

function downloadBlob(blob, fileName) {
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
}

function toDateInputValue(date) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function toTimeInputValue(date) {
  const hh = String(date.getHours()).padStart(2, "0");
  const mm = String(date.getMinutes()).padStart(2, "0");
  return `${hh}:${mm}`;
}

function setBusy(isBusy, message) {
  generateBtn.disabled = isBusy;
  statusEl.textContent = message;
}
