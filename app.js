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
const CHINESE_NUMBERS = ["一", "二", "三", "四", "五", "六", "七"];
const SUMMARY_POOL = [
  "\u67e5\u503c\u73ed\u53f0\u74b0\u5883\u3001\u5e38\u7528\u7c3f\u518a\u53ca\u8fa6\u516c\u7269\u54c1\u64fa\u653e\u60c5\u5f62\uff0c\u73fe\u5834\u7dad\u6301\u6574\u6f54\uff0c\u7269\u54c1\u5747\u80fd\u5206\u985e\u6536\u7d0d\u3002",
  "\u67e5\u52e4\u52d9\u7d00\u9304\u3001\u5de5\u4f5c\u7d00\u9304\u53ca\u51fa\u5165\u767b\u8a18\u60c5\u5f62\uff0c\u5404\u9805\u8cc7\u6599\u5747\u80fd\u4f9d\u52e4\u52d9\u57f7\u884c\u60c5\u5f62\u8a18\u8f09\u3002",
  "\u67e5\u5ef3\u820d\u516c\u5171\u7a7a\u9593\u53ca\u5099\u52e4\u5340\u57df\u74b0\u5883\uff0c\u52d5\u7dda\u7dad\u6301\u66a2\u901a\uff0c\u672a\u898b\u660e\u986f\u96dc\u7269\u5806\u7f6e\u6216\u5f71\u97ff\u5b89\u5168\u60c5\u5f62\u3002",
  "\u67e5\u8eca\u8f1b\u4f7f\u7528\u3001\u4fdd\u990a\u53ca\u57fa\u672c\u6aa2\u67e5\u7d00\u9304\uff0c\u76f8\u95dc\u9805\u76ee\u5747\u80fd\u6309\u65e5\u5e38\u7ba1\u7406\u6d41\u7a0b\u8fa6\u7406\u3002",
  "\u67e5\u8eca\u4e0a\u5668\u6750\u3001\u8017\u6750\u53ca\u5fc5\u8981\u88dd\u5099\u6e05\u9ede\u60c5\u5f62\uff0c\u6578\u91cf\u53ca\u64fa\u653e\u4f4d\u7f6e\u5927\u81f4\u7b26\u5408\u52e4\u52d9\u9700\u6c42\u3002",
  "\u67e5\u500b\u4eba\u57fa\u672c\u88dd\u5099\u53ca\u5171\u7528\u88dd\u5099\u7ba1\u7406\u60c5\u5f62\uff0c\u540c\u4ec1\u5747\u80fd\u4f9d\u898f\u5b9a\u653e\u7f6e\u4e26\u7dad\u6301\u53ef\u7528\u72c0\u614b\u3002",
  "\u67e5\u901a\u8a0a\u8a2d\u5099\u3001\u96fb\u529b\u53ca\u57fa\u672c\u6e2c\u8a66\u60c5\u5f62\uff0c\u5404\u9805\u8a2d\u5099\u529f\u80fd\u5c1a\u7a31\u6b63\u5e38\uff0c\u53ef\u652f\u61c9\u4e00\u822c\u52e4\u52d9\u4f7f\u7528\u3002",
  "\u67e5\u672c\u65e5\u8a13\u7df4\u3001\u6f14\u7df4\u6216\u5e38\u614b\u6e2c\u8a66\u5b89\u6392\u60c5\u5f62\uff0c\u53c3\u8207\u4eba\u54e1\u5747\u80fd\u914d\u5408\u57f7\u884c\u4e26\u8a18\u9304\u76f8\u95dc\u6210\u679c\u3002",
  "\u67e5\u696d\u52d9\u627f\u8fa6\u9032\u5ea6\u53ca\u5f85\u8fa6\u4e8b\u9805\u8ffd\u8e64\u60c5\u5f62\uff0c\u5404\u9805\u5de5\u4f5c\u5747\u5df2\u6309\u512a\u5148\u9806\u5e8f\u7ba1\u5236\u8fa6\u7406\u3002",
  "\u67e5\u4eba\u54e1\u5099\u52e4\u3001\u4f11\u606f\u53ca\u57fa\u672c\u5b89\u5168\u7ba1\u7406\u60c5\u5f62\uff0c\u73fe\u5834\u79e9\u5e8f\u826f\u597d\uff0c\u672a\u898b\u660e\u986f\u7570\u5e38\u60c5\u4e8b\u3002"
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
  const count = randomInt(5, 7);
  const picked = shuffle(SUMMARY_POOL).slice(0, count - 2);
  const endHour = (Number(data.hour) + 2) % 24;
  const first = `\u67e5${data.hour}\u6642\u81f3${endHour}\u6642\u503c\u73ed\u4eba\u54e1\uff3f\uff3f\uff3f\uff0c\u4f9d\u898f\u65bc\u503c\u52e4\u53f0\u53d7\u7406\u76f8\u95dc\u6848\u4ef6\uff0c\u8b66\u89ba\u6027\u53ca\u670d\u52d9\u614b\u5ea6\u826f\u597d\uff1b\u672c\u65e5\u4e3b\u7ba1\u70ba\u5206\u968a\u9577\uff3f\uff3f\uff3f\u3002`;
  const second = `\u67e5\u7763\u5c0e\u671f\u9593\u4eba\u54e1\u5728\u968a\u60c5\u5f62\uff0c\u52e4\u52d9\u8868\u986f\u793a\u4e0a\u73ed\u8b66\u6d88${data.firefighterCount}\u540d\u3001\u5f79\u7537${data.conscriptCount}\u540d\uff0c\u7763\u5c0e\u7576\u4e0b\u5747\u5728\u968a\uff0c\u672a\u6709\u9055\u53cd\u52e4\u4f11\u6216\u98f2\u9152\u60c5\u4e8b\u3002`;
  return [first, second, ...picked].map((item, index) => `${CHINESE_NUMBERS[index]}\u3001${item}`);
}

function shuffle(items) {
  const copy = [...items];
  for (let i = copy.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [copy[i], copy[j]] = [copy[j], copy[i]];
  }
  return copy;
}

function randomInt(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
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
