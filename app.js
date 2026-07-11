const TEMPLATE_PATH = "template.docx";
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
  setBusy(true, "產生中...");

  try {
    const data = readForm();
    const template = await fetchTemplate();
    const output = buildDocx(template, data);
    downloadBlob(output, makeFileName(data));
    setBusy(false, "已產生 Word 檔。");
  } catch (error) {
    console.error(error);
    setBusy(false, "產生失敗，請確認 template.docx 與網路 CDN 可讀取。");
  }
});

function readForm() {
  const date = new Date(`${dateInput.value}T00:00:00`);
  const [hour, minute] = timeInput.value.split(":").map(Number);
  const rocYear = date.getFullYear() - 1911;
  const supervisor = document.querySelector("#supervisor").value.trim();
  const unit = document.querySelector("#unit").value.trim();

  return {
    rocYear: String(rocYear),
    month: String(date.getMonth() + 1),
    day: String(date.getDate()),
    hour: String(hour),
    minute: String(minute).padStart(2, "0"),
    supervisor,
    unit
  };
}

async function fetchTemplate() {
  const response = await fetch(TEMPLATE_PATH);
  if (!response.ok) {
    throw new Error(`Template fetch failed: ${response.status}`);
  }
  return await response.arrayBuffer();
}

function buildDocx(templateBuffer, data) {
  const zip = new PizZip(templateBuffer);
  const xmlPath = "word/document.xml";
  const parser = new DOMParser();
  const serializer = new XMLSerializer();
  const xml = zip.file(xmlPath).asText();
  const documentXml = parser.parseFromString(xml, "application/xml");
  const textNodes = Array.from(documentXml.getElementsByTagName("w:t"));
  const unitParts = splitUnit(data.unit);

  setText(textNodes, TEXT_NODE_MAP.rocCenturyPrefix, data.rocYear.slice(0, -1) || "0");
  setText(textNodes, TEXT_NODE_MAP.rocYearLastDigit, data.rocYear.slice(-1));
  setText(textNodes, TEXT_NODE_MAP.month, data.month);
  setText(textNodes, TEXT_NODE_MAP.day, data.day);
  setText(textNodes, TEXT_NODE_MAP.hour, data.hour);
  setText(textNodes, TEXT_NODE_MAP.minuteTens, data.minute[0]);
  setText(textNodes, TEXT_NODE_MAP.minuteOnes, data.minute[1]);
  setText(textNodes, TEXT_NODE_MAP.supervisor, data.supervisor);
  setText(textNodes, TEXT_NODE_MAP.unitName, unitParts.name);
  setText(textNodes, TEXT_NODE_MAP.unitSuffix, unitParts.suffix);

  zip.file(xmlPath, serializer.serializeToString(documentXml));
  return zip.generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });
}

function setText(textNodes, index, value) {
  if (!textNodes[index]) {
    throw new Error(`Template text node ${index} not found`);
  }
  textNodes[index].textContent = value;
}

function splitUnit(unit) {
  if (unit.endsWith("分隊")) {
    return { name: unit.slice(0, -2), suffix: "分隊" };
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
