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
  let xml = xmlFile.asText();
  xml = replaceTextNodes(xml, {
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
  });

  zip.file(xmlPath, xml);
  return zip.generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });
}

function replaceTextNodes(xml, replacements) {
  let index = 0;
  const seen = new Set();
  const nextXml = xml.replace(/(<(?:\w+:)?t\b[^>]*>)([\s\S]*?)(<\/(?:\w+:)?t>)/g, (match, open, text, close) => {
    if (!Object.prototype.hasOwnProperty.call(replacements, index)) {
      index += 1;
      return match;
    }

    const value = escapeXml(String(replacements[index]));
    seen.add(index);
    index += 1;
    return `${open}${value}${close}`;
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
