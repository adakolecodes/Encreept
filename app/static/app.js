const MAX_BYTES = 60 * 1024 * 1024;

const sheet = document.querySelector(".sheet");
const form = document.getElementById("form");
const submit = document.getElementById("submit");
const statusLine = document.getElementById("status");
const watermarkOn = document.getElementById("watermark-on");
const watermarkLabel = document.getElementById("watermark-label");
const secondSlot = document.getElementById("image-slot-2");
const addImage = document.getElementById("add-image");

const MODES = {
  protect: { endpoint: "/api/protect", button: "Protect document", done: "Protected" },
  insert: { endpoint: "/api/insert", button: "Add images and protect", done: "Images added and protected" },
};

function drawSeal() {
  // A rosette of interlaced hypotrochoid bands.
  const seal = document.getElementById("seal");
  const bands = [
    { R: 96, r: 7, d: 16, turns: 7 },
    { R: 74, r: 9, d: 20, turns: 9 },
    { R: 46, r: 6, d: 15, turns: 6 },
  ];
  const paths = bands.map(({ R, r, d, turns }) => {
    const points = [];
    for (let i = 0; i <= 1400; i++) {
      const t = (i / 1400) * Math.PI * 2 * turns;
      const k = R - d - r;
      const x = k * Math.cos(t) + d * Math.cos(((k + r) / r) * t);
      const y = k * Math.sin(t) - d * Math.sin(((k + r) / r) * t);
      points.push(`${x.toFixed(2)} ${y.toFixed(2)}`);
    }
    return `<path d="M${points.join("L")}Z"/>`;
  });
  seal.innerHTML = paths.join("");
}

function setStatus(message, isError = false) {
  statusLine.textContent = message;
  statusLine.classList.toggle("is-error", isError);
}

function extensionOf(name) {
  return name.slice(name.lastIndexOf(".")).toLowerCase();
}

// A drop zone that remembers one chosen file of the allowed types.
function dropZone(id, extensions, wrongType) {
  const zone = document.getElementById(id);
  const input = zone.querySelector("input");
  const title = zone.querySelector("[data-title]");
  const hint = zone.querySelector("[data-hint]");
  const empty = { title: title.textContent, hint: hint.textContent };
  const state = { file: null };

  state.clear = () => {
    state.file = null;
    input.value = "";
    title.textContent = empty.title;
    hint.textContent = empty.hint;
  };

  function choose(file) {
    sheet.dataset.state = "";
    if (!extensions.includes(extensionOf(file.name))) {
      state.clear();
      return setStatus(wrongType, true);
    }
    state.file = file;
    title.textContent = file.name;
    hint.textContent = "Choose or drop another file to replace it.";
    setStatus("");
  }

  input.addEventListener("change", () => {
    if (input.files.length) choose(input.files[0]);
  });
  for (const type of ["dragenter", "dragover"]) {
    zone.addEventListener(type, (event) => {
      event.preventDefault();
      zone.classList.add("is-over");
    });
  }
  for (const type of ["dragleave", "drop"]) {
    zone.addEventListener(type, (event) => {
      event.preventDefault();
      zone.classList.remove("is-over");
    });
  }
  zone.addEventListener("drop", (event) => {
    if (event.dataTransfer.files.length) choose(event.dataTransfer.files[0]);
  });
  return state;
}

const IMAGE_TYPES = [".png", ".jpg", ".jpeg"];
const zones = {
  document: dropZone(
    "document-drop",
    [".pdf", ".doc", ".docx", ".ppt", ".pptx"],
    "Choose a PDF, Word (.doc, .docx) or PowerPoint (.ppt, .pptx) file.",
  ),
  pdf: dropZone("pdf-drop", [".pdf"], "Choose a PDF file. To protect a Word or PowerPoint file, use Protect a document."),
  image1: dropZone("image-drop-1", IMAGE_TYPES, "Choose a PNG or JPEG image."),
  image2: dropZone("image-drop-2", IMAGE_TYPES, "Choose a PNG or JPEG image."),
};

function currentMode() {
  return form.elements.mode.value;
}

// Returns the form to send, or an error message to show.
function buildUpload() {
  const body = new FormData();
  if (currentMode() === "protect") {
    if (!zones.document.file) return "Choose a document first.";
    body.append("file", zones.document.file);
  } else {
    if (!zones.pdf.file) return "Choose a PDF first.";
    body.append("file", zones.pdf.file);
    for (const number of secondSlot.hidden ? [1] : [1, 2]) {
      const image = zones[`image${number}`].file;
      if (!image) return number === 1 ? "Choose image 1." : "Choose image 2, or remove it.";
      const position = document.getElementById(`position-${number}`).value;
      body.append(`image${number}`, image);
      body.append(`position${number}`, position);
      if (position === "before") {
        const page = document.getElementById(`page-${number}`).value.trim();
        if (!/^[1-9]\d*$/.test(page)) return `Enter the page that image ${number} goes before.`;
        body.append(`page${number}`, page);
      }
    }
  }

  if (watermarkOn.checked) {
    if (!watermarkLabel.value.trim()) return "Enter the watermark text, or untick Add a watermark.";
    body.append("watermark", watermarkLabel.value);
  }
  body.append("quality", form.elements.quality.value);

  let total = 0;
  for (const value of body.values()) if (value instanceof File) total += value.size;
  if (total > MAX_BYTES) return "These files are over the 60 MB limit together.";
  return body;
}

function downloadName(response) {
  const header = response.headers.get("Content-Disposition") || "";
  const match = header.match(/filename\*=UTF-8''([^;]+)/);
  return match ? decodeURIComponent(match[1]) : "document-encreept.pdf";
}

async function errorMessage(response) {
  try {
    return (await response.json()).detail;
  } catch {
    return "The server stopped while working on this document, which usually means it's too large or complex. Save it as a PDF, upload that instead, and wait a minute before trying.";
  }
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const body = buildUpload();
  if (typeof body === "string") return setStatus(body, true);
  const mode = MODES[currentMode()];

  submit.disabled = true;
  sheet.dataset.state = "working";
  setStatus("Working on your document. Keep this page open: long files can take several minutes or more.");

  try {
    const response = await fetch(mode.endpoint, { method: "POST", body });
    if (!response.ok) {
      sheet.dataset.state = "";
      return setStatus(await errorMessage(response), true);
    }
    const url = URL.createObjectURL(await response.blob());
    const link = document.createElement("a");
    link.href = url;
    link.download = downloadName(response);
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    sheet.dataset.state = "done";
    setStatus(`${mode.done}. ${link.download} is in your downloads.`);
  } catch {
    sheet.dataset.state = "";
    setStatus("The connection dropped before the document was ready. Check your connection and try again.", true);
  } finally {
    submit.disabled = false;
  }
});

for (const radio of form.elements.mode) {
  radio.addEventListener("change", () => {
    const insert = currentMode() === "insert";
    document.getElementById("protect-fields").hidden = insert;
    document.getElementById("insert-fields").hidden = !insert;
    submit.textContent = MODES[currentMode()].button;
    sheet.dataset.state = "";
    setStatus("");
  });
}

for (const number of [1, 2]) {
  const position = document.getElementById(`position-${number}`);
  const page = document.getElementById(`page-${number}`);
  position.addEventListener("change", () => {
    page.hidden = position.value !== "before";
    if (!page.hidden) page.focus();
  });
}

addImage.addEventListener("click", () => {
  secondSlot.hidden = false;
  addImage.hidden = true;
  document.querySelector("#image-drop-2 input").focus();
});

document.getElementById("remove-image").addEventListener("click", () => {
  zones.image2.clear();
  document.getElementById("position-2").value = "last";
  document.getElementById("page-2").hidden = true;
  secondSlot.hidden = true;
  addImage.hidden = false;
  addImage.focus();
  setStatus("");
});

function updateWatermarkCount() {
  const left = watermarkLabel.maxLength - watermarkLabel.value.length;
  document.getElementById("watermark-count").textContent =
    `${left} of ${watermarkLabel.maxLength} characters left`;
}

watermarkOn.addEventListener("change", () => {
  document.getElementById("watermark-fields").hidden = !watermarkOn.checked;
  watermarkLabel.disabled = !watermarkOn.checked;
  if (watermarkOn.checked) watermarkLabel.focus();
  else setStatus("");
});

watermarkLabel.addEventListener("input", updateWatermarkCount);

drawSeal();
