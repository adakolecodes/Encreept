const MAX_BYTES = 25 * 1024 * 1024;
const EXTENSIONS = [".pdf", ".doc", ".docx", ".ppt", ".pptx"];

const sheet = document.querySelector(".sheet");
const form = document.getElementById("form");
const drop = document.getElementById("drop");
const input = document.getElementById("file");
const title = document.getElementById("drop-title");
const submit = document.getElementById("submit");
const statusLine = document.getElementById("status");

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

function chooseFile(file) {
  sheet.dataset.state = "";
  const extension = file.name.slice(file.name.lastIndexOf(".")).toLowerCase();
  if (!EXTENSIONS.includes(extension)) {
    submit.disabled = true;
    return setStatus("Choose a PDF, Word (.doc, .docx) or PowerPoint (.ppt, .pptx) file.", true);
  }
  if (file.size > MAX_BYTES) {
    submit.disabled = true;
    return setStatus("This file is over the 25 MB limit.", true);
  }
  title.textContent = file.name;
  document.getElementById("drop-hint").textContent = "Choose or drop another file to replace it.";
  submit.disabled = false;
  setStatus("");
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
    return "The server couldn't protect this document. Try again in a minute.";
  }
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!input.files.length) return;

  submit.disabled = true;
  sheet.dataset.state = "working";
  setStatus("Protecting your document. Long files and first use after a quiet spell can take a few minutes.");

  try {
    const response = await fetch("/api/protect", { method: "POST", body: new FormData(form) });
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
    setStatus(`Protected. ${link.download} is in your downloads.`);
  } catch {
    sheet.dataset.state = "";
    setStatus("The connection dropped before the document was ready. Check your connection and try again.", true);
  } finally {
    submit.disabled = false;
  }
});

input.addEventListener("change", () => {
  if (input.files.length) chooseFile(input.files[0]);
});

for (const type of ["dragenter", "dragover"]) {
  drop.addEventListener(type, (event) => {
    event.preventDefault();
    drop.classList.add("is-over");
  });
}
for (const type of ["dragleave", "drop"]) {
  drop.addEventListener(type, (event) => {
    event.preventDefault();
    drop.classList.remove("is-over");
  });
}
drop.addEventListener("drop", (event) => {
  if (!event.dataTransfer.files.length) return;
  input.files = event.dataTransfer.files;
  chooseFile(input.files[0]);
});

drawSeal();
