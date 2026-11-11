
"use strict";

const $ = id => document.getElementById(id);
const MAX_FILE_SIZE = 30 * 1024 * 1024;

let questions = [];
let position = 0;
let score = 0;
let selected = -1;
let answered = false;
let ocrWorker = null;

const screens = [
  $("setupScreen"),
  $("quizScreen"),
  $("resultScreen")
];

function show(element) {
  screens.forEach(x => x.classList.add("hidden"));
  element.classList.remove("hidden");
  window.scrollTo(0, 0);
}

function status(message, error = false) {
  $("statusMessage").textContent = message;
  $("statusMessage").classList.toggle("error", error);
}

$("pptFile").addEventListener("change", e => {
  const file = e.target.files[0];
  if (!file) return;

  if (!/\.pptx$/i.test(file.name) ||
      file.size > MAX_FILE_SIZE) {
    status("Choose a .pptx file no larger than 30 MB.", true);
    e.target.value = "";
    $("fileName").textContent = "No file selected";
    return;
  }

  $("fileName").textContent =
    `✓ ${file.name} (${(file.size / 1048576).toFixed(1)} MB)`;

  status("Presentation ready.");
});

function xmlDoc(str) {
  const doc = new DOMParser().parseFromString(
    str, "application/xml"
  );

  if (doc.getElementsByTagName("parsererror").length) {
    throw Error("Invalid slide XML.");
  }

  return doc;
}

function slideText(doc) {
  const ns =
    "http://schemas.openxmlformats.org/drawingml/2006/main";

  return [...doc.getElementsByTagNameNS(ns, "p")]
    .map(p =>
      [...p.getElementsByTagNameNS(ns, "t")]
        .map(t => t.textContent)
        .join("")
        .trim()
    )
    .filter(Boolean)
    .join("\n");
}

async function imagesForSlide(zip, number, doc) {
  const file = zip.file(
    `ppt/slides/_rels/slide${number}.xml.rels`
  );

  if (!file) return [];

  const rels = xmlDoc(await file.async("string"));

  const byId = new Map(
    [...rels.getElementsByTagName("Relationship")]
      .map(r => [
        r.getAttribute("Id"),
        {
          target: r.getAttribute("Target"),
          type: r.getAttribute("Type") || ""
        }
      ])
  );

  const ns =
    "http://schemas.openxmlformats.org/drawingml/2006/main";

  const rns =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  const paths = [];

  for (const blip of doc.getElementsByTagNameNS(ns, "blip")) {
    const rel = byId.get(
      blip.getAttributeNS(rns, "embed")
    );

    if (!rel || !rel.type.endsWith("/image") ||
        !rel.target) {
      continue;
    }

    const url = new URL(
      rel.target,
      "https://ppt.local/ppt/slides/"
    );

    if (url.origin !== "https://ppt.local") continue;

    const path = decodeURIComponent(
      url.pathname.slice(1)
    );

    if (
      /\.(png|jpe?g|webp)$/i.test(path) &&
      zip.file(path) &&
      !paths.includes(path)
    ) {
      paths.push(path);
    }
  }

  return paths;
}

async function imageCanvas(bytes, filename) {
  const mime = filename.toLowerCase().endsWith(".png")
    ? "image/png"
    : filename.toLowerCase().endsWith(".webp")
      ? "image/webp"
      : "image/jpeg";

  const url = URL.createObjectURL(
    new Blob([bytes], { type: mime })
  );

  try {
    const img = new Image();
    img.src = url;
    await img.decode();

    const scale = Math.min(
      1,
      1900 / Math.max(img.width, img.height)
    );

    const canvas = document.createElement("canvas");

    canvas.width = Math.max(
      1, Math.round(img.width * scale)
    );

    canvas.height = Math.max(
      1, Math.round(img.height * scale)
    );

    const ctx = canvas.getContext("2d");

    if (!ctx) throw Error("Canvas not supported.");

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(
      0, 0, canvas.width, canvas.height
    );

    ctx.drawImage(
      img, 0, 0, canvas.width, canvas.height
    );

    return canvas;

  } finally {
    URL.revokeObjectURL(url);
  }
}

async function getOCR() {
  if (!window.Tesseract) {
    throw Error(
      "OCR library could not load. Check your connection."
    );
  }

  if (!ocrWorker) {
    status("Loading English OCR language data...");

    ocrWorker = await Tesseract.createWorker(
      "eng", 1
    );
  }

  return ocrWorker;
}

async function readPPT(file) {
  if (!window.JSZip) {
    throw Error(
      "PowerPoint ZIP reader did not load."
    );
  }

  const zip = await JSZip.loadAsync(file);
  const files = [];

  zip.forEach((path, entry) => {
    const m = /^ppt\/slides\/slide(\d+)\.xml$/.exec(path);

    if (m && !entry.dir) {
      files.push({
        number: Number(m[1]),
        path
      });
    }
  });

  files.sort((a, b) => a.number - b.number);

  if (!files.length) {
    throw Error("This is not a readable PPTX file.");
  }

  if (files.length > 100) {
    throw Error("Maximum 100 slides supported.");
  }

  const slides = [];

  for (let i = 0; i < files.length; i++) {
    const s = files[i];

    status(
      `Reading slide ${i + 1} of ${files.length}...`
    );

    const doc = xmlDoc(
      await zip.file(s.path).async("string")
    );

    let text = slideText(doc);

    const paths = await imagesForSlide(
      zip, s.number, doc
    );

    // OCR up to three images per slide.
    for (const path of paths.slice(0, 3)) {
      try {
        const bytes = await zip.file(path)
          .async("uint8array");

        const canvas = await imageCanvas(
          bytes, path
        );

        const worker = await getOCR();

        status(
          `Reading image text on slide ${s.number}...`
        );

        const recognized = (
          await worker.recognize(canvas)
        ).data.text.trim();

        if (recognized) {
          text += "\n" + recognized;
        }

      } catch (e) {
        console.warn(
          "OCR issue:", s.number, e
        );
      }
    }

    if (text.trim()) {
      slides.push({
        number: s.number,
        text: text.trim().slice(0, 6500)
      });
    }
  }

  if (!slides.length) {
    throw Error(
      "No legible text detected. Use clearer slide pictures."
    );
  }

  return slides;
}

// API: always read the raw response first.
// This gives a useful error even if Vercel returns HTML.
async function aiQuiz(slides, count, instructions) {
  let response;

  try {
    response = await fetch("/api/quiz", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        slides, count, instructions
      })
    });

  } catch {
    throw Error(
      "Cannot connect to the Vercel API."
    );
  }

  const raw = await response.text();

  let data;

  try {
    data = JSON.parse(raw);
  } catch {
    throw Error(
      `API returned HTTP ${response.status}, not JSON. ` +
      "Check api/quiz.js and Vercel logs."
    );
  }

  if (!response.ok) {
    throw Error(
      data.error || `API error ${response.status}`
    );
  }

  if (!Array.isArray(data.questions) ||
      !data.questions.length) {
    throw Error("No AI questions received.");
  }

  return data.questions;
}

// Simple rule-based offline fallback.
// This is not a generative language model.
function localQuiz(slides, count) {
  const pool = [];

  for (const s of slides) {
    for (let line of s.text.split(
      /\n|(?<=[.!?])\s+/
    )) {
      line = line.replace(/\s+/g, " ").trim();

      if (
        line.length >= 35 &&
        line.length <= 240 &&
        /[a-z]{4}/i.test(line)
      ) {
        pool.push({
          line,
          slide: s.number
        });
      }
    }
  }

  if (!pool.length) {
    throw Error(
      "Not enough legible text for local quiz questions."
    );
  }

  const terms = [
    ...new Set(
      pool.flatMap(x =>
        x.line.match(/\b[A-Za-z]{5,}\b/g) || []
      )
    )
  ];

  const result = [];

  for (let i = 0; i < count; i++) {
    const s = pool[i % pool.length];

    const word = (
      s.line.match(/\b[A-Za-z]{5,}\b/g) || []
    ).find(w =>
      !/^(which|their|there|these|those|about|would|could|should|after|before|where|while|other|using|through)$/i.test(w)
    );

    const evidence = s.line;

    const type =
      i % 4 === 0 ? "mcq" :
      i % 4 === 1 ? "one_word" :
      i % 4 === 2 ? "short" : "long";

    const blank = word
      ? s.line.replace(
          new RegExp(`\\b${word}\\b`, "i"),
          "____"
        )
      : s.line;

    const q = {
      type,
      question:
        "Complete this statement from the presentation: " +
        blank,
      answer: word || s.line,
      explanation:
        "This word or statement appears in the slide text.",
      evidence,
      slideNumber: s.slide
    };

    if (type === "mcq" && word) {
      const wrong = terms.filter(x =>
        x.toLowerCase() !== word.toLowerCase() &&
        !s.line.toLowerCase().includes(
          x.toLowerCase()
        )
      ).slice(0, 3);

      if (wrong.length === 3) {
        q.options = [word, ...wrong];

        q.options.push(
          ...q.options.splice(0, i % 4)
        );

        q.correctIndex = q.options.indexOf(word);
      } else {
        q.type = "one_word";
      }

    } else if (
      type === "short" ||
      type === "long" ||
      !word
    ) {
      q.type = type === "long" ? "long" : "short";

      q.question =
        `What does slide ${s.slide} state ` +
        "about this topic? Use the slide's own wording.";

      q.answer = s.line;
    }

    result.push(q);
  }

  return result;
}

// Start quiz generation.
$("generateBtn").addEventListener(
  "click",
  async () => {
    const file = $("pptFile").files[0];

    if (!file) {
      status(
        "Select a PowerPoint first.", true
      );
      return;
    }

    const button = $("generateBtn");
    button.disabled = true;
    button.textContent = "Reading slides...";

    try {
      const slides = await readPPT(file);

      const count = Number(
        $("questionCount").value
      );

      const mode = $("quizMode").value;

      if (mode === "local") {
        questions = localQuiz(slides, count);
      } else {
        status("Generating questions...");

        try {
          questions = await aiQuiz(
            slides,
            count,
            $("instructions").value.trim()
          );

        } catch (error) {
          if (mode === "ai") throw error;

          console.warn(
            "AI fallback:", error.message
          );

          questions = localQuiz(
            slides, count
          );
        }
      }

      position = 0;
      score = 0;

      renderQuestion();
      show($("quizScreen"));

    } catch (e) {
      status(
        e.message || "Could not create quiz.",
        true
      );
      console.error(e);

    } finally {
      button.disabled = false;
      button.textContent = "✦ Generate quiz";
    }
  }
);

function addLine(parent, value, tag = "p") {
  const e = document.createElement(tag);

  e.textContent = value;
  parent.appendChild(e);

  return e;
}

function renderQuestion() {
  const q = questions[position];

  selected = -1;
  answered = false;

  $("studentLabel").textContent =
    ($("studentName").value.trim() || "Student") +
    "'s quiz";

  $("questionNumber").textContent =
    `Question ${position + 1} of ${questions.length}`;

  $("progressBar").style.width =
    `${100 * (position + 1) / questions.length}%`;

  $("questionType").textContent = ({
    mcq: "Choose the correct answer",
    one_word: "One-word answer",
    short: "Short answer",
    long: "Answer the following"
  })[q.type] || "Question";

  $("slideNumber").textContent =
    "Slide " + q.slideNumber;

  $("questionText").textContent =
    q.question;

  $("feedback").className =
    "feedback hidden";

  $("feedback").replaceChildren();

  $("nextBtn").classList.add("hidden");
  $("submitBtn").classList.remove("hidden");

  const area = $("answerArea");
  area.replaceChildren();

  if (
    q.type === "mcq" &&
    Array.isArray(q.options)
  ) {
    q.options.forEach((option, i) => {
      const b = document.createElement("button");

      b.type = "button";
      b.className = "option";

      b.textContent =
        `${String.fromCharCode(65 + i)}. ${option}`;

      b.addEventListener("click", () => {
        if (answered) return;

        selected = i;

        area.querySelectorAll(".option")
          .forEach(x =>
            x.classList.remove("selected")
          );

        b.classList.add("selected");
      });

      area.appendChild(b);
    });

  } else {
    const field = document.createElement("textarea");

    field.id = "writtenAnswer";
    field.className = "answer-input";
    field.rows = q.type === "long" ? 7 : 3;
    field.placeholder = "Type your answer...";

    area.appendChild(field);
  }
}

function showFeedback(
  q, heading, correct = null
) {
  const box = $("feedback");

  box.replaceChildren();

  box.className =
    "feedback" +
    (correct === null
      ? ""
      : correct ? " success" : " error");

  addLine(box, heading, "strong");
  addLine(box, "Answer: " + q.answer);

  addLine(
    box,
    "Explanation: " + q.explanation
  );

  addLine(
    box,
    `Slide ${q.slideNumber} evidence: "${q.evidence}"`
  );
}

$("submitBtn").addEventListener(
  "click",
  () => {
    if (answered) return;

    const q = questions[position];

    if (
      q.type === "mcq" &&
      Array.isArray(q.options)
    ) {
      if (selected < 0) {
        alert("Choose an option first.");
        return;
      }

      answered = true;

      const correct =
        selected === q.correctIndex;

      if (correct) score++;

      $("answerArea")
        .querySelectorAll(".option")
        .forEach((b, i) => {
          b.disabled = true;

          if (i === q.correctIndex) {
            b.classList.add("correct");
          }

          if (i === selected && !correct) {
            b.classList.add("wrong");
          }
        });

      showFeedback(
        q,
        correct ? "✓ Correct" : "Not quite",
        correct
      );

      $("nextBtn").classList.remove("hidden");

    } else {
      const field = $("writtenAnswer");

      if (!field.value.trim()) {
        alert("Type your answer first.");
        return;
      }

      answered = true;
      field.disabled = true;

      showFeedback(q, "Compare your answer");

      const box = $("feedback");

      addLine(
        box,
        "Did your answer cover the key points?"
      );

      for (
        const [label, gotIt] of [
          ["✓ Yes, I got it right", true],
          ["I need more practice", false]
        ]
      ) {
        const b = document.createElement("button");

        b.type = "button";

        b.className = gotIt
          ? "primary"
          : "secondary";

        b.textContent = label;

        b.onclick = () => {
          if (gotIt) score++;

          box.querySelectorAll("button")
            .forEach(x => x.remove());

          $("nextBtn").classList.remove(
            "hidden"
          );
        };

        box.appendChild(b);
      }
    }

    $("submitBtn").classList.add("hidden");
  }
);

$("nextBtn").addEventListener(
  "click",
  () => {
    position++;

    if (position < questions.length) {
      renderQuestion();
      window.scrollTo(0, 0);
      return;
    }

    $("resultName").textContent =
      "Well done, " +
      ($("studentName").value.trim() || "Student") +
      "!";

    $("finalScore").textContent =
      `${score} / ${questions.length}`;

    $("resultMessage").textContent =
      "Written answers are self-assessed. Keep practising!";

    show($("resultScreen"));
  }
);

$("restartBtn").addEventListener(
  "click",
  () => {
    position = 0;
    score = 0;

    renderQuestion();
    show($("quizScreen"));
  }
);

$("newQuizBtn").addEventListener(
  "click",
  () => {
    questions = [];
    position = 0;
    score = 0;

    show($("setupScreen"));
  }
);

window.addEventListener(
  "beforeunload",
  () => {
    if (ocrWorker) {
      ocrWorker.terminate();
    }
  }
);
