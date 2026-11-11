
"use strict";

// SlideSpark Fast Browser AI
// Requires JSZip, Tesseract.js and WebGPU.
// No OpenAI API key.

const $ = id => document.getElementById(id);
const MAX_SIZE = 30 * 1024 * 1024;

const MODEL_IDS = [
  "Llama-3.2-1B-Instruct-q4f32_1-MLC",
  "Llama-3.2-1B-Instruct-q4f16_1-MLC"
];

let engine = null;
let engineModel = "";
let ocrWorker = null;
let questions = [];
let current = 0;
let score = 0;
let selected = -1;
let answered = false;
let working = false;
let lastError = "";

let cachedFile = null;
let cachedMode = "";
let cachedSlides = null;

const screens = [
  $("setupScreen"),
  $("quizScreen"),
  $("resultScreen")
];

function status(message, error = false) {
  $("statusMessage").textContent = message;
  $("statusMessage").classList.toggle("error", error);
}

function progress(value) {
  $("loadingBar").style.width =
    `${Math.max(0, Math.min(100, value))}%`;
}

function show(screen) {
  screens.forEach(s => s.classList.add("hidden"));
  screen.classList.remove("hidden");
  window.scrollTo(0, 0);
}

function diagnostics(message) {
  lastError = message;
  $("diagnosticText").textContent = message;
  $("diagnosticPanel").classList.remove("hidden");
}

function clearDiagnostics() {
  lastError = "";
  $("diagnosticPanel").classList.add("hidden");
}

$("copyErrorBtn").onclick = () => {
  if (navigator.clipboard) {
    navigator.clipboard.writeText(lastError)
      .then(() => status("Error details copied."))
      .catch(() => status("Copy the details manually."));
  } else {
    status("Copy the details manually.");
  }
};

function normalize(value) {
  return String(value || "")
    .normalize("NFKC")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function addText(parent, value, tag = "p") {
  const el = document.createElement(tag);
  el.textContent = String(value ?? "");
  parent.appendChild(el);
  return el;
}

// ------------------------------------
// FILE UPLOAD
// ------------------------------------

$("pptFile").onchange = event => {
  const file = event.target.files[0];

  cachedFile = null;
  cachedSlides = null;

  if (!file) return;

  if (
    !/\.pptx$/i.test(file.name) ||
    file.size > MAX_SIZE
  ) {
    event.target.value = "";
    $("fileName").textContent = "No file selected";
    status("Select a PPTX up to 30 MB.", true);
    return;
  }

  $("fileName").textContent =
    `✓ ${file.name} (${(file.size / 1048576).toFixed(1)} MB)`;

  clearDiagnostics();
  progress(0);
  status("PowerPoint ready.");
};

// ------------------------------------
// XML HELPERS
// ------------------------------------

function parseXML(xml) {
  const doc = new DOMParser().parseFromString(
    xml,
    "application/xml"
  );

  if (doc.getElementsByTagName("parsererror").length) {
    throw Error("PowerPoint contains invalid slide XML.");
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

// ------------------------------------
// FIND SLIDE IMAGES
// ------------------------------------

async function slideImages(zip, number, doc) {
  const file = zip.file(
    `ppt/slides/_rels/slide${number}.xml.rels`
  );

  if (!file) return [];

  const relDoc = parseXML(
    await file.async("string")
  );

  const relations = new Map();

  for (
    const rel of relDoc.getElementsByTagName(
      "Relationship"
    )
  ) {
    relations.set(rel.getAttribute("Id"), {
      target: rel.getAttribute("Target"),
      type: rel.getAttribute("Type") || ""
    });
  }

  const aNS =
    "http://schemas.openxmlformats.org/drawingml/2006/main";

  const rNS =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  const images = [];

  for (
    const blip of doc.getElementsByTagNameNS(
      aNS,
      "blip"
    )
  ) {
    const rel = relations.get(
      blip.getAttributeNS(rNS, "embed")
    );

    if (
      !rel ||
      !rel.target ||
      !rel.type.endsWith("/image")
    ) continue;

    const url = new URL(
      rel.target,
      "https://ppt.local/ppt/slides/"
    );

    if (url.origin !== "https://ppt.local") continue;

    const path = decodeURIComponent(
      url.pathname.slice(1)
    );

    if (
      /\.(png|jpg|jpeg|webp)$/i.test(path) &&
      zip.file(path) &&
      !images.includes(path)
    ) {
      images.push(path);
    }
  }

  return images;
}

// ------------------------------------
// IMAGE OCR
// ------------------------------------

async function toCanvas(bytes, path) {
  const mime = /\.png$/i.test(path)
    ? "image/png"
    : /\.webp$/i.test(path)
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
      2,
      2000 / Math.max(img.width, img.height)
    );

    const canvas = document.createElement("canvas");

    canvas.width = Math.max(
      1,
      Math.round(img.width * scale)
    );

    canvas.height = Math.max(
      1,
      Math.round(img.height * scale)
    );

    const ctx = canvas.getContext("2d");

    if (!ctx) throw Error("Canvas unavailable.");

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

    return canvas;

  } finally {
    URL.revokeObjectURL(url);
  }
}

async function getOCR() {
  if (!window.Tesseract) {
    throw Error("OCR library did not load.");
  }

  if (!ocrWorker) {
    status("Loading OCR language data...");

    ocrWorker = await window.Tesseract.createWorker(
      "eng",
      1
    );
  }

  return ocrWorker;
}

// ------------------------------------
// FAST POWERPOINT READING
// ------------------------------------

async function readPPT(file, mode) {
  if (
    cachedFile === file &&
    cachedMode === mode &&
    cachedSlides
  ) {
    status("Using previously extracted slide text.");
    progress(40);
    return cachedSlides;
  }

  if (!window.JSZip) {
    throw Error("JSZip could not load.");
  }

  const zip = await window.JSZip.loadAsync(file);
  const entries = [];

  zip.forEach((path, entry) => {
    const match =
      /^ppt\/slides\/slide(\d+)\.xml$/.exec(path);

    if (match && !entry.dir) {
      entries.push({
        number: Number(match[1]),
        path
      });
    }
  });

  entries.sort((a, b) => a.number - b.number);

  if (!entries.length) {
    throw Error("No PPTX slides found.");
  }

  if (entries.length > 100) {
    throw Error("Maximum 100 slides supported.");
  }

  const imageLimit =
    mode === "text" ? 0 :
    mode === "thorough" ? 5 : 2;

  const ocrCache = new Map();
  const slides = [];

  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];

    status(
      `Reading slide ${i + 1}/${entries.length}...`
    );

    progress(
      5 + 35 * (i + 1) / entries.length
    );

    const doc = parseXML(
      await zip.file(entry.path).async("string")
    );

    const pieces = [];
    const nativeText = slideText(doc).trim();

    if (nativeText) pieces.push(nativeText);

    if (imageLimit > 0) {
      const paths = await slideImages(
        zip,
        entry.number,
        doc
      );

      for (const path of paths.slice(0, imageLimit)) {
        try {
          let recognized = ocrCache.get(path);

          if (recognized === undefined) {
            status(
              `OCR: slide ${entry.number}...`
            );

            const bytes = await zip.file(path)
              .async("uint8array");

            const canvas = await toCanvas(
              bytes,
              path
            );

            const worker = await getOCR();

            const data = (
              await worker.recognize(canvas)
            ).data;

            recognized =
              data.text?.trim() || "";

            if (
              Number(data.confidence) < 45 ||
              recognized.length < 25
            ) {
              recognized = "";
            }

            ocrCache.set(path, recognized);
          }

          if (
            recognized &&
            !normalize(nativeText).includes(
              normalize(recognized)
            )
          ) {
            pieces.push(recognized);
          }

        } catch (error) {
          console.warn(
            `OCR issue, slide ${entry.number}`,
            error
          );
        }
      }
    }

    const combined = pieces
      .join("\n")
      .replace(/[ \t]+/g, " ")
      .trim();

    if (combined.length >= 35) {
      slides.push({
        number: entry.number,
        text: combined.slice(0, 4200)
      });
    }
  }

  if (!slides.length) {
    throw Error(
      "No readable slide text. " +
      "Try Thorough mode or clearer slide images."
    );
  }

  cachedFile = file;
  cachedMode = mode;
  cachedSlides = slides;

  return slides;
}

// ------------------------------------
// LOAD AND REUSE BROWSER AI
// ------------------------------------

async function loadAI() {
  if (engine) return engine;

  if (!window.isSecureContext) {
    throw Error(
      "WebGPU requires HTTPS."
    );
  }

  if (!navigator.gpu) {
    throw Error(
      "WebGPU is unavailable. " +
      "Use updated Chrome/Edge with " +
      "hardware acceleration enabled."
    );
  }

  const adapter = await navigator.gpu.requestAdapter();

  if (!adapter) {
    throw Error(
      "No usable WebGPU adapter. " +
      "Check chrome://gpu and your graphics drivers."
    );
  }

  status("Loading free browser AI library...");

  let webllm;

  try {
    webllm = await import(
      "https://esm.run/@mlc-ai/web-llm@0.2.88"
    );
  } catch (error) {
    throw Error(
      "WebLLM library failed: " + error.message
    );
  }

  const available = new Set(
    webllm.prebuiltAppConfig.model_list.map(
      m => m.model_id
    )
  );

  const f16 = adapter.features.has("shader-f16");

  const choices = MODEL_IDS
    .filter(id => available.has(id))
    .filter(
      id => f16 || !id.includes("q4f16")
    );

  if (!choices.length) {
    throw Error(
      "No compatible model found in WebLLM catalogue."
    );
  }

  const errors = [];

  for (const model of choices) {
    try {
      status(`Loading AI model: ${model}`);

      const result = await webllm.CreateMLCEngine(
        model,
        {
          initProgressCallback: state => {
            const percent =
              Number.isFinite(state.progress)
                ? Math.round(state.progress * 100)
                : 0;

            progress(40 + percent * 0.3);

            status(
              `AI model loading: ${percent}%\n` +
              (state.text || model)
            );
          }
        }
      );

      engine = result;
      engineModel = model;

      status(`AI loaded: ${model}`);
      return engine;

    } catch (error) {
      errors.push(
        `${model}: ${error.message || error}`
      );
      console.warn(
        "WebLLM load issue:",
        error
      );
    }
  }

  throw Error(
    "AI models failed to load:\n" +
    errors.join("\n")
  );
}

// ------------------------------------
// QUESTION DISTRIBUTION
// ------------------------------------

function questionTypes(count) {
  const pattern = [
    "mcq", "mcq", "mcq",
    "one_word", "one_word",
    "short", "short",
    "long", "long", "long"
  ];

  return Array.from(
    { length: count },
    (_, i) => pattern[i % pattern.length]
  );
}

// ------------------------------------
// READ MODEL JSON
// ------------------------------------

function parseAIJSON(text) {
  const raw = String(text || "")
    .replace(/<think>[\s\S]*?<\/think>/g, "")
    .replace(/```json|```/gi, "")
    .trim();

  try {
    return JSON.parse(raw);
  } catch {}

  const first = raw.indexOf("{");
  const last = raw.lastIndexOf("}");

  if (first < 0 || last <= first) {
    throw Error("AI did not return valid JSON.");
  }

  return JSON.parse(
    raw.slice(first, last + 1)
  );
}

// ------------------------------------
// VALIDATE MODEL QUESTIONS
// ------------------------------------

function validateQuestion(q, slide, type, seen) {
  if (!q || typeof q !== "object") {
    throw Error("Invalid question object.");
  }

  const question = String(
    q.question || ""
  ).trim();

  const answer = String(
    q.answer || ""
  ).trim();

  const explanation = String(
    q.explanation || ""
  ).trim();

  const evidence = String(
    q.evidence || ""
  ).trim();

  if (
    question.length < 12 ||
    !answer ||
    explanation.length < 6 ||
    evidence.length < 8
  ) {
    throw Error("Incomplete AI question.");
  }

  if (
    /____|fill in the blank|what does slide \d+|complete the statement/i
      .test(question)
  ) {
    throw Error("Unhelpful question rejected.");
  }

  if (
    !normalize(slide.text).includes(
      normalize(evidence)
    )
  ) {
    throw Error(
      "Question evidence not found in slide."
    );
  }

  if (seen.has(normalize(question))) {
    throw Error("Duplicate question rejected.");
  }

  const result = {
    type,
    slideNumber: slide.number,
    question,
    answer,
    explanation,
    evidence
  };

  if (type === "mcq") {
    const options = Array.isArray(q.options)
      ? q.options.map(x => String(x).trim())
      : [];

    const index = Number(q.correctIndex);

    if (
      options.length !== 4 ||
      options.some(x => !x) ||
      new Set(options.map(normalize)).size !== 4 ||
      !Number.isInteger(index) ||
      index < 0 ||
      index > 3 ||
      normalize(options[index]) !==
        normalize(answer)
    ) {
      throw Error("Invalid multiple-choice question.");
    }

    result.options = options;
    result.correctIndex = index;
  }

  return result;
}

// ------------------------------------
// FAST BATCH QUESTION GENERATION
// ------------------------------------

async function generateQuiz(slides, count, instructions) {
  const ai = await loadAI();

  const result = [];
  const seen = new Set();
  const types = questionTypes(count);

  const usable = slides.filter(
    s => s.text.length >= 70
  );

  if (!usable.length) {
    throw Error(
      "Not enough readable text for reliable questions."
    );
  }

  // Try two questions per request.
  // On failure retry with individual questions.
  for (let i = 0; i < count;) {
    const requested = types.slice(i, i + 2);
    let added = [];

    for (let attempt = 0; attempt < 3; attempt++) {
      const slide = usable[
        (Math.floor(i / 2) + attempt) % usable.length
      ];

      status(
        `Generating questions ${i + 1}–` +
        `${i + requested.length} of ${count}...`
      );

      progress(
        70 + 28 * i / count
      );

      const batchTypes =
        attempt === 2
          ? [requested[0]]
          : requested;

      const prompt = `
You are generating school quiz questions.

Create exactly ${batchTypes.length}
meaningful questions.

Types in exact order:
${batchTypes.join(", ")}

ONLY use these source facts:
SLIDE ${slide.number}:
${slide.text.slice(0, 1650)}

Teacher preferences:
${instructions.slice(0, 300)}

Rules:
- Ask specific educational questions.
- Do not invent facts or missing information.
- Do not ask "What does slide say?"
- No fill-in-the-blanks.
- Each question needs an exact short
  quote from the slide in evidence.
- MCQs have four distinct options.
- correctIndex is 0, 1, 2, or 3.
- MCQ answer exactly equals its correct option.
- Other types have options [] and
  correctIndex -1.
- Output JSON only.

Format:
{
  "questions": [
    {
      "question": "specific question",
      "answer": "correct answer",
      "explanation": "brief explanation",
      "evidence": "verbatim slide quote",
      "options": [],
      "correctIndex": -1
    }
  ]
}
`;

      try {
        const response =
          await ai.chat.completions.create({
            messages: [
              {
                role: "system",
                content:
                  "You write accurate educational " +
                  "questions using only supplied text. " +
                  "Return a JSON object, no markdown."
              },
              {
                role: "user",
                content: prompt
              }
            ],
            temperature: 0.15,
            max_tokens: 750
          });

        const raw =
          response.choices?.[0]
            ?.message?.content || "";

        const parsed = parseAIJSON(raw);

        if (
          !Array.isArray(parsed.questions) ||
          parsed.questions.length !==
            batchTypes.length
        ) {
          throw Error(
            "AI returned an incorrect question count."
          );
        }

        const proposed = [];
        const temporarySeen = new Set(seen);

        for (let j = 0; j < batchTypes.length; j++) {
          const cleaned = validateQuestion(
            parsed.questions[j],
            slide,
            batchTypes[j],
            temporarySeen
          );

          proposed.push(cleaned);

          temporarySeen.add(
            normalize(cleaned.question)
          );
        }

        added = proposed;
        break;

      } catch (error) {
        console.warn(
          `Generation attempt ${attempt + 1}:`,
          error
        );
      }
    }

    if (!added.length) {
      throw Error(
        `Could not create reliable question ${i + 1}. ` +
        "Try a smaller quiz or Thorough reading mode."
      );
    }

    for (const q of added) {
      result.push(q);
      seen.add(normalize(q.question));
    }

    i += added.length;
  }

  progress(100);
  return result;
}

// ------------------------------------
// START GENERATION
// ------------------------------------

$("generateBtn").onclick = async () => {
  if (working) return;

  const file = $("pptFile").files[0];

  if (!file) {
    status("Upload a PowerPoint first.", true);
    return;
  }

  working = true;
  clearDiagnostics();

  const button = $("generateBtn");
  button.disabled = true;
  button.textContent = "Generating quiz...";

  try {
    progress(2);

    const mode = $("readMode").value;

    const slides = await readPPT(file, mode);

    questions = await generateQuiz(
      slides,
      Number($("questionCount").value),
      $("instructions").value.trim()
    );

    current = 0;
    score = 0;

    renderQuestion();
    show($("quizScreen"));

  } catch (error) {
    console.error("SlideSpark error:", error);

    status(
      "Quiz generation failed. " +
      "See the technical details below.",
      true
    );

    diagnostics(
      `Browser: ${navigator.userAgent}\n` +
      `WebGPU: ${Boolean(navigator.gpu)}\n` +
      `AI model: ${engineModel || "Not loaded"}\n` +
      `Error: ${error.message || error}`
    );

  } finally {
    working = false;
    button.disabled = false;
    button.textContent = "✦ Generate quiz";
  }
};

// ------------------------------------
// DISPLAY QUIZ
// ------------------------------------

function renderQuestion() {
  const q = questions[current];

  selected = -1;
  answered = false;

  $("studentLabel").textContent =
    ($("studentName").value.trim() || "Student") +
    "'s quiz";

  $("questionNumber").textContent =
    `Question ${current + 1} of ${questions.length}`;

  $("progressBar").style.width =
    `${100 * (current + 1) / questions.length}%`;

  $("questionType").textContent = ({
    mcq: "Multiple choice",
    one_word: "One-word answer",
    short: "Short answer",
    long: "Detailed answer"
  })[q.type] || "Question";

  $("slideNumber").textContent =
    `Slide ${q.slideNumber}`;

  $("questionText").textContent = q.question;

  $("feedback").className = "feedback hidden";
  $("feedback").replaceChildren();

  $("nextBtn").classList.add("hidden");
  $("submitBtn").classList.remove("hidden");

  const area = $("answerArea");
  area.replaceChildren();

  if (q.type === "mcq") {
    q.options.forEach((option, index) => {
      const button = document.createElement("button");

      button.className = "option";
      button.type = "button";

      button.textContent =
        `${String.fromCharCode(65 + index)}. ${option}`;

      button.onclick = () => {
        if (answered) return;

        selected = index;

        area.querySelectorAll(".option")
          .forEach(x =>
            x.classList.remove("selected")
          );

        button.classList.add("selected");
      };

      area.appendChild(button);
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

// ------------------------------------
// ANSWER FEEDBACK
// ------------------------------------

function showFeedback(q, title, correct = null) {
  const box = $("feedback");
  box.replaceChildren();

  box.className =
    "feedback" +
    (correct === null
      ? ""
      : correct ? " success" : " error");

  addText(box, title, "strong");
  addText(box, "Correct answer: " + q.answer);
  addText(box, "Explanation: " + q.explanation);
  addText(
    box,
    `Evidence from slide ${q.slideNumber}: ${q.evidence}`
  );
}

// ------------------------------------
// CHECK ANSWER
// ------------------------------------

$("submitBtn").onclick = () => {
  if (answered) return;

  const q = questions[current];

  if (q.type === "mcq") {
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
      .forEach((button, i) => {
        button.disabled = true;

        if (i === q.correctIndex) {
          button.classList.add("correct");
        }

        if (i === selected && !correct) {
          button.classList.add("wrong");
        }
      });

    showFeedback(
      q,
      correct ? "✓ Correct!" : "Incorrect answer",
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

    addText(
      $("feedback"),
      "Did your answer include the important points?"
    );

    [
      ["✓ Yes, I got it right", true],
      ["I need more practice", false]
    ].forEach(([label, correct]) => {
      const button = document.createElement("button");

      button.className =
        correct ? "primary" : "secondary";

      button.type = "button";
      button.textContent = label;

      button.onclick = () => {
        if (correct) score++;

        $("feedback").querySelectorAll("button")
          .forEach(x => x.remove());

        $("nextBtn").classList.remove("hidden");
      };

      $("feedback").appendChild(button);
    });
  }

  $("submitBtn").classList.add("hidden");
};

// ------------------------------------
// NEXT QUESTION
// ------------------------------------

$("nextBtn").onclick = () => {
  current++;

  if (current < questions.length) {
    renderQuestion();
    window.scrollTo(0, 0);
    return;
  }

  const name =
    $("studentName").value.trim() || "Student";

  $("resultName").textContent =
    `Well done, ${name}!`;

  $("finalScore").textContent =
    `${score} / ${questions.length}`;

  $("resultMessage").textContent =
    "Written answers are self-assessed.";

  show($("resultScreen"));
};

$("restartBtn").onclick = () => {
  current = 0;
  score = 0;
  renderQuestion();
  show($("quizScreen"));
};

$("newQuizBtn").onclick = () => {
  questions = [];
  current = 0;
  score = 0;
  progress(0);
  show($("setupScreen"));
};

window.addEventListener("beforeunload", () => {
  if (ocrWorker) {
    ocrWorker.terminate();
  }
});
