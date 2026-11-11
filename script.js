
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
  screens.forEach(screen => screen.classList.add("hidden"));
  element.classList.remove("hidden");
  window.scrollTo(0, 0);
}

function status(message, error = false) {
  $("statusMessage").textContent = message;
  $("statusMessage").classList.toggle("error", error);
}

// AI is the default. Local/automatic question generation is disabled
// because the old local algorithm produced unreliable questions.
const quizMode = $("quizMode");

if (quizMode) {
  quizMode.value = "ai";

  [...quizMode.options].forEach(option => {
    if (option.value === "ai") {
      option.textContent = "AI only (recommended)";
    } else {
      option.disabled = true;
      option.textContent += " — disabled";
    }
  });
}

const instructionsField = $("instructions");

if (instructionsField) {
  instructionsField.value =
    "Create meaningful educational questions using ONLY " +
    "the content of the uploaded PowerPoint. " +
    "For 10 questions, use 3 MCQs, 2 one-word answers, " +
    "2 short answers, and 3 detailed-answer questions. " +
    "Ask clear, specific, concept-based questions. " +
    "Do not create random fill-in-the-blank questions. " +
    "Do not ask vague questions such as 'What does slide 5 say?' " +
    "Do not invent facts or answers. " +
    "Include four plausible options for MCQs, " +
    "the correct answer, a useful explanation, " +
    "an exact evidence quote, and the source slide number. " +
    "If OCR text is unclear, avoid questions about it. " +
    "Make the quiz suitable for a school student.";
}

// --------------------------------------------------
// FILE UPLOAD
// --------------------------------------------------

$("pptFile").addEventListener("change", event => {
  const file = event.target.files[0];

  if (!file) return;

  if (!/\.pptx$/i.test(file.name)) {
    status("Please upload a .pptx PowerPoint file.", true);
    event.target.value = "";
    return;
  }

  if (file.size > MAX_FILE_SIZE) {
    status("Maximum supported file size is 30 MB.", true);
    event.target.value = "";
    return;
  }

  $("fileName").textContent =
    "✓ " + file.name +
    " (" + (file.size / 1048576).toFixed(1) + " MB)";

  status("PowerPoint uploaded successfully.");
});

// --------------------------------------------------
// READ POWERPOINT XML
// --------------------------------------------------

function xmlDoc(content) {
  const doc = new DOMParser().parseFromString(
    content,
    "application/xml"
  );

  if (doc.getElementsByTagName("parsererror").length) {
    throw new Error("Could not read PowerPoint slide XML.");
  }

  return doc;
}

function slideText(doc) {
  const namespace =
    "http://schemas.openxmlformats.org/drawingml/2006/main";

  return [...doc.getElementsByTagNameNS(namespace, "p")]
    .map(paragraph => {
      return [...paragraph.getElementsByTagNameNS(
        namespace,
        "t"
      )]
        .map(node => node.textContent)
        .join("")
        .trim();
    })
    .filter(Boolean)
    .join("\n");
}

// --------------------------------------------------
// FIND IMAGES IN EACH SLIDE
// --------------------------------------------------

async function imagesForSlide(zip, slideNumber, doc) {
  const relFile = zip.file(
    "ppt/slides/_rels/slide" + slideNumber + ".xml.rels"
  );

  if (!relFile) return [];

  const rels = xmlDoc(await relFile.async("string"));

  const relationships = new Map(
    [...rels.getElementsByTagName("Relationship")]
      .map(node => [
        node.getAttribute("Id"),
        {
          target: node.getAttribute("Target"),
          type: node.getAttribute("Type") || ""
        }
      ])
  );

  const drawingNS =
    "http://schemas.openxmlformats.org/drawingml/2006/main";

  const relationshipNS =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  const paths = [];

  for (const blip of doc.getElementsByTagNameNS(
    drawingNS,
    "blip"
  )) {
    const relationId = blip.getAttributeNS(
      relationshipNS,
      "embed"
    );

    const relation = relationships.get(relationId);

    if (
      !relation ||
      !relation.type.endsWith("/image") ||
      !relation.target
    ) {
      continue;
    }

    const url = new URL(
      relation.target,
      "https://ppt.local/ppt/slides/"
    );

    if (url.origin !== "https://ppt.local") continue;

    const path = decodeURIComponent(url.pathname.slice(1));

    if (
      /\.(png|jpg|jpeg|webp)$/i.test(path) &&
      zip.file(path) &&
      !paths.includes(path)
    ) {
      paths.push(path);
    }
  }

  return paths;
}

// --------------------------------------------------
// CONVERT IMAGE FOR TEXT RECOGNITION
// --------------------------------------------------

async function imageCanvas(bytes, filename) {
  const lower = filename.toLowerCase();

  const mime = lower.endsWith(".png")
    ? "image/png"
    : lower.endsWith(".webp")
      ? "image/webp"
      : "image/jpeg";

  const blob = new Blob([bytes], { type: mime });
  const url = URL.createObjectURL(blob);

  try {
    const img = new Image();

    img.src = url;
    await img.decode();

    // Keep reasonably high resolution for OCR.
    const scale = Math.min(
      1.5,
      2400 / Math.max(img.width, img.height)
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

    if (!ctx) {
      throw new Error("Canvas is not supported.");
    }

    ctx.fillStyle = "#ffffff";

    ctx.fillRect(
      0,
      0,
      canvas.width,
      canvas.height
    );

    ctx.drawImage(
      img,
      0,
      0,
      canvas.width,
      canvas.height
    );

    return canvas;
  } finally {
    URL.revokeObjectURL(url);
  }
}

// --------------------------------------------------
// LOAD OCR ENGINE
// --------------------------------------------------

async function getOCR() {
  if (!window.Tesseract) {
    throw new Error(
      "Image text recognition could not load. " +
      "Check your internet connection."
    );
  }

  if (!ocrWorker) {
    status("Loading image text recognition...");

    ocrWorker = await Tesseract.createWorker("eng", 1);
  }

  return ocrWorker;
}

// --------------------------------------------------
// EXTRACT POWERPOINT SLIDES
// --------------------------------------------------

async function readPPT(file) {
  if (!window.JSZip) {
    throw new Error(
      "PowerPoint reader could not load. Refresh the page."
    );
  }

  const zip = await JSZip.loadAsync(file);
  const slideFiles = [];

  zip.forEach((path, entry) => {
    const match =
      /^ppt\/slides\/slide(\d+)\.xml$/.exec(path);

    if (match && !entry.dir) {
      slideFiles.push({
        number: Number(match[1]),
        path
      });
    }
  });

  slideFiles.sort((a, b) => a.number - b.number);

  if (!slideFiles.length) {
    throw new Error("No readable PowerPoint slides found.");
  }

  if (slideFiles.length > 100) {
    throw new Error("Maximum 100 slides are supported.");
  }

  const slides = [];

  for (let i = 0; i < slideFiles.length; i++) {
    const slide = slideFiles[i];

    status(
      "Reading slide " +
      (i + 1) +
      " of " +
      slideFiles.length +
      "..."
    );

    const xml = await zip.file(slide.path).async("string");
    const doc = xmlDoc(xml);

    const textParts = [];

    const nativeText = slideText(doc).trim();

    if (nativeText) {
      textParts.push(nativeText);
    }

    const images = await imagesForSlide(
      zip,
      slide.number,
      doc
    );

    // OCR text in embedded slide pictures.
    // Limit images to avoid extremely slow processing.
    for (const path of images.slice(0, 5)) {
      try {
        status(
          "Reading picture text on slide " +
          slide.number +
          "..."
        );

        const bytes = await zip.file(path)
          .async("uint8array");

        const canvas = await imageCanvas(
          bytes,
          path
        );

        const worker = await getOCR();

        const result = await worker.recognize(canvas);

        const recognizedText =
          (result.data.text || "").trim();

        const confidence =
          Number(result.data.confidence || 0);

        // Reject very unreliable OCR results.
        if (
          recognizedText.length >= 20 &&
          confidence >= 40
        ) {
          textParts.push(recognizedText);
        }
      } catch (error) {
        console.warn(
          "Could not OCR image on slide",
          slide.number,
          error
        );
      }
    }

    const combinedText = textParts
      .join("\n")
      .replace(/\r/g, "")
      .replace(/[ \t]+/g, " ")
      .trim();

    if (combinedText.length >= 20) {
      slides.push({
        number: slide.number,
        text: combinedText.slice(0, 6500)
      });
    }
  }

  if (!slides.length) {
    throw new Error(
      "No readable slide text was found. " +
      "Use clearer slide images or a PowerPoint " +
      "with editable text."
    );
  }

  return slides;
}

// --------------------------------------------------
// CONNECT TO OPENAI VIA VERCEL
// --------------------------------------------------

async function aiQuiz(slides, count, instructions) {
  let response;

  try {
    response = await fetch("/api/quiz", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        slides,
        count,
        instructions
      })
    });
  } catch {
    throw new Error(
      "Could not connect to AI server. " +
      "Please check the internet connection."
    );
  }

  const raw = await response.text();

  let data;

  try {
    data = JSON.parse(raw);
  } catch {
    throw new Error(
      "AI server returned an invalid response. " +
      "Check the Vercel deployment logs."
    );
  }

  if (!response.ok) {
    throw new Error(
      data.error ||
      "AI service failed with status " + response.status
    );
  }

  if (
    !Array.isArray(data.questions) ||
    data.questions.length !== count
  ) {
    throw new Error(
      "AI did not generate the requested number " +
      "of questions. Please retry."
    );
  }

  // Basic browser-side checks.
  const slideNumbers = new Set(
    slides.map(slide => slide.number)
  );

  const seen = new Set();

  for (const question of data.questions) {
    const questionText =
      String(question.question || "").trim();

    if (
      !questionText ||
      !question.answer ||
      !question.explanation ||
      !slideNumbers.has(Number(question.slideNumber))
    ) {
      throw new Error(
        "AI generated an incomplete or unsupported " +
        "question. Please try again."
      );
    }

    const key = questionText.toLowerCase();

    if (seen.has(key)) {
      throw new Error(
        "AI generated duplicate questions. Please retry."
      );
    }

    seen.add(key);
  }

  return data.questions;
}

// --------------------------------------------------
// GENERATE QUIZ
// --------------------------------------------------

$("generateBtn").addEventListener("click", async () => {
  const file = $("pptFile").files[0];

  if (!file) {
    status("Please select a PowerPoint first.", true);
    return;
  }

  const button = $("generateBtn");

  button.disabled = true;
  button.textContent = "Preparing AI quiz...";

  try {
    status("Reading your PowerPoint...");

    const slides = await readPPT(file);

    const count = Number(
      $("questionCount").value
    );

    const instructions =
      $("instructions").value.trim();

    status(
      "Generating " +
      count +
      " meaningful AI questions from " +
      slides.length +
      " readable slides..."
    );

    // AI ONLY. No unreliable local fallback.
    questions = await aiQuiz(
      slides,
      count,
      instructions
    );

    position = 0;
    score = 0;

    renderQuestion();
    show($("quizScreen"));

  } catch (error) {
    console.error(
      "Quiz generation error:",
      error
    );

    status(
      error.message || "Could not create quiz.",
      true
    );
  } finally {
    button.disabled = false;
    button.textContent = "✦ Generate quiz";
  }
});

// --------------------------------------------------
// QUIZ DISPLAY HELPERS
// --------------------------------------------------

function addLine(parent, value, tag = "p") {
  const element = document.createElement(tag);

  element.textContent = value;
  parent.appendChild(element);

  return element;
}

function renderQuestion() {
  const q = questions[position];

  if (!q) return;

  selected = -1;
  answered = false;

  const name =
    $("studentName").value.trim() || "Student";

  $("studentLabel").textContent =
    name + "'s quiz";

  $("questionNumber").textContent =
    "Question " +
    (position + 1) +
    " of " +
    questions.length;

  $("progressBar").style.width =
    (100 * (position + 1) / questions.length) +
    "%";

  const labels = {
    mcq: "Multiple-choice question",
    one_word: "One-word answer",
    short: "Short answer",
    long: "Detailed answer"
  };

  $("questionType").textContent =
    labels[q.type] || "Question";

  $("slideNumber").textContent =
    "Slide " + q.slideNumber;

  $("questionText").textContent =
    q.question;

  $("feedback").className = "feedback hidden";
  $("feedback").replaceChildren();

  $("nextBtn").classList.add("hidden");
  $("submitBtn").classList.remove("hidden");

  const area = $("answerArea");
  area.replaceChildren();

  if (
    q.type === "mcq" &&
    Array.isArray(q.options)
  ) {
    q.options.forEach((option, index) => {
      const button = document.createElement("button");

      button.type = "button";
      button.className = "option";

      button.textContent =
        String.fromCharCode(65 + index) +
        ". " +
        option;

      button.addEventListener("click", () => {
        if (answered) return;

        selected = index;

        area.querySelectorAll(".option")
          .forEach(item => {
            item.classList.remove("selected");
          });

        button.classList.add("selected");
      });

      area.appendChild(button);
    });

  } else {
    const field = document.createElement("textarea");

    field.id = "writtenAnswer";
    field.className = "answer-input";

    field.rows = q.type === "long" ? 7 : 3;

    field.placeholder =
      "Type your answer here...";

    area.appendChild(field);
  }
}

// --------------------------------------------------
// SHOW ANSWERS AND EXPLANATIONS
// --------------------------------------------------

function showFeedback(
  question,
  heading,
  correct = null
) {
  const feedback = $("feedback");

  feedback.replaceChildren();

  feedback.className =
    "feedback" +
    (correct === null
      ? ""
      : correct
        ? " success"
        : " error");

  addLine(feedback, heading, "strong");

  addLine(
    feedback,
    "Correct answer: " + question.answer
  );

  addLine(
    feedback,
    "Explanation: " + question.explanation
  );

  addLine(
    feedback,
    "Source: Slide " + question.slideNumber
  );

  if (question.evidence) {
    addLine(
      feedback,
      "Slide evidence: " + question.evidence
    );
  }
}

// --------------------------------------------------
// CHECK ANSWER
// --------------------------------------------------

$("submitBtn").addEventListener("click", () => {
  if (answered) return;

  const q = questions[position];

  if (
    q.type === "mcq" &&
    Array.isArray(q.options)
  ) {
    if (selected < 0) {
      alert("Please choose an answer first.");
      return;
    }

    answered = true;

    const correct =
      selected === q.correctIndex;

    if (correct) score++;

    $("answerArea")
      .querySelectorAll(".option")
      .forEach((button, index) => {
        button.disabled = true;

        if (index === q.correctIndex) {
          button.classList.add("correct");
        }

        if (index === selected && !correct) {
          button.classList.add("wrong");
        }
      });

    showFeedback(
      q,
      correct
        ? "✓ Correct answer!"
        : "Incorrect answer",
      correct
    );

    $("nextBtn").classList.remove("hidden");

  } else {
    const field = $("writtenAnswer");

    if (!field.value.trim()) {
      alert("Please type your answer first.");
      return;
    }

    answered = true;
    field.disabled = true;

    showFeedback(
      q,
      "Compare your answer with the correct answer"
    );

    const feedback = $("feedback");

    addLine(
      feedback,
      "Did your answer cover the important points?"
    );

    const choices = [
      ["✓ Yes, I got it right", true],
      ["I need more practice", false]
    ];

    choices.forEach(([label, correct]) => {
      const button = document.createElement("button");

      button.type = "button";

      button.className =
        correct ? "primary" : "secondary";

      button.textContent = label;

      button.addEventListener("click", () => {
        if (correct) score++;

        feedback.querySelectorAll("button")
          .forEach(item => item.remove());

        $("nextBtn").classList.remove("hidden");
      });

      feedback.appendChild(button);
    });
  }

  $("submitBtn").classList.add("hidden");
});

// --------------------------------------------------
// NEXT QUESTION
// --------------------------------------------------

$("nextBtn").addEventListener("click", () => {
  position++;

  if (position < questions.length) {
    renderQuestion();
    window.scrollTo(0, 0);
    return;
  }

  const name =
    $("studentName").value.trim() || "Student";

  $("resultName").textContent =
    "Well done, " + name + "!";

  $("finalScore").textContent =
    score + " / " + questions.length;

  $("resultMessage").textContent =
    "Your quiz is complete. " +
    "Written answers were self-assessed.";

  show($("resultScreen"));
});

// --------------------------------------------------
// RESTART QUIZ
// --------------------------------------------------

$("restartBtn").addEventListener("click", () => {
  position = 0;
  score = 0;

  renderQuestion();
  show($("quizScreen"));
});

// --------------------------------------------------
// CREATE NEW QUIZ
// --------------------------------------------------

$("newQuizBtn").addEventListener("click", () => {
  questions = [];
  position = 0;
  score = 0;

  show($("setupScreen"));
});

// --------------------------------------------------
// CLEAN UP OCR
// --------------------------------------------------

window.addEventListener("beforeunload", () => {
  if (ocrWorker) {
    ocrWorker.terminate();
  }
});
