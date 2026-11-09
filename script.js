
"use strict";

// SLIDESPARK AI - POWERPOINT VISION VERSION
// Maximum PowerPoint file size: 30 MB

const MAX_FILE_SIZE = 30 * 1024 * 1024;
const $ = (id) => document.getElementById(id);

let questions = [];
let currentIndex = 0;
let score = 0;
let selectedOption = null;
let answered = false;

const setupScreen = $("setupScreen");
const quizScreen = $("quizScreen");
const resultScreen = $("resultScreen");

function showScreen(screen) {
  [setupScreen, quizScreen, resultScreen].forEach((item) => {
    item.classList.add("hidden");
  });
  screen.classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function setStatus(message, error = false) {
  const element = $("statusMessage");
  element.textContent = message;
  element.classList.toggle("error", error);
}

function formatSize(size) {
  return (size / (1024 * 1024)).toFixed(2) + " MB";
}

$("pptFile").addEventListener("change", (event) => {
  const file = event.target.files[0];

  if (!file) {
    $("fileName").textContent = "No presentation selected";
    return;
  }

  if (!file.name.toLowerCase().endsWith(".pptx")) {
    setStatus("Please select a PPTX file.", true);
    event.target.value = "";
    return;
  }

  if (file.size > MAX_FILE_SIZE) {
    setStatus("Maximum PowerPoint size is 30 MB.", true);
    event.target.value = "";
    return;
  }

  $("fileName").textContent =
    "✅ " + file.name + " (" + formatSize(file.size) + ")";
  setStatus("PowerPoint selected successfully.");
});

// Convert each slide picture into a smaller JPEG.
// This helps stay within Vercel request limits.

async function compressImage(blob) {
  const url = URL.createObjectURL(blob);

  try {
    const img = new Image();
    img.src = url;
    await img.decode();

    const maxWidth = 1200;
    const maxHeight = 1200;
    const scale = Math.min(
      1,
      maxWidth / img.width,
      maxHeight / img.height
    );

    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(img.width * scale));
    canvas.height = Math.max(1, Math.round(img.height * scale));

    const ctx = canvas.getContext("2d");

    if (!ctx) {
      throw new Error("Image conversion is unavailable.");
    }

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

    return canvas.toDataURL("image/jpeg", 0.65);
  } finally {
    URL.revokeObjectURL(url);
  }
}

function parseXML(source) {
  const doc = new DOMParser().parseFromString(
    source,
    "application/xml"
  );

  if (doc.getElementsByTagName("parsererror").length) {
    throw new Error("Invalid PowerPoint slide XML.");
  }

  return doc;
}

function getSlideText(xml) {
  const ns =
    "http://schemas.openxmlformats.org/drawingml/2006/main";

  const paragraphs = Array.from(
    xml.getElementsByTagNameNS(ns, "p")
  );

  return paragraphs.map((paragraph) => {
    const pieces = paragraph.getElementsByTagNameNS(ns, "t");

    return Array.from(pieces)
      .map((piece) => piece.textContent)
      .join("")
      .trim();
  }).filter(Boolean).join("\n").slice(0, 6000);
}

async function getSlideImage(zip, slideNumber, slideXML) {
  const relPath =
    `ppt/slides/_rels/slide${slideNumber}.xml.rels`;

  const relationshipsFile = zip.file(relPath);

  if (!relationshipsFile) {
    return null;
  }

  const relationshipXML = parseXML(
    await relationshipsFile.async("string")
  );

  const relationships = new Map();

  Array.from(
    relationshipXML.getElementsByTagName("Relationship")
  ).forEach((rel) => {
    relationships.set(rel.getAttribute("Id"), {
      type: rel.getAttribute("Type") || "",
      target: rel.getAttribute("Target") || ""
    });
  });

  const drawingNS =
    "http://schemas.openxmlformats.org/drawingml/2006/main";

  const relNS =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  const blips = Array.from(
    slideXML.getElementsByTagNameNS(drawingNS, "blip")
  );

  for (const blip of blips) {
    const id = blip.getAttributeNS(relNS, "embed");
    const rel = relationships.get(id);

    if (!rel || !rel.type.endsWith("/image")) {
      continue;
    }

    if (!rel.target) continue;

    let imagePath;

    if (rel.target.startsWith("/")) {
      imagePath = rel.target.replace(/^\/+/, "");
    } else {
      const resolved = new URL(
        rel.target,
        "https://slidespark.local/ppt/slides/"
      );

      if (resolved.origin !== "https://slidespark.local") {
        continue;
      }

      imagePath = decodeURIComponent(
        resolved.pathname.replace(/^\/+/, "")
      );
    }

    if (!/\.(png|jpg|jpeg|webp)$/i.test(imagePath)) {
      continue;
    }

    const imageFile = zip.file(imagePath);
    if (!imageFile) continue;

    const blob = await imageFile.async("blob");

    try {
      return await compressImage(blob);
    } catch (err) {
      console.warn("Could not read slide image:", err);
    }
  }

  return null;
}

// Read editable text and embedded pictures.

async function extractSlides(file) {
  if (typeof JSZip === "undefined") {
    throw new Error(
      "PowerPoint reader is unavailable. Check your internet."
    );
  }

  if (file.size > MAX_FILE_SIZE) {
    throw new Error("Maximum file size is 30 MB.");
  }

  const zip = await JSZip.loadAsync(file);
  const files = [];

  zip.forEach((path, entry) => {
    const match = path.match(
      /^ppt\/slides\/slide(\d+)\.xml$/
    );

    if (match && !entry.dir) {
      files.push({
        number: Number(match[1]),
        path
      });
    }
  });

  files.sort((a, b) => a.number - b.number);

  if (!files.length) {
    throw new Error("No PowerPoint slides found.");
  }

  if (files.length > 100) {
    throw new Error("Maximum 100 slides supported.");
  }

  const slides = [];

  for (const fileInfo of files) {
    setStatus(
      `Opening slide ${fileInfo.number} of ${files.length}...`
    );

    const xml = parseXML(
      await zip.file(fileInfo.path).async("string")
    );

    const text = getSlideText(xml);
    const image = await getSlideImage(
      zip,
      fileInfo.number,
      xml
    );

    if (text || image) {
      slides.push({
        number: fileInfo.number,
        text,
        image
      });
    }
  }

  if (!slides.length) {
    throw new Error(
      "No supported text or slide pictures were found."
    );
  }

  return slides;
}

// Send one image at a time for AI Vision reading.
// The full PowerPoint never needs to be uploaded.

async function callAPI(payload) {
  const response = await fetch("/api/quiz", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  let data;

  try {
    data = await response.json();
  } catch {
    throw new Error(
      "Server returned an unexpected response. Check Vercel logs."
    );
  }

  if (!response.ok) {
    throw new Error(
      data.error || "AI service request failed."
    );
  }

  return data;
}

async function readSlidesWithAI(slides) {
  const readableSlides = [];

  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i];

    let text = slide.text || "";

    if (slide.image) {
      setStatus(
        `🔍 AI reading slide ${slide.number} (${i + 1}/${slides.length})...`
      );

      const result = await callAPI({
        mode: "read",
        slideNumber: slide.number,
        image: slide.image
      });

      const imageText = String(result.text || "").trim();

      text = [text, imageText]
        .filter(Boolean)
        .join("\n");
    }

    if (text.trim()) {
      readableSlides.push({
        number: slide.number,
        text: text.slice(0, 6000)
      });
    }
  }

  if (!readableSlides.length) {
    throw new Error(
      "AI could not read educational content from these slides."
    );
  }

  return readableSlides;
}

// Generate quiz.

$("generateBtn").addEventListener("click", async () => {
  const file = $("pptFile").files[0];

  if (!file) {
    setStatus("Please select your PowerPoint.", true);
    return;
  }

  const button = $("generateBtn");
  button.disabled = true;
  button.textContent = "⏳ Creating Quiz...";

  try {
    setStatus("📚 Opening PowerPoint...");

    const slides = await extractSlides(file);
    const readableSlides = await readSlidesWithAI(slides);

    setStatus("✨ Preparing questions from your slides...");

    const result = await callAPI({
      mode: "generate",
      slides: readableSlides,
      count: Number($("questionCount").value),
      instructions: $("instructions").value.trim()
    });

    if (!Array.isArray(result.questions) ||
        result.questions.length === 0) {
      throw new Error("No questions were generated.");
    }

    questions = result.questions;
    currentIndex = 0;
    score = 0;

    renderQuestion();
    showScreen(quizScreen);
    setStatus("");
  } catch (error) {
    console.error(error);
    setStatus(error.message || "Something went wrong.", true);
  } finally {
    button.disabled = false;
    button.textContent = "✨ Generate My Quiz";
  }
});

// Quiz interface

function renderQuestion() {
  const question = questions[currentIndex];

  selectedOption = null;
  answered = false;

  const name = $("studentName").value.trim() || "Student";

  $("studentLabel").textContent = name + "'s Quiz";
  $("questionNumber").textContent =
    `Question ${currentIndex + 1} of ${questions.length}`;

  $("progressBar").style.width =
    `${((currentIndex + 1) / questions.length) * 100}%`;

  const labels = {
    mcq: "Choose the Correct Answer",
    one_word: "One-Word Answer",
    short: "Short Answer",
    long: "Answer the Following"
  };

  $("questionType").textContent =
    labels[question.type] || "Question";

  $("slideNumber").textContent =
    "📑 Slide " + question.slideNumber;

  $("questionText").textContent = question.question;

  $("feedback").replaceChildren();
  $("feedback").className = "feedback hidden";

  $("submitBtn").classList.remove("hidden");
  $("submitBtn").disabled = false;
  $("nextBtn").classList.add("hidden");

  const area = $("answerArea");
  area.replaceChildren();

  if (question.type === "mcq") {
    question.options.forEach((option, index) => {
      const button = document.createElement("button");

      button.type = "button";
      button.className = "option";
      button.textContent =
        `${String.fromCharCode(65 + index)}. ${option}`;

      button.addEventListener("click", () => {
        if (answered) return;

        selectedOption = index;

        area.querySelectorAll(".option").forEach((item) => {
          item.classList.remove("selected");
        });

        button.classList.add("selected");
      });

      area.appendChild(button);
    });
  } else {
    const input = document.createElement("textarea");

    input.id = "writtenAnswer";
    input.className = "answer-input";
    input.rows = question.type === "long" ? 7 : 3;
    input.placeholder = "✍️ Type your answer here...";

    area.appendChild(input);
  }
}

function showFeedback(title, question, correct) {
  const box = $("feedback");
  box.replaceChildren();
  box.className = "feedback";

  if (correct === true) box.classList.add("success");
  if (correct === false) box.classList.add("error");

  const heading = document.createElement("strong");
  heading.textContent = title;
  box.appendChild(heading);

  const lines = [
    "✅ Answer: " + question.answer,
    "💡 Explanation: " + question.explanation,
    `📑 Slide ${question.slideNumber}: "${question.evidence}"`
  ];

  lines.forEach((line) => {
    const p = document.createElement("p");
    p.textContent = line;
    box.appendChild(p);
  });
}

$("submitBtn").addEventListener("click", () => {
  if (answered) return;

  const question = questions[currentIndex];

  if (question.type === "mcq") {
    if (selectedOption === null) {
      alert("Please select an option.");
      return;
    }

    answered = true;

    const correct =
      selectedOption === question.correctIndex;

    if (correct) score++;

    document.querySelectorAll(".option").forEach(
      (button, index) => {
        button.disabled = true;

        if (index === question.correctIndex) {
          button.classList.add("correct");
        }

        if (index === selectedOption && !correct) {
          button.classList.add("wrong");
        }
      }
    );

    showFeedback(
      correct ? "✅ Correct!" : "❌ Not Quite!",
      question,
      correct
    );

    $("nextBtn").classList.remove("hidden");
  } else {
    const written = $("writtenAnswer").value.trim();

    if (!written) {
      alert("Please type your answer.");
      return;
    }

    answered = true;
    $("writtenAnswer").disabled = true;

    showFeedback("📖 Check Your Answer", question, null);

    const feedback = $("feedback");

    const message = document.createElement("p");
    message.textContent =
      "Compare your answer with the model answer above.";

    const yes = document.createElement("button");
    yes.className = "primary-btn";
    yes.textContent = "✅ I got it right";

    const no = document.createElement("button");
    no.className = "secondary-btn";
    no.textContent = "📚 I need more practice";

    yes.addEventListener("click", () => {
      score++;
      finishWrittenAnswer();
    });

    no.addEventListener("click", finishWrittenAnswer);

    feedback.appendChild(message);
    feedback.appendChild(yes);
    feedback.appendChild(no);
  }

  $("submitBtn").classList.add("hidden");
});

function finishWrittenAnswer() {
  $("feedback").querySelectorAll("button").forEach(
    (button) => button.remove()
  );

  $("nextBtn").classList.remove("hidden");
}

$("nextBtn").addEventListener("click", () => {
  currentIndex++;

  if (currentIndex < questions.length) {
    renderQuestion();
    window.scrollTo({ top: 0, behavior: "smooth" });
  } else {
    showResults();
  }
});

function showResults() {
  const name = $("studentName").value.trim() || "Student";
  const percent = Math.round(
    (score / questions.length) * 100
  );

  $("resultName").textContent =
    "Well done, " + name + "!";

  $("finalScore").textContent =
    `${score} / ${questions.length}`;

  $("resultMessage").textContent =
    `You scored ${percent}%. Written answers were self-assessed.`;

  showScreen(resultScreen);
}

$("restartBtn").addEventListener("click", () => {
  currentIndex = 0;
  score = 0;
  renderQuestion();
  showScreen(quizScreen);
});

$("newQuizBtn").addEventListener("click", () => {
  questions = [];
  currentIndex = 0;
  score = 0;
  setStatus("");
  showScreen(setupScreen);
});

console.log("SlideSpark AI Vision - 30 MB version loaded");
