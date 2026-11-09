
"use strict";

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

function setStatus(message, isError = false) {
  $("statusMessage").textContent = message;
  $("statusMessage").classList.toggle("error", isError);
}

$("pptFile").addEventListener("change", (event) => {
  const file = event.target.files[0];
  $("fileName").textContent = file
    ? file.name
    : "No presentation selected";
});

async function extractSlides(file) {
  if (typeof JSZip === "undefined") {
    throw new Error(
      "PowerPoint reader could not load. Check your internet connection."
    );
  }

  if (!file.name.toLowerCase().endsWith(".pptx")) {
    throw new Error("Please upload a .pptx PowerPoint file.");
  }

  if (file.size > 15 * 1024 * 1024) {
    throw new Error("Please use a PowerPoint smaller than 15 MB.");
  }

  const zip = await JSZip.loadAsync(file);
  const slideFiles = [];

  zip.forEach((path, entry) => {
    const match = path.match(/^ppt\/slides\/slide(\d+)\.xml$/);

    if (match && !entry.dir) {
      slideFiles.push({
        number: Number(match[1]),
        path
      });
    }
  });

  slideFiles.sort((a, b) => a.number - b.number);

  if (slideFiles.length === 0) {
    throw new Error("No PowerPoint slides were found.");
  }

  if (slideFiles.length > 100) {
    throw new Error("Please use a presentation with 100 slides or fewer.");
  }

  const slides = [];

  for (const slide of slideFiles) {
    const xmlText = await zip.file(slide.path).async("string");

    const xml = new DOMParser().parseFromString(
      xmlText,
      "application/xml"
    );

    if (xml.getElementsByTagName("parsererror").length) {
      throw new Error("Could not read the slide XML.");
    }

    const paragraphs = Array.from(
      xml.getElementsByTagNameNS(
        "http://schemas.openxmlformats.org/drawingml/2006/main",
        "p"
      )
    );

    const lines = paragraphs.map((paragraph) => {
      const textNodes = paragraph.getElementsByTagNameNS(
        "http://schemas.openxmlformats.org/drawingml/2006/main",
        "t"
      );

      return Array.from(textNodes)
        .map((node) => node.textContent)
        .join("")
        .trim();
    }).filter(Boolean);

    const text = lines.join("\n").slice(0, 6000);

    if (text) {
      slides.push({
        number: slide.number,
        text
      });
    }
  }

  if (slides.length === 0) {
    throw new Error(
      "No editable text was found. This version cannot read text inside slide images."
    );
  }

  return slides;
}

$("generateBtn").addEventListener("click", async () => {
  const file = $("pptFile").files[0];

  if (!file) {
    setStatus("Please select a PowerPoint first.", true);
    return;
  }

  const name = $("studentName").value.trim() || "Student";
  const instructions = $("instructions").value.trim();
  const count = Number($("questionCount").value);

  const button = $("generateBtn");
  button.disabled = true;
  button.textContent = "Generating...";

  try {
    setStatus("Reading PowerPoint slides...");

    const slides = await extractSlides(file);

    setStatus(
      `Found text on ${slides.length} slides. Generating questions...`
    );

    const response = await fetch("/api/quiz", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        slides,
        instructions,
        count,
        name
      })
    });

    const data = await response.json();

    if (!response.ok) {
      throw new Error(data.error || "Quiz generation failed.");
    }

    if (!Array.isArray(data.questions) || !data.questions.length) {
      throw new Error("The AI did not return any questions.");
    }

    questions = data.questions;
    currentIndex = 0;
    score = 0;

    renderQuestion();
    showScreen(quizScreen);
    setStatus("");
  } catch (error) {
    setStatus(error.message || "Something went wrong.", true);
  } finally {
    button.disabled = false;
    button.textContent = "✨ Generate Quiz";
  }
});

function renderQuestion() {
  const question = questions[currentIndex];

  selectedOption = null;
  answered = false;

  $("studentLabel").textContent =
    $("studentName").value.trim() || "Student";

  $("questionNumber").textContent =
    `Question ${currentIndex + 1} of ${questions.length}`;

  $("progressBar").style.width =
    `${((currentIndex + 1) / questions.length) * 100}%`;

  const typeNames = {
    mcq: "Choose the correct answer",
    one_word: "One-word answer",
    short: "Short answer",
    long: "Answer the following"
  };

  $("questionType").textContent =
    typeNames[question.type] || "Question";

  $("slideNumber").textContent =
    `Slide ${question.slideNumber}`;

  $("questionText").textContent = question.question;

  $("feedback").className = "feedback hidden";
  $("feedback").replaceChildren();

  $("submitBtn").classList.remove("hidden");
  $("submitBtn").disabled = false;
  $("nextBtn").classList.add("hidden");

  const answerArea = $("answerArea");
  answerArea.replaceChildren();

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

        document.querySelectorAll(".option").forEach((item) => {
          item.classList.remove("selected");
        });

        button.classList.add("selected");
      });

      answerArea.appendChild(button);
    });
  } else {
    const textarea = document.createElement("textarea");
    textarea.id = "writtenAnswer";
    textarea.className = "answer-input";

    textarea.placeholder =
      question.type === "one_word"
        ? "Enter your answer..."
        : "Write your answer here...";

    textarea.rows = question.type === "long" ? 7 : 3;
    answerArea.appendChild(textarea);
  }
}

$("submitBtn").addEventListener("click", () => {
  if (answered) return;

  const question = questions[currentIndex];

  if (question.type === "mcq") {
    if (selectedOption === null) {
      alert("Please select an answer.");
      return;
    }

    answered = true;

    const correct = selectedOption === question.correctIndex;

    if (correct) {
      score++;
    }

    document.querySelectorAll(".option").forEach((button, index) => {
      button.disabled = true;

      if (index === question.correctIndex) {
        button.classList.add("correct");
      }

      if (index === selectedOption && !correct) {
        button.classList.add("wrong");
      }
    });

    showFeedback(
      correct ? "✅ Correct Answer!" : "❌ Not Quite!",
      question,
      correct
    );

    $("nextBtn").classList.remove("hidden");
  } else {
    const written = $("writtenAnswer").value.trim();

    if (!written) {
      alert("Please write your answer first.");
      return;
    }

    answered = true;
    $("writtenAnswer").disabled = true;

    showFeedback(
      "📖 Compare Your Answer",
      question,
      null
    );

    const feedback = $("feedback");

    const prompt = document.createElement("p");
    prompt.textContent =
      "Did your answer match the main points in the model answer?";
    feedback.appendChild(prompt);

    const correctBtn = document.createElement("button");
    correctBtn.className = "primary-btn";
    correctBtn.textContent = "✅ Yes, I got it right";

    const practiceBtn = document.createElement("button");
    practiceBtn.className = "secondary-btn";
    practiceBtn.textContent = "📚 I need more practice";

    correctBtn.addEventListener("click", () => {
      score++;
      finishWrittenAnswer();
    });

    practiceBtn.addEventListener("click", finishWrittenAnswer);

    feedback.appendChild(correctBtn);
    feedback.appendChild(practiceBtn);
  }

  $("submitBtn").classList.add("hidden");
});

function finishWrittenAnswer() {
  $("feedback").querySelectorAll("button").forEach((button) => {
    button.remove();
  });

  $("nextBtn").classList.remove("hidden");
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

  const answer = document.createElement("p");
  answer.textContent = `Answer: ${question.answer}`;
  box.appendChild(answer);

  const explanation = document.createElement("p");
  explanation.textContent =
    `Explanation: ${question.explanation}`;
  box.appendChild(explanation);

  const evidence = document.createElement("p");
  evidence.textContent =
    `Slide ${question.slideNumber} evidence: "${question.evidence}"`;
  box.appendChild(evidence);
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
  const percentage = Math.round((score / questions.length) * 100);

  $("resultName").textContent = `Well done, ${name}!`;
  $("finalScore").textContent = `${score} / ${questions.length}`;

  $("resultMessage").textContent =
    `Your score is ${percentage}%. Written answers were self-assessed.`;

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
  showScreen(setupScreen);
});

