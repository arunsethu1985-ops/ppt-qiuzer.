
"use strict";

// ============================================
// SLIDESPARK AI QUIZ - VERSION 2.0
// Maximum PowerPoint size: 30 MB
// ============================================

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

// ============================================
// SCREEN MANAGEMENT
// ============================================

function showScreen(screen) {
  [setupScreen, quizScreen, resultScreen].forEach((element) => {
    element.classList.add("hidden");
  });

  screen.classList.remove("hidden");

  window.scrollTo({
    top: 0,
    behavior: "smooth"
  });
}

function setStatus(message, isError = false) {
  const status = $("statusMessage");

  status.textContent = message;
  status.classList.toggle("error", isError);
}

function formatFileSize(bytes) {
  return (bytes / (1024 * 1024)).toFixed(2) + " MB";
}

// ============================================
// POWERPOINT UPLOAD - 30 MB
// ============================================

$("pptFile").addEventListener("change", function (event) {
  const file = event.target.files[0];

  if (!file) {
    $("fileName").textContent = "No presentation selected";
    return;
  }

  if (!file.name.toLowerCase().endsWith(".pptx")) {
    $("fileName").textContent = "Invalid file";
    setStatus("Please upload a .pptx PowerPoint file.", true);
    event.target.value = "";
    return;
  }

  if (file.size > MAX_FILE_SIZE) {
    $("fileName").textContent = "File exceeds 30 MB";
    setStatus(
      "Please use a PowerPoint smaller than 30 MB.",
      true
    );
    event.target.value = "";
    return;
  }

  $("fileName").textContent =
    "✅ " + file.name + " (" + formatFileSize(file.size) + ")";

  setStatus("Presentation selected successfully.");
});

// ============================================
// EXTRACT TEXT FROM POWERPOINT
// ============================================

async function extractSlides(file) {
  if (typeof JSZip === "undefined") {
    throw new Error(
      "PowerPoint reader not loaded. Please check your internet connection."
    );
  }

  if (!file.name.toLowerCase().endsWith(".pptx")) {
    throw new Error("Only .pptx files are supported.");
  }

  if (file.size > MAX_FILE_SIZE) {
    throw new Error(
      "Please use a PowerPoint smaller than 30 MB."
    );
  }

  const zip = await JSZip.loadAsync(file);

  const slideFiles = [];

  zip.forEach((path, entry) => {
    const match = path.match(
      /^ppt\/slides\/slide(\d+)\.xml$/
    );

    if (match && !entry.dir) {
      slideFiles.push({
        number: Number(match[1]),
        path: path
      });
    }
  });

  slideFiles.sort((a, b) => a.number - b.number);

  if (slideFiles.length === 0) {
    throw new Error("No slides found in this PowerPoint.");
  }

  if (slideFiles.length > 100) {
    throw new Error(
      "This version supports a maximum of 100 slides."
    );
  }

  const slides = [];

  for (const slide of slideFiles) {
    const xmlString = await zip.file(slide.path).async("string");

    const xml = new DOMParser().parseFromString(
      xmlString,
      "application/xml"
    );

    if (xml.getElementsByTagName("parsererror").length) {
      throw new Error(
        "Unable to read PowerPoint slide " + slide.number
      );
    }

    const namespace =
      "http://schemas.openxmlformats.org/drawingml/2006/main";

    const paragraphs = Array.from(
      xml.getElementsByTagNameNS(namespace, "p")
    );

    const lines = [];

    for (const paragraph of paragraphs) {
      const textElements =
        paragraph.getElementsByTagNameNS(namespace, "t");

      const text = Array.from(textElements)
        .map((element) => element.textContent)
        .join("")
        .trim();

      if (text) {
        lines.push(text);
      }
    }

    const slideText = lines.join("\n").slice(0, 6000);

    if (slideText.length > 0) {
      slides.push({
        number: slide.number,
        text: slideText
      });
    }
  }

  if (slides.length === 0) {
    throw new Error(
      "No editable text found in this PowerPoint. " +
      "If your slides contain text inside images, " +
      "image-reading AI is required."
    );
  }

  return slides;
}

// ============================================
// GENERATE QUESTIONS USING AI
// ============================================

$("generateBtn").addEventListener("click", async function () {
  const file = $("pptFile").files[0];

  if (!file) {
    setStatus("Please select a PowerPoint first.", true);
    return;
  }

  if (file.size > MAX_FILE_SIZE) {
    setStatus(
      "Please use a PowerPoint smaller than 30 MB.",
      true
    );
    return;
  }

  const name = $("studentName").value.trim() || "Student";
  const instructions = $("instructions").value.trim();
  const count = Number($("questionCount").value);

  const button = $("generateBtn");

  button.disabled = true;
  button.textContent = "⏳ Generating Quiz...";

  try {
    setStatus("📖 Reading your PowerPoint...");

    const slides = await extractSlides(file);

    setStatus(
      `✅ Read ${slides.length} slides containing text. Generating questions...`
    );

    const response = await fetch("/api/quiz", {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        slides: slides,
        instructions: instructions,
        count: count,
        name: name
      })
    });

    let data;

    try {
      data = await response.json();
    } catch {
      throw new Error(
        "The server returned an invalid response. " +
        "Check the Vercel deployment logs."
      );
    }

    if (!response.ok) {
      throw new Error(
        data.error || "Unable to generate the quiz."
      );
    }

    if (!Array.isArray(data.questions) ||
        data.questions.length === 0) {
      throw new Error("No questions were generated.");
    }

    questions = data.questions;
    currentIndex = 0;
    score = 0;

    renderQuestion();
    showScreen(quizScreen);
    setStatus("");

  } catch (error) {
    console.error("SlideSpark Error:", error);

    setStatus(
      error.message || "Something went wrong.",
      true
    );
  } finally {
    button.disabled = false;
    button.textContent = "✨ Generate My Quiz";
  }
});

// ============================================
// DISPLAY ONE QUESTION AT A TIME
// ============================================

function renderQuestion() {
  const question = questions[currentIndex];

  selectedOption = null;
  answered = false;

  const name = $("studentName").value.trim() || "Student";

  $("studentLabel").textContent = name + "'s Quiz";

  $("questionNumber").textContent =
    `Question ${currentIndex + 1} of ${questions.length}`;

  const percentage =
    ((currentIndex + 1) / questions.length) * 100;

  $("progressBar").style.width = percentage + "%";

  const typeNames = {
    mcq: "Choose the correct answer",
    one_word: "One Word Answer",
    short: "Short Answer",
    long: "Answer the Following"
  };

  $("questionType").textContent =
    typeNames[question.type] || "Question";

  $("slideNumber").textContent =
    "📑 Slide " + question.slideNumber;

  $("questionText").textContent = question.question;

  $("feedback").className = "feedback hidden";
  $("feedback").replaceChildren();

  $("submitBtn").classList.remove("hidden");
  $("submitBtn").disabled = false;

  $("nextBtn").classList.add("hidden");

  const answerArea = $("answerArea");
  answerArea.replaceChildren();

  if (question.type === "mcq") {
    renderMultipleChoice(question, answerArea);
  } else {
    renderWrittenAnswer(question, answerArea);
  }
}

// ============================================
// MULTIPLE CHOICE QUESTIONS
// ============================================

function renderMultipleChoice(question, container) {
  question.options.forEach((option, index) => {
    const button = document.createElement("button");

    button.type = "button";
    button.className = "option";

    const letter = String.fromCharCode(65 + index);

    button.textContent = letter + ". " + option;

    button.addEventListener("click", function () {
      if (answered) return;

      selectedOption = index;

      document.querySelectorAll(".option").forEach((item) => {
        item.classList.remove("selected");
      });

      button.classList.add("selected");
    });

    container.appendChild(button);
  });
}

// ============================================
// WRITTEN ANSWERS
// ============================================

function renderWrittenAnswer(question, container) {
  const textarea = document.createElement("textarea");

  textarea.id = "writtenAnswer";
  textarea.className = "answer-input";

  textarea.placeholder =
    question.type === "one_word"
      ? "✍️ Type your one-word answer..."
      : "✍️ Write your answer here...";

  textarea.rows = question.type === "long" ? 7 : 3;

  container.appendChild(textarea);
}

// ============================================
// CHECK ANSWERS
// ============================================

$("submitBtn").addEventListener("click", function () {
  if (answered) return;

  const question = questions[currentIndex];

  if (question.type === "mcq") {
    if (selectedOption === null) {
      alert("Please select an answer first.");
      return;
    }

    answered = true;

    const correct =
      selectedOption === question.correctIndex;

    if (correct) {
      score++;
    }

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
      correct ? "✅ Correct Answer!" : "❌ Not Quite!",
      question,
      correct
    );

    $("nextBtn").classList.remove("hidden");

  } else {
    const input = $("writtenAnswer");
    const written = input.value.trim();

    if (!written) {
      alert("Please write your answer first.");
      return;
    }

    answered = true;
    input.disabled = true;

    showFeedback(
      "📖 Compare Your Answer",
      question,
      null
    );

    const feedback = $("feedback");

    const message = document.createElement("p");
    message.textContent =
      "Compare your response with the model answer. " +
      "Did you answer correctly?";

    feedback.appendChild(message);

    const correctButton = document.createElement("button");
    correctButton.type = "button";
    correctButton.className = "primary-btn";
    correctButton.textContent = "✅ Yes, I got it right";

    const practiceButton = document.createElement("button");
    practiceButton.type = "button";
    practiceButton.className = "secondary-btn";
    practiceButton.textContent = "📚 I need more practice";

    correctButton.addEventListener("click", function () {
      score++;
      finishWrittenAnswer();
    });

    practiceButton.addEventListener("click", function () {
      finishWrittenAnswer();
    });

    feedback.appendChild(correctButton);
    feedback.appendChild(practiceButton);
  }

  $("submitBtn").classList.add("hidden");
});

function finishWrittenAnswer() {
  $("feedback").querySelectorAll("button").forEach(
    (button) => button.remove()
  );

  $("nextBtn").classList.remove("hidden");
}

// ============================================
// DISPLAY CORRECT ANSWER AND SLIDE REFERENCE
// ============================================

function showFeedback(title, question, correct) {
  const box = $("feedback");

  box.replaceChildren();
  box.className = "feedback";

  if (correct === true) {
    box.classList.add("success");
  }

  if (correct === false) {
    box.classList.add("error");
  }

  const heading = document.createElement("strong");
  heading.textContent = title;

  const answer = document.createElement("p");
  answer.textContent = "✅ Answer: " + question.answer;

  const explanation = document.createElement("p");
  explanation.textContent =
    "💡 Explanation: " + question.explanation;

  const source = document.createElement("p");
  source.textContent =
    `📑 Slide ${question.slideNumber}: "${question.evidence}"`;

  box.appendChild(heading);
  box.appendChild(answer);
  box.appendChild(explanation);
  box.appendChild(source);
}

// ============================================
// NEXT QUESTION
// ============================================

$("nextBtn").addEventListener("click", function () {
  currentIndex++;

  if (currentIndex < questions.length) {
    renderQuestion();

    window.scrollTo({
      top: 0,
      behavior: "smooth"
    });
  } else {
    showResults();
  }
});

// ============================================
// FINAL RESULTS
// ============================================

function showResults() {
  const name = $("studentName").value.trim() || "Student";

  const percentage = Math.round(
    (score / questions.length) * 100
  );

  $("resultName").textContent =
    "Amazing work, " + name + "!";

  $("finalScore").textContent =
    score + " / " + questions.length;

  let message = "";

  if (percentage >= 90) {
    message = "🌟 Outstanding performance!";
  } else if (percentage >= 70) {
    message = "🎉 Great job! Keep learning.";
  } else if (percentage >= 50) {
    message = "👍 Good effort! Keep practicing.";
  } else {
    message = "📚 Keep studying. You can improve!";
  }

  $("resultMessage").textContent =
    message + " Your score is " + percentage +
    "%. Written answers were self-assessed.";

  showScreen(resultScreen);
}

// ============================================
// RESTART QUIZ
// ============================================

$("restartBtn").addEventListener("click", function () {
  currentIndex = 0;
  score = 0;

  renderQuestion();
  showScreen(quizScreen);
});

// ============================================
// CREATE NEW QUIZ
// ============================================

$("newQuizBtn").addEventListener("click", function () {
  questions = [];
  currentIndex = 0;
  score = 0;

  setStatus("");
  showScreen(setupScreen);
});

console.log("SlideSpark AI Quiz v2.0 - 30 MB limit");
