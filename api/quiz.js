
"use strict";

// SlideSpark AI Vision Backend
// Compatible with script.js using:
// mode: "read" and mode: "generate"

const OPENAI_URL = "https://api.openai.com/v1/chat/completions";
const MODEL = "gpt-4o-mini";

const TYPES = ["mcq", "one_word", "short", "long"];

function clean(value) {
  return String(value || "").trim();
}

function normalize(value) {
  return clean(value)
    .normalize("NFKC")
    .toLowerCase()
    .replace(/\s+/g, " ");
}

async function askAI(messages, maxTokens = 4000) {
  const controller = new AbortController();

  const timeout = setTimeout(() => {
    controller.abort();
  }, 55000);

  try {
    const response = await fetch(OPENAI_URL, {
      method: "POST",
      signal: controller.signal,
      headers: {
        "Content-Type": "application/json",
        Authorization:
          `Bearer ${process.env.OPENAI_API_KEY}`
      },
      body: JSON.stringify({
        model: MODEL,
        temperature: 0.1,
        max_tokens: maxTokens,
        response_format: {
          type: "json_object"
        },
        messages
      })
    });

    const data = await response.json();

    if (!response.ok) {
      console.error(
        "OpenAI error:",
        data.error?.message || response.status
      );

      throw new Error(
        "AI service error. Check your API key, " +
        "credits and model access."
      );
    }

    const output =
      data.choices?.[0]?.message?.content;

    if (!output) {
      throw new Error("AI returned no response.");
    }

    try {
      return JSON.parse(output);
    } catch {
      throw new Error(
        "AI returned invalid JSON. Please retry."
      );
    }
  } catch (error) {
    if (error.name === "AbortError") {
      throw new Error(
        "AI request timed out. Please retry."
      );
    }

    throw error;
  } finally {
    clearTimeout(timeout);
  }
}

// =====================================
// MODE 1 - READ SLIDE IMAGE
// =====================================

async function readSlide(body) {
  const image = body.image;

  if (
    typeof image !== "string" ||
    !/^data:image\/jpeg;base64,[A-Za-z0-9+/=]+$/.test(image)
  ) {
    throw new Error("Invalid slide image format.");
  }

  if (image.length > 3 * 1024 * 1024) {
    throw new Error(
      "Slide image is too large for processing."
    );
  }

  const messages = [
    {
      role: "system",
      content: `
You are SlideSpark's PowerPoint image reader.

Read only the text actually visible in the image.

Instructions:
1. Extract headings, sentences and bullet points.
2. Preserve original terminology.
3. Preserve numbers and definitions accurately.
4. Do not add explanations.
5. Do not use outside knowledge.
6. Do not guess unreadable text.
7. Preserve the order of visible information.
8. If no readable text exists, return empty text.

Return JSON in this exact format:

{
  "text": "The text visible on the slide"
}
`
    },
    {
      role: "user",
      content: [
        {
          type: "text",
          text:
            "Transcribe the educational content " +
            "visible in this PowerPoint slide."
        },
        {
          type: "image_url",
          image_url: {
            url: image,
            detail: "high"
          }
        }
      ]
    }
  ];

  const result = await askAI(messages, 3000);

  return {
    text: clean(result.text).slice(0, 6000)
  };
}

// =====================================
// VALIDATE GENERATED QUESTIONS
// =====================================

function validateQuestions(items, slides, count) {
  if (!Array.isArray(items)) {
    throw new Error(
      "AI did not return a valid question list."
    );
  }

  if (items.length !== count) {
    throw new Error(
      `AI generated ${items.length} questions ` +
      `instead of ${count}. Please retry.`
    );
  }

  const slideMap = new Map(
    slides.map((slide) => [
      slide.number,
      slide.text
    ])
  );

  const seen = new Set();
  const valid = [];

  for (const item of items) {
    const type = clean(item.type);

    if (!TYPES.includes(type)) {
      throw new Error(
        "AI returned an unsupported question type."
      );
    }

    const slideNumber = Number(item.slideNumber);
    const sourceText = slideMap.get(slideNumber);

    if (!sourceText) {
      throw new Error(
        `Question refers to an unavailable slide: ${slideNumber}.`
      );
    }

    const question = clean(item.question);
    const answer = clean(item.answer);
    const explanation = clean(item.explanation);
    const evidence = clean(item.evidence);

    if (
      !question ||
      !answer ||
      !explanation ||
      !evidence
    ) {
      throw new Error(
        "AI returned an incomplete question."
      );
    }

    // Confirm that the quoted evidence appears
    // in the actual extracted slide text.

    if (
      !normalize(sourceText).includes(
        normalize(evidence)
      )
    ) {
      throw new Error(
        `Evidence for a question was not found ` +
        `on slide ${slideNumber}. Please retry.`
      );
    }

    const signature = normalize(question);

    if (seen.has(signature)) {
      throw new Error(
        "AI generated duplicate questions."
      );
    }

    seen.add(signature);

    const cleaned = {
      type,
      question,
      answer,
      explanation,
      evidence,
      slideNumber
    };

    if (type === "mcq") {
      if (
        !Array.isArray(item.options) ||
        item.options.length !== 4
      ) {
        throw new Error(
          "Multiple-choice question must have four options."
        );
      }

      const options = item.options.map(clean);
      const correctIndex = Number(item.correctIndex);

      if (
        options.some((option) => !option) ||
        new Set(options.map(normalize)).size !== 4 ||
        !Number.isInteger(correctIndex) ||
        correctIndex < 0 ||
        correctIndex > 3
      ) {
        throw new Error(
          "AI returned invalid multiple-choice options."
        );
      }

      if (
        normalize(options[correctIndex]) !==
        normalize(answer)
      ) {
        throw new Error(
          "AI returned an inconsistent correct answer."
        );
      }

      cleaned.options = options;
      cleaned.correctIndex = correctIndex;
    }

    valid.push(cleaned);
  }

  return valid;
}

// =====================================
// MODE 2 - GENERATE QUIZ
// =====================================

async function generateQuiz(body) {
  const count = Number(body.count);
  const instructions =
    clean(body.instructions).slice(0, 1500);

  if (
    !Number.isInteger(count) ||
    count < 1 ||
    count > 20
  ) {
    throw new Error(
      "Select between 1 and 20 questions."
    );
  }

  if (
    !Array.isArray(body.slides) ||
    body.slides.length === 0
  ) {
    throw new Error(
      "No PowerPoint slide content received."
    );
  }

  const slides = body.slides
    .filter((slide) =>
      Number.isInteger(Number(slide.number)) &&
      Number(slide.number) > 0 &&
      typeof slide.text === "string" &&
      slide.text.trim()
    )
    .slice(0, 100)
    .map((slide) => ({
      number: Number(slide.number),
      text: slide.text.slice(0, 6000)
    }));

  if (!slides.length) {
    throw new Error(
      "No readable PowerPoint content was found."
    );
  }

  const slideContent = slides.map((slide) => {
    return (
      `SLIDE ${slide.number}\n` +
      `${slide.text}`
    );
  }).join("\n\n").slice(0, 55000);

  const systemPrompt = `
You are SlideSpark, an educational quiz generator.

You create quizzes strictly from supplied
PowerPoint presentation content.

IMPORTANT:
The presentation text may have been extracted
from slide images using AI Vision.

SOURCE RULES:
1. Use ONLY the supplied PowerPoint content.
2. Do not add information from general knowledge.
3. Never introduce topics absent from the slides.
4. Every question must have a supported answer.
5. Include the correct source slide number.
6. Include a short exact quote from that slide
   in the "evidence" field.
7. The evidence must be consecutive words
   copied from the supplied slide text.
8. Do not fabricate quotations.
9. Do not repeat questions.
10. Use simple, clear educational language.
11. Follow the requested question distribution.
12. Generate exactly the requested number.
13. Return valid JSON only.

QUESTION TYPES:

"mcq":
Choose the correct answer.
Exactly four options and one correct answer.

"one_word":
A question answered using one word.

"short":
A question requiring a short explanation.

"long":
An answer-the-following question requiring
a more detailed explanation.

DEFAULT DISTRIBUTION FOR 10 QUESTIONS:
3 mcq
2 one_word
2 short
3 long

If the teacher requests another distribution,
follow that distribution instead.

For mcq:
correctIndex must be 0, 1, 2 or 3.
answer must match the correct option.

For other types:
options must be an empty array.
correctIndex must be -1.

Return this JSON object:

{
  "questions": [
    {
      "type": "mcq",
      "question": "Question text",
      "options": [
        "Option A",
        "Option B",
        "Option C",
        "Option D"
      ],
      "correctIndex": 0,
      "answer": "Option A",
      "explanation": "Explanation supported by the PPT",
      "slideNumber": 1,
      "evidence": "Exact words from the slide"
    }
  ]
}
`;

  const userPrompt = `
Generate exactly ${count} quiz questions.

TEACHER INSTRUCTIONS:
${instructions || "Create a balanced educational quiz."}

POWERPOINT CONTENT:
${slideContent}

Use only this material.
Return the requested JSON object.
`;

  const result = await askAI(
    [
      {
        role: "system",
        content: systemPrompt
      },
      {
        role: "user",
        content: userPrompt
      }
    ],
    8000
  );

  return {
    questions: validateQuestions(
      result.questions,
      slides,
      count
    )
  };
}

// =====================================
// VERCEL API HANDLER
// =====================================

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");

    return res.status(405).json({
      error: "Only POST requests are supported."
    });
  }

  if (!process.env.OPENAI_API_KEY) {
    return res.status(500).json({
      error:
        "OPENAI_API_KEY is missing. " +
        "Add it in Vercel Environment Variables."
    });
  }

  try {
    const body = req.body || {};

    if (body.mode === "read") {
      const result = await readSlide(body);

      return res.status(200).json(result);
    }

    if (body.mode === "generate") {
      const result = await generateQuiz(body);

      return res.status(200).json(result);
    }

    return res.status(400).json({
      error:
        "Invalid request mode. " +
        "Expected 'read' or 'generate'."
    });

  } catch (error) {
    console.error(
      "SlideSpark API Error:",
      error.message
    );

    return res.status(500).json({
      error:
        error.message ||
        "Unexpected quiz generation error."
    });
  }
}
