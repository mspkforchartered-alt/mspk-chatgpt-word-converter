const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const { exec } = require("child_process");

const app = express();

app.use(cors());
app.use(express.json());

/* =========================================
   🔥 SMART TABLE CONVERTER
========================================= */
function convertTables(text) {

  const lines = text.split("\n");

  let result = [];

  for (let line of lines) {

    // =====================================
    // ✅ Already markdown table
    // Keep it untouched
    // =====================================
    if (
      line.includes("|") &&
      line.split("|").length >= 3
    ) {

      // Clean repeated ||
      line = line.replace(/\|\|+/g, "|");

      // Remove leading/trailing extra |
      line = line.replace(/^\|+/, "|");
      line = line.replace(/\|+$/, "|");

      result.push(line);

      continue;
    }

    // =====================================
    // ✅ Space separated table rows
    // =====================================
    if (/\s{2,}/.test(line.trim())) {

      const cols =
        line.trim().split(/\s{2,}/);

      if (cols.length >= 2) {

        result.push(
          "| " + cols.join(" | ") + " |"
        );

        continue;
      }
    }

    result.push(line);
  }

  return result.join("\n");
}

/* =========================================
   🔥 SMART FORMATTER
========================================= */
function formatText(text) {

  const lines = text.split("\n");

  return lines.map(line => {

    let trimmed = line.trim();

    if (!trimmed) return line;

    // Skip markdown structures
    if (
      trimmed.startsWith("#") ||
      trimmed.startsWith("|") ||
      trimmed.startsWith("-") ||
      trimmed.startsWith("*")
    ) {
      return line;
    }

    // Skip already formatted equations
    if (trimmed.includes("$$")) {
      return line;
    }

    // =========================================
    // ✅ ONLY CONVERT REAL DIVISIONS
    // Avoid decimals like 7.76
    // =========================================
    trimmed = trimmed.replace(
      /(\([^\)]+\)|[A-Za-z]+|\d+)\s*\/\s*(\([^\)]+\)|[A-Za-z]+|\d+)/g,
      "\\frac{$1}{$2}"
    );

    // =========================================
    // ✅ Finance variable fixes
    // =========================================
    trimmed = trimmed.replace(/\bKp\b/g, "K_p");
    trimmed = trimmed.replace(/\bKe\b/g, "K_e");
    trimmed = trimmed.replace(/\bKd\b/g, "K_d");

    // =========================================
    // ✅ Multiplication
    // =========================================
    trimmed = trimmed.replace(/\btimes\b/g, "\\times");

    // =========================================
    // 🔥 STRICT EQUATION DETECTION
    // =========================================
    const isRealEquation =
      /\\frac/.test(trimmed) ||
      /\^/.test(trimmed) ||
      /[A-Za-z]\s*=.*[+\-*/]/.test(trimmed);

    const isSimpleAssignment =
      /^[A-Za-z0-9]+\s*=/.test(trimmed) &&
      !/[+\-*/^]/.test(trimmed);

    if (
      isRealEquation &&
      !isSimpleAssignment &&
      trimmed.length < 150
    ) {
      return `$$ ${trimmed} $$`;
    }

    return line;

  }).join("\n");
}

/* =========================================
   🔥 CLEAN BRACKETS
========================================= */
function cleanBrackets(text) {

  return text
    .replace(/^\s*\[\s*$/gm, "")
    .replace(/^\s*\]\s*$/gm, "")
    .replace(/\[(.*?)\]/g, "$1")
    .replace(/\[\d+\]/g, "");
}

/* =========================================
   🚀 API
========================================= */
app.post("/generate-docx", (req, res) => {

  let text = req.body.text;

  const inputPath =
    path.join(__dirname, "input.md");

  const outputPath =
    path.join(__dirname, "output.docx");

  try {

    // =====================================
    // 🔥 TABLE CONVERSION
    // =====================================
    let formattedText =
      convertTables(text);

    // =====================================
    // 🔥 EQUATION FORMATTING
    // =====================================
    formattedText =
      formatText(formattedText);

    // =====================================
    // 🔥 CLEANUP
    // =====================================
    formattedText =
      cleanBrackets(formattedText);

    fs.writeFileSync(
      inputPath,
      formattedText
    );

    const command =
      `pandoc "${inputPath}" -o "${outputPath}" --from markdown --standalone`;

    exec(command, (err) => {

      if (err) {

        console.error(
          "Pandoc Error:",
          err
        );

        return res
          .status(500)
          .send("❌ Conversion failed");
      }

      res.download(
        outputPath,
        "MSPK.docx",
        () => {

          try {

            fs.unlinkSync(inputPath);
            fs.unlinkSync(outputPath);

          } catch {}

        }
      );

    });

  } catch (error) {

    console.error(
      "Server Error:",
      error
    );

    res
      .status(500)
      .send("❌ Server error");
  }
});

/* =========================================
   🚀 START SERVER
========================================= */
app.listen(3000, () => {

  console.log(
    "🚀 MSPK Backend running on port 3000"
  );

});