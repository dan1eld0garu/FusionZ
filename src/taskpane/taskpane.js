// import fs from "fs";

/* global Word console */

// To avoid duplicates we need to have the range of text (ideally the paragraph) where
// this edit needs to be. For now hardcoding it.

const uniqueStyles = [];

export async function insertText(text) {
  // const sample_response = JSON.parse(fs.readFileSync("./src/taskpane/sample_response_1.json", "utf8"));

  await Word.run(async (context) => {
    let clauseRangeResult = await getClauseRange(context, "Delivery of Battery Packages to Site");
    if (clauseRangeResult.foundRange) {
      clauseRangeResult.range.load("text");
      await context.sync();
      console.log("Clause range text: " + clauseRangeResult.range.text);
    }
    console.log("End of the execution");
  });
}

async function getClauseRange(context, clauseHeaderText) {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items/style/name,items/text");
  await context.sync();

  let startRange = null;
  let endRange = null;

  for (let i = 0; i < paragraphs.items.length; i++) {
    const paragraph = paragraphs.items[i];

    // ============ Debug code ====================
    if (!uniqueStyles.includes(paragraph.style)) {
      uniqueStyles.push(paragraph.style);
    }

    if (paragraph.text.includes(clauseHeaderText) && paragraph.style === "heading 1") {
      startRange = paragraph.getRange();
      continue;
    }

    if (startRange && paragraph.style === "heading 1") {
      endRange = paragraph.getRange();
      break;
    }
  }

  console.log("Is this still executing");

  // ============ Debug code - print styles ====================
  // for (var style in uniqueStyles) {
  //   console.log("Paragraph style:" + uniqueStyles[style]);
  // }
  // for (let i = 0; i < paragraphs.items.length; i++) {
  //   const paragraph = paragraphs.items[i];
  //   if (paragraph.style === "heading 1") {
  //     console.log("Paragraph with style 'heading 1': " + paragraph.text);
  //   }
  // }
  let result = {
    foundRange: false,
    range: null,
  };

  if (startRange && endRange) {
    // console.log("Found range text: " + startRange.text + " " + endRange.text);
    try {
      // Expand startRange to include endRange
      result.foundRange = true;
      result.range = startRange.expandTo(endRange);
      return result;
    } catch (error) {
      console.error("Error during context.sync():", error);
    }
  } else {
    console.log("startRange or endRange is not defined");
    if (!startRange) {
      console.log("startRange is not defined");
    }
    if (!endRange) {
      console.log("endRange is not defined");
    }

    return Promise.resolve(result);
  }
}
