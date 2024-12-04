/* global Word console */

const sample_response = {
  "Clause 24": {
    "Delivery of Battery Packages to Site and Urgent Protection": {
      "24.1 Notice to Ship and Delivery of Battery Packages to Site": {
        "(b)": {
          action: "replace",
          changes: [
            {
              replace: "no later than 5 Business Days",
              with: "no later than 5 Business Days or a mutually agreed-upon timeframe",
            },
          ],
        },
        "(d)": {
          action: "replace",
          changes: [
            {
              replace: "commence shipping the relevant Battery Package",
              with: "commence shipping the relevant Battery Package within a reasonable time",
            },
            {
              add: "subject to any unforeseen logistical or transportation challenges",
              after: "within a reasonable time",
            },
          ],
        },
        "(e)(ii)": {
          action: "add",
          changes: [
            {
              add: "If the Principal delays issuing the Notice to Ship beyond the agreed Date for Shipment, any additional storage costs or logistics incurred by the Battery Supplier shall be compensated by the Principal, or the Battery Supplier may claim an extension of time in accordance with clause 19.3.",
              position: "end of sub-clause",
            },
          ],
        },
        "(f)": {
          action: "replace",
          changes: [
            {
              replace: "will be borne by the Battery Supplier",
              with: "will be borne by the Principal, except in cases where the Battery Supplier has failed to follow the shipping instructions and timelines as set forth by the Principal in the Notice to Ship",
            },
          ],
        },
        "(g)": {
          action: "add",
          changes: [
            {
              add: "provided that such costs were incurred due to delays directly caused by the Battery Supplier’s failure to comply with the agreed shipping timelines",
              after: "a debt due and payable from the Battery Supplier to the Principal",
            },
          ],
        },
      },
      "24.2 Urgent Protection": {
        "(a)": {
          action: "replace",
          changes: [
            {
              replace: "fails to take the action",
              with: "fails to take such action despite receiving written notice from the Principal with reasonable time to respond",
            },
            {
              add: "If the action was action which the Battery Supplier should have taken at the Battery Supplier’s cost, the reasonable cost incurred by the Principal in the circumstances will be a debt due and payable immediately from the Battery Supplier to the Principal.",
              position: "end of sub-clause",
            },
          ],
        },
        "(b)": {
          action: "add",
          changes: [
            {
              add: "If the Principal takes action without prior notice due to a genuine emergency, the Battery Supplier shall only be liable for reasonable costs directly related to the protection of the Equipment.",
              position: "end of sub-clause",
            },
          ],
        },
      },
    },
  },
};
// this edit needs to be. For now hardcoding it.

const uniqueStyles = [];

export async function insertText(text) {
  // const sample_response = JSON.parse(fs.readFileSync("./src/taskpane/sample_response_1.json", "utf8"));

  await Word.run(async (context) => {
    // 1) Extract the clause range for improved search results
    let clauseRangeResult = await getClauseRange(context, "Delivery of Battery Packages to Site");
    if (clauseRangeResult.foundRange) {
      clauseRangeResult.range.load("text");
      await context.sync();
      console.log("Clause range text: " + clauseRangeResult.range.text);
    }

    // 2) Search in range and apply edits. Enable track changes
    context.document.body.trackRevisions = true;
    const allSearchResults = [];

    for (const clause in sample_response) {
      for (const subClause in sample_response[clause]) {
        for (const subSubClause in sample_response[clause][subClause]) {
          const actions = sample_response[clause][subClause][subSubClause];
          for (const actionKey in actions) {
            const action = actions[actionKey];
            if (action.action === "replace") {
              for (const change of action.changes) {
                const searchResults = clauseRangeResult.range.search(change.replace, {
                  matchCase: true,
                  matchWholeWord: true,
                });
                searchResults.load("items");
                console.log("pushing" + searchResults + " " + change);
                allSearchResults.push({ searchResults, change });
              }
            }
          }
        }
      }
    }
    try {
      await context.sync();
    } catch (e) {
      console.error("Error during context.sync():", e);
      if (e.debugInfo) {
        console.error("Debug info:", e.debugInfo);
      }
    }

    console.log("After sync");

    for (const result of allSearchResults) {
      const { searchResults } = result;
      if (searchResults.items.length > 0) {
        console.log("We have some search result" + searchResults[0]);
        // for (let i = 0; i < searchResults.items.length; i++) {
        //   searchResults.items[i].insertText(change.with, Word.InsertLocation.replace);
        // }
      } else {
        console.log("No result");
      }
    }

    await context.sync();

    // context.document.body.trackRevisions = false;

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
