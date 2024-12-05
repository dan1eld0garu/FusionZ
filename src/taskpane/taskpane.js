/* global Word console */

import OpenAI from "openai";

const openai = new OpenAI({
  apiKey: "secret_key_to_be_added",
  dangerouslyAllowBrowser: true,
});

const step1Prompt = `Only return the JSON code with no extra text. Analyse and give suggested changes to this supply contract from the seller's perspective based on the governing law in this contract. Explain why such changes are suggested and show us how exactly we should make the changes.

You are tasked with reviewing and editing a contract document. Your edits must be provided in a structured JSON format to communicate changes clearly. The JSON will support one specific operation:
replace: Propose a replacement for a specific word or phrase and include an explanation for the change.

JSON Structure
The JSON output must follow this structure:
json
Copy code
{
  "edits": [
    {
      "operation": "replace",
      "target": "original_word_or_phrase",
      "replacement": "suggested_word_or_phrase",
      "explanation": "Reason why this change is required."
    }
  ]
}

Instructions
1.	Carefully identify all necessary edits in the given contract text.
2.	Use the replace operation to propose changes where a word or phrase needs correction, clarification, or improvement.
3.	Include a clear and concise explanation for each suggested edit, explaining why the change is necessary.
4.	Ensure each change is explicit—do not summarize edits.

Example
Original Sentence:
"The contractor shall complete the work within 30 days."
Proposed Edits:
•	Replace "shall" with "must" to strengthen the legal obligation.
•	Replace "30" with "thirty (30)" for legal formatting consistency.
Output JSON:
{
  "edits": [
    {
      "operation": "replace",
      "target": "shall",
      "replacement": "must",
      "explanation": "Replace 'shall' with 'must' to strengthen the legal obligation."
    },
    {
      "operation": "replace",
      "target": "30",
      "replacement": "thirty (30)",
      "explanation": "Replace '30' with 'thirty (30)' for consistency with legal formatting standards."
    }
  ]
}
Provide your edits in the specified JSON format. Each edit must include a clear explanation for the proposed change. Only return the JSON code with no extra text.`;

const step2Prompt = `"Only return the JSON code, with no extra text. Extract the full text of all the clauses with any title containing 'Liability' or 'Indemnity', including variations such as 'Limitation of Liability', 'Defects Liability', 'Liability Limitations', 'Limitations on Liability & Indemnity', or similar phrases. Start from the section header (e.g., '9 Limitations on Liability & Indemnity') and include all paragraphs, numbered subsections, and bullet points that follow. Continue extracting the section until the next major numbered section header (e.g., '10 Limited Warranty') or equivalent. Retain the original document formatting, including numbering, indentation, and bullet points. Confirm the last extracted line to ensure completeness. Extract and return the text content of the uploaded Word document exactly as it is, without adding any explanations, commentary, or extra sentences.

Edit the provided Limitation of Liability clause to make it more pro-seller. Use clear, concise, and commercially favorable language to:
Cap the seller's total liability to a defined limit, such as the total fees paid by the buyer
Exclude indirect, consequential, punitive, and special damages.
Ensure liability exclusions apply even if the seller has been advised of potential damages.
Maintain compliance with the relevant jurisdiction (e.g., South Australia).
For each sub-section of the clause, provide explanations for the changes made, focusing on minimizing the seller's financial exposure while ensuring enforceability.

You are tasked with reviewing and editing a contract document. Your edits must be provided in a structured JSON format to communicate changes clearly. The JSON will support one specific operation:
replace: Propose a replacement for a specific word or phrase and include an explanation for the change.

JSON Structure
The JSON output must follow this structure:
json
Copy code
{
  ""edits"": [
    {
      ""operation"": ""replace"",
      ""target"": ""original_word_or_phrase"",
      ""replacement"": ""suggested_word_or_phrase"",
      ""explanation"": ""Reason why this change is required.""
    }
  ]
}

Instructions
1.	Carefully identify all necessary edits in the given contract text.
2.	Use the replace operation to propose changes where a word or phrase needs correction, clarification, or improvement.
3.	Include a clear and concise explanation for each suggested edit, explaining why the change is necessary.
4.	Ensure each change is explicit—do not summarize edits.

Example
Original Sentence:
""The contractor shall complete the work within 30 days.""
Proposed Edits:
•	Replace ""shall"" with ""must"" to strengthen the legal obligation.
•	Replace ""30"" with ""thirty (30)"" for legal formatting consistency.
Output JSON:
{
  ""edits"": [
    {
      ""operation"": ""replace"",
      ""target"": ""shall"",
      ""replacement"": ""must"",
      ""explanation"": ""Replace 'shall' with 'must' to strengthen the legal obligation.""
    },
    {
      ""operation"": ""replace"",
      ""target"": ""30"",
      ""replacement"": ""thirty (30)"",
      ""explanation"": ""Replace '30' with 'thirty (30)' for consistency with legal formatting standards.""
    }
  ]
}
Provide your edits in the specified JSON format. Each edit must include a clear explanation for the proposed change. Only return the JSON code, with no extra text."`;

async function getEditsFromOpenAI(documentText, systemPrompt) {
  const wordLimit = 100;
  const words = documentText.split(" ");
  const limitedText = words.slice(0, wordLimit).join(" ");
  console.log("Debugging some text" + limitedText);

  // const systemPrompt = `The following text is a legal contract and I need to extract all the clauses. Return a JSON array with
  // the following format: [{clauseStart: "", clauseEnd: "", clauseTitle: ""}] containing a couple of original words for each clause start and end that will help me extract the
  // text for each clause`;
  const promptContent = systemPrompt + "\n\n Contract text:" + documentText;
  let response;
  try {
    response = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages: [
        {
          role: "user",
          content: [
            {
              type: "text",
              text: promptContent,
            },
          ],
        },
      ],
      temperature: 1,
      max_tokens: 2048,
      top_p: 1,
      frequency_penalty: 0,
      presence_penalty: 0,
      response_format: {
        type: "text",
      },
    });
    // response = "test response";
  } catch (e) {
    console.log(e);
  }
  if (response) {
    if (response.choices && response.choices.length > 0) {
      console.log("Response received from OpenAI: " + response.choices[0].message.content);
      return response.choices[0].message.content;
    } else {
      console.log("No valid choices received from OpenAI");
      return "";
    }
  } else {
    console.log("No response");
    return "";
  }
}

const DEBUG = false;

const sample_response = `Response received from OpenAI: Below are recommended changes in a structured JSON format for the supply contract from the seller's perspective. Each change is described in terms of the operation (replacement), the target phrase, the suggested replacement, and the reason for the change.
Suggested Changes in JSON Format: 
{ 
  "edits": [ 
    { 
      "operation": "replace", 
      "target": "the Principal must pay the Battery Supplier the Contract Price in accordance with this Contract.", 
      "replacement": "the Principal is obligated to pay the Battery Supplier the Contract Price in strict compliance with the terms of this Contract.", 
      "explanation": "Rewording emphasizes the need for strict compliance, which strengthens the seller's position in case of disputes about payment." 
    }, 
    { 
      "operation": "replace", 
      "target": "No reliance will be placed by either party on any representations, promise or other inducement made or given or alleged to be made or given by one party to the other party prior to the Execution Date.", 
      "replacement": "Both parties acknowledge that no reliance is placed on any representations, promises, or inducements made prior to the Execution Date.", 
      "explanation": "Changing ‘No reliance will be placed’ to “Both parties acknowledge” conveys mutual understanding and agreement, which provides better legal protection." 
    }, 
    { 
      "operation": "replace", 
      "target": "its obligations under the Contract are valid and binding and are enforceable against it in accordance with its respective terms subject to any necessary stamping and registration, the availability of equitable remedies and Laws relating to the enforcement of creditor’s rights.", 
      "replacement": "its obligations under the Contract are valid, binding, and enforceable in accordance with the terms, subject only to applicable laws governing contractual compliance and enforcement.", 
      "explanation": "This adjustment avoids overly complex language and clarifies that enforceability is limited to existing laws rather than being subject to undefined conditions." 
    }, 
    { 
      "operation": "replace", 
      "target": "The Battery Supplier must comply with, and ensure that its Subcontractors and all other persons engaged in the performance of the Works comply with: the WHS Law; and the WHS Management Plan.", 
      "replacement": "The Battery Supplier must ensure compliance with WHS Law and promote adherence to the WHS Management Plan among its Subcontractors and all personnel involved in the Works.", 
      "explanation": "This rephrasing emphasizes the responsibility of the Battery Supplier to instill compliance throughout its teams rather than merely ensuring compliance." 
    }, 
    { 
      "operation": "replace", 
      "target": "The Battery Supplier acknowledges and agrees that nothing in this clause 17 prejudices any other right which the Principal may have against the Battery Supplier arising out of the failure of the Battery Supplier to provide Equipment or the Works in accordance with this Contract.", 
      "replacement": "The Battery Supplier acknowledges that this clause 17 does not limit the Principal’s rights to claim for failures related to Equipment or Works not delivered in compliance with this Contract.", 
      "explanation": "This revision clarifies the seller's acknowledgment while reinforcing the contractual rights of the Principal without ambiguity." 
    }, 
    { 
      "operation": "replace", 
      "target": "The Principal may at any time direct the Battery Supplier to carry out a Variation by issuing a Variation Order pursuant to clause 18.1.", 
      "replacement": "The Principal retains the authority to mandate Variations by issuing a Variation Order under clause 18.1, provided such obligations are within the agreed scope of work.", 
      "explanation": "Emphasizing ‘within the agreed scope of work’ ensures that the Principal does not exercise this power in an arbitrary manner while still holding authority." 
    }, 
    { 
      "operation": "replace", 
      "target": "the Battery Supplier acknowledges and agrees that the submission of the Battery Supplier's Program or an updated Battery Supplier’s Program does not relieve the Battery Supplier of its obligations to achieve Practical Completion.", 
      "replacement": "The Battery Supplier acknowledges that while submitting the Battery Supplier's Program, it remains solely responsible for achieving Practical Completion as per the Contract stipulations.", 
      "explanation": "This change asserts the Battery Supplier's continuous accountability towards meeting completion deadlines despite submitting schedules." 
    }, 
    { 
      "operation": "replace", 
      "target": "The Battery Supplier must indemnify and hold harmless the Principal against all Loss which the Principal suffers or incurs, or may suffer or incur, arising out of, or in connection with a breach of clause 30.1(a).", 
      "replacement": "The Battery Supplier must indemnify the Principal fully for any Loss incurred due to breaches of clause 30.1(a), ensuring that comprehensive coverage is in place.", 
      "explanation": "This emphasizes the extent of indemnification, reinforcing the Battery Supplier's responsibility for losses incurred by the Principal due to non-compliance." 
    }, 
    { 
      "operation": "replace", 
      "target": "In the event of loss or damage to the Equipment being caused by a combination of Excepted Risks and other risks, any such direction and consequential valuation made under clause 18.8 shall take into account the proportional responsibility of the Battery Supplier.",  
      "replacement": "If losses occur due to a combination of Excepted Risks and other risks, any directions and consequential valuations made under clause 18.8 shall reflect the Battery Supplier's responsibility proportionately.", 
      "explanation": "Clarifying that valuations should proportionally reflect responsibility ensures fairness in obligations assigned, thus protecting the Battery Supplier’s interests." 
    }, 
    { 
      "operation": "replace", 
      "target": "the aggregate liability of the Battery Supplier to the Principal arising out of or in connection with this Contract will in no event exceed an amount equal to the amount set out in Item 32 of Annexure Part A (Liability Limitation).", 
      "replacement": "The total liability of the Battery Supplier to the Principal for claims related to this Contract shall not exceed the amount specified in Item 32 of Annexure Part A, barring exceptions outlined in this Contract.", 
      "explanation": "This modification maintains clarity while reinforcing the limitations of liability, ensuring that the scope of liability is understood by both parties." 
    }, 
    { 
      "operation": "replace", 
      "target": "Payments made by the Principal to any Battery Supplier Responsible Party for any Equipment delivered or work executed for the purposes of this Contract may be deducted or paid from the proceeds of any Security.", 
      "replacement": "The Principal may offset payments made to any Battery Supplier Responsible Party against the proceeds of Security, ensuring that obligations are met efficiently without undue delay.", 
      "explanation": "Clarifying that payments can be offset stresses the efficiency of payment handling while protecting the Principal's interests." 
    } 
  ] 
} 
  

Summary of Suggested Changes: 

Clarity and Conciseness: Many replacements focus on making the language clearer and more concise to reduce ambiguity. 

Responsibility Emphasis: Several changes emphasize the Battery Supplier's ongoing responsibility for performance obligations. 

Protection of Rights: Modifications are made to better safeguard the interests of the Battery Supplier. 

** contract compliance:** Adjustments made in terms of contract compliance highlight the need for mutual awareness and acknowledgments without displacing obligations. 

Data Protection: Changes relating to indemnification and liability enhance the seller's ability to limit exposure while ensuring performance obligations are adequately addressed. 

These adjustments can help create a more balanced contract that better protects the seller’s interests while ensuring compliance with legal frameworks and obligations.`;

const uniqueStyles = [];

export async function deleteAllComments() {
  await Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();

    for (let i = 0; i < comments.items.length; i++) {
      comments.items[i].delete();
    }

    await context.sync();
    console.log("Deleted all comments.");
  });
}

export async function clearEdits() {
  await Word.run(async (context) => {
    const body = context.document.body;
    const trackedChanges = body.getTrackedChanges();
    trackedChanges.rejectAll();
    console.log("Rejected all tracked changes.");
    await deleteAllComments();
  });
}

function extractAndParseJSON(text) {
  const startIndex = text.indexOf("{");
  const endIndex = text.lastIndexOf("}");
  if (startIndex === -1 || endIndex === -1 || startIndex >= endIndex) {
    throw new Error("Invalid JSON string");
  }
  const jsonString = text.substring(startIndex, endIndex + 1);
  return JSON.parse(jsonString);
}

async function applyEdits(context, clauseRangeResult = null, editsResult) {
  context.document.body.trackRevisions = true;
  context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;

  const sample_response_parsed = extractAndParseJSON(sample_response);
  console.log("Parsed JSON" + sample_response_parsed);

  // const jsonString = JSON.stringify(sample_response_parsed, null, 2);
  //console.log("JSON Stringified Object: " + jsonString);

  const allSearchResults = [];

  for (const edit in sample_response_parsed.edits) {
    const editContent = sample_response_parsed.edits[edit];
    if (editContent.operation === "replace") {
      const searchResults = context.document.body.search(editContent.target);

      searchResults.load("items");
      console.log("Pushing search results for:", editContent.target);
      allSearchResults.push({ searchResults, editContent });
    }
  }

  console.log("Search results: " + JSON.stringify(allSearchResults));

  //   for (const subClause in sample_response[clause]) {
  //     for (const subSubClause in sample_response[clause][subClause]) {
  //       const actions = sample_response[clause][subClause][subSubClause];
  //       for (const actionKey in actions) {
  //         const action = actions[actionKey];
  //         if (action.action === "replace") {
  //           for (const change of action.changes) {
  //             if (!change.replace || typeof change.replace !== "string" || change.replace.trim() === "") {
  //               console.error("Invalid search term:", change.replace);
  //               continue;
  //             }

  //             const searchResults = clauseRangeResult.range.search(change.replace, {
  //               matchCase: true,
  //               matchWholeWord: true,
  //             });

  //             searchResults.load("items");
  //             console.log("Pushing search results for:", change.replace);
  //             allSearchResults.push({ searchResults, change });
  //           }
  //         }
  //       }
  //     }
  //   }
  // }

  try {
    await context.sync();
  } catch (e) {
    console.error("Error during context.sync():", e);
    if (e.debugInfo) {
      console.error("Debug info:", e.debugInfo);
    }
  }

  // for (const result of allSearchResults) {
  //   const { searchResults, change } = result;
  //   if (searchResults.items.length > 0) {
  //     console.log("We have some search result" + searchResults.items[0]);
  //     for (let i = 0; i < searchResults.items.length; i++) {
  //       searchResults.items[i].insertText(change.with, Word.InsertLocation.replace);
  //       if (change.explanation) {
  //         searchResults.items[i].insertComment(change.explanation);
  //       }
  //     }
  //   } else {
  //     console.log("No result");
  //   }
  // }

  // const jsonString = JSON.stringify(allSearchResults, null, 2);
  // console.log("JSON Stringified Objectt: " + jsonString);

  // // {
  // //   "searchResults": {
  // //     "items": [
  // //       {
  // //         "hyperlink": null,
  // //         "isEmpty": false,
  // //         "style": "heading 3",
  // //         "styleBuiltIn": "Heading3",
  // //         "text": "The Principal must pay the Battery Supplier the Contract Price in accordance with this Contract."
  // //       }
  // //     ]
  // //   },
  // //   "editContent": {
  // //     "operation": "replace",
  // //     "target": "the Principal must pay the Battery Supplier the Contract Price in accordance with this Contract.",
  // //     "replacement": "the Principal is obligated to pay the Battery Supplier the Contract Price in strict compliance with the terms of this Contract.",
  // //     "explanation": "Rewording emphasizes the need for strict compliance, which strengthens the seller's position in case of disputes about payment."
  // //   }
  // // }

  for (const result of allSearchResults) {
    const { searchResults, editContent } = result;
    if (searchResults.items.length > 0) {
      // console.log("We have some search result" + searchResults.items[0]);
      const jsonString = JSON.stringify(searchResults.items[0].text, null, 2);
      console.log("JSON Stringified Object: " + jsonString);

      for (let i = 0; i < searchResults.items.length; i++) {
        // console.log("JSON Stringified Object: " + JSON.stringify(editContent, null, 2));
        searchResults.items[i].insertText(editContent.replacement, Word.InsertLocation.replace);
        if (editContent.explanation) {
          searchResults.items[i].insertComment(editContent.explanation);
        }
      }
    } else {
      console.log("No result");
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
}

export async function processDocument() {
  // const sample_response = JSON.parse(fs.readFileSync("./src/taskpane/sample_response_1.json", "utf8"));

  await Word.run(async (context) => {
    // 1) Extract the clause range for improved search results
    // let clauseRangeResult = await getClauseRange(context, "Delivery of Battery Packages to Site");
    // if (clauseRangeResult.foundRange) {
    //   clauseRangeResult.range.load("text");
    //   await context.sync();
    //   if (DEBUG) {
    //     console.log("Clause range text: " + clauseRangeResult.range.text);
    //   }
    // }

    // Get the edits with step 1 prompt
    const body = context.document.body;
    body.load("text");
    await context.sync();
    // let step1Edits = getEditsFromOpenAI(body.text, step1Prompt);
    // let step2Edits = getEditsFromOpenAI(body.text, step2Prompt);
    let step2Edits = "";

    if (DEBUG) {
      console.log("Clause range text: " + step2Edits);
    }

    // 2) Apply edits
    await applyEdits(context, null, step2Edits);

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
