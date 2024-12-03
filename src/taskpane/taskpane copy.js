/* global Word console */

export async function insertText(text) {
  console.log("Inserting text: " + text);

  // We need to identify the strcuture of the document clauses. Are they using headers?
  // Sub-clauses are using sub-headers? The replace algorithm will depend on the structure.
  // As a simplification, for now assume that the structure is header -> sub-header -> sub-sub-header.

  // I need to have a list of all the headers so that I can restrict the search range for a specific clause
  // and have more accurate results.
  // I can extract all this using the AI model, but for now, let's assume that we have a list of headers.

  // Extract the clause headers from the document, send that to the AI model to get the changes in a JSON format.
  // and then apply the changes to the document.

  // Search for the paragraph, send it to OpenAI, get the changes, and apply them back to the paragraph.

  const sampleResponse = {
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

  

  const clauseSampleText = `Delivery of Battery Packages to Site and Urgent Protection

  Notice to Ship and Delivery of Battery Packages to Site

  The Battery Supplier must notify the Principal in writing when it is ready to ship the Battery Packages from the Place of Manufacture and Storage to Site (Notice of Readiness to Ship).

  Without limiting the obligation of the Battery Supplier to achieve Practical Completion by the Date for Practical Completion, the Notice of Readiness to Ship must be given no later than 5 Business Days prior to the relevant Date for Shipment specified for each Battery Package.

  Following the receipt by the Principal of a Notice of Readiness to Ship, the Principal will notify the Battery Supplier in writing that it may proceed to ship the relevant Battery Package from the Place of Manufacture and Storage to Site by issuing a Notice to Ship with respect to the relevant Battery Package.

  Without limiting the obligation of the Battery Supplier to achieve Practical Completion by the Date for Practical Completion, following receipt by the Battery Supplier of a Notice to Ship, the Battery Supplier must commence shipping the relevant Battery Package from the Place of Manufacture and Storage to the Site.

  The Battery Supplier acknowledges and agrees:

  the Battery Supplier must not commence shipping a Battery Package specified in clause 24.1(a) from the Place of Manufacture and Storage to the Site until the Principal has issued a Notice to Ship with respect to the relevant Battery Package;

  without limiting clause 24.1(e)(iii), if the Notice of Readiness to Ship is given by the Battery Supplier, or the Battery Package is otherwise ready to ship, on a date which is earlier than 5 Business Days prior to the Date for Shipment for that Battery Package, the Principal is not required to issue a Notice to Ship earlier than the Date for Shipment for that Battery Package and the Battery Supplier is required, at the Battery Supplier’s cost, to arrange for temporary storage of the Battery Package at the Place of Manufacture and Storage or another appropriate facility until the Notice to Ship is issued by the Principal with respect to that Battery Package and the Battery Supplier has commenced shipping the relevant Battery Package; and

  the Principal is not required to issue a Notice to Ship in respect of any Battery Package on or before the Date for Shipment for that Battery Package and may issue a Notice to Ship at any time (including before it has received a Notice of Readiness to Ship from the Battery Supplier), however the issue by the Principal of a Notice to Ship after the Date for Shipment for that Battery Package may be grounds for an extension of time under clause 19.3.

  If the Battery Supplier commences shipment of the relevant Battery Package prior to receiving a Notice to Ship from the Principal and the D&C Contractor is not ready to receive and install the Battery Packages at Site by the time that the relevant Battery Package arrives at the Site then:

  the costs of transporting the relevant Battery Package from the Site to a temporary storage facility (including the costs of loading the Battery Package at the Site and unloading the Battery Package at the Temporary Storage Facility);

  the costs of storing the relevant Battery Package at the Temporary Storage Facility; and

  the costs of transporting the relevant Battery Package from the Temporary Storage Facility to Site (including the costs of loading the Battery Package at the Temporary Storage Facility and unloading the Battery Package at Site),

  will be borne by the Battery Supplier.

  Where the costs of transportation and storage are to be borne by the Battery Supplier, the amount paid by the Principal to third parties for such transportation and storage will be a debt due and payable from the Battery Supplier to the Principal.

  Urgent protection

  If urgent action is necessary to protect the Equipment, Facility, other property or people and the Battery Supplier fails to take the action, in addition to any other remedies of the Principal, the Principal may take the necessary action. If the action was action which the Battery Supplier should have taken at the Battery Supplier’s cost, the reasonable cost incurred by the Principal in the circumstances will be a debt due and payable immediately from the Battery Supplier to the Principal.

  If time permits, the Principal will give the Battery Supplier prior written notice of the intention to take action pursuant to this clause 24.2.`;

  async function getDocumentStructure(context) {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items/style,items/text,items/level");

    await context.sync();

    console.log("Paragraphs:", paragraphs.items);
    const headingStyles = ["Heading 1", "Heading 2", "Heading 3"];

    // Build the structure of the document
    const structure = [];
    paragraphs.items.forEach((paragraph) => {
      if (headingStyles.includes(paragraph.style)) {
        structure.push({
          level: headingStyles.indexOf(paragraph.style) + 1,
          text: paragraph.text.trim(),
          paragraph: paragraph,
        });
      }
    });

    return structure;
  }

  // async function getClauseContent(context, clauseHeading, nextClauseHeading) {
  //   const clauseRange = clauseHeading.paragraph.getRange("After");
  //   let clauseContentRange;

  //   if (nextClauseHeading) {
  //     // Expand to the next clause heading
  //     clauseContentRange = clauseRange.expandTo(nextClauseHeading.paragraph);
  //   } else {
  //     // Expand to the end of the document
  //     clauseContentRange = clauseRange.expandTo(context.document.body.getRange("End"));
  //   }

  //   clauseContentRange.load("text");
  //   await context.sync();

  //   return clauseContentRange.text;
  // }

  async function getEditsFromOpenAI() {
    // Replace with your server endpoint that handles OpenAI API requests
    // const apiUrl = "/api/get-edits";

    // const response = await fetch(apiUrl, {
    //   method: "POST",
    //   headers: {
    //     "Content-Type": "application/json",
    //   },
    //   body: JSON.stringify({ clauseText }),
    // });

    // const data = await response.json();

    // return data.edits;
    // console.log(clauseText);
    return sampleResponse;
  }

// What should this logic do? What is the easiest way to do this?
// If the document is hierarchical, I just need a visitor style algorithm
// that goes through the document and applies the changes to the clauses
// The range of search will always be the start of the clause and the start of the next clause.
// ultimatedly you get to one paragraph. the range is always the start of the next paragraph.
// Some of the clauses won't have edits, so I can just skip them.
// If I have the strcture of the document and a way to find where to start to apply the edits.

// A list of clauses and subclauses and a list of edits. What is the easiest way to apply
// the edits to the clauses?

// Should you iterate through 
// the edits or through the document strcuture?
// How to easlily
// I'm going to represent the document as 

// I can have a list with all the paragraphs.
// How do I applu a replace to a paragraph?
// Build it using a bottom - up approach

// The edits are always going to be inside a paragraph,
// then search inside the paragraph for the text to replace.

// To have no mistakes, find the exact paragraph.
// The document structure could be a map with the path to the paragraph
// and the paragraph itself.

// The edits could then have this path to the paragraph and the edits for it
/// Cause/sub-clause/sub-sub-clause -> edits.

// It's the paragraph range and the way to uniquly identify it is through
// it's path.

// If I solve a clause, I can solve for all the other clauses.
// What paragraph is this? What is the path to this paragraph?

  async function applyEditsToClause(context, clauseContentRange, edits) {
    const searchPromises = [];
    for (const [, subClauseEdits] of Object.entries(edits)) {
      for (const [, subClauseChanges] of Object.entries(subClauseEdits)) {
        for (const [, subClauseSubChanges] of Object.entries(subClauseChanges)) {
          for (const change of subClauseSubChanges.changes) {
            if (change.action === "replace") {
              const searchResults = clauseContentRange.search(change.replace, {
                matchCase: false,
                matchWholeWord: false,
              });
              searchResults.load("items");
              searchPromises.push(searchResults);
            } else if (change.action === "add") {
              // if (change.position === "end of sub-clause") {
              //   clauseContentRange.insertText(change.add, Word.InsertLocation.end);
              // } else if (change.after) {
              //   const searchResults = clauseContentRange.search(change.after, {
              //     matchCase: false,
              //     matchWholeWord: false,
              //   });
              //   searchResults.load("items");
              //   searchPromises.push(searchResults);
              // }
            }
          }
        }
      }
    }

    await context.sync();

    for (const searchResults of searchPromises) {
      if (searchResults.items.length > 0) {
        const change = searchResults.items[0];
        if (change.action === "replace") {
          searchResults.items[0].insertText(change.with, Word.InsertLocation.replace);
        } else if (change.action === "add" && change.after) {
          searchResults.items[0].insertText(change.add, Word.InsertLocation.after);
        }
      }
    }

    await context.sync();
  }

  // Main function
  // I need to get the text for all the paragraphs
  await Word.run(async (context) => {
    // Step 1: Get document structure
    const structure = await getDocumentStructure(context);
    console.log("Document structure:", structure);

    // Step 2: Iterate through clauses
    for (let i = 0; i < structure.length; i++) {
      const clauseHeading = structure[i];
      const nextClauseHeading = structure[i + 1];

      // Only process top-level clauses (e.g., Heading 1)
      if (clauseHeading.level === 1) {
        // Step 3: Get clause content
        const clauseRange = clauseHeading.paragraph.getRange("After");
        let clauseContentRange;

        if (nextClauseHeading) {
          clauseContentRange = clauseRange.expandTo(nextClauseHeading.paragraph);
        } else {
          clauseContentRange = clauseRange.expandTo(context.document.body.getRange("End"));
        }

        clauseContentRange.load("text");
        await context.sync();

        const clauseText = clauseContentRange.text;

        // Step 4: Send content to OpenAI API
        const edits = await getEditsFromOpenAI(clauseText);

        // Step 5: Apply edits back to the document
        await applyEditsToClause(context, clauseContentRange, edits);
      }
    }

    await context.sync();
    console.log("Document processing complete.");
  });

  // Main function
  await Word.run(async (context) => {
    // Step 1: Get document structure
    const structure = await getDocumentStructure(context);
    console.log("Document structure test:", structure);

    const edits = await getEditsFromOpenAI(clauseSampleText);

    // Step 5: Apply edits back to the document
    const clauseRange = clauseHeading.paragraph.getRange("After");
    let clauseContentRange;

    if (nextClauseHeading) {
      clauseContentRange = clauseRange.expandTo(nextClauseHeading.paragraph);
    } else {
      clauseContentRange = clauseRange.expandTo(context.document.body.getRange("End"));
    }

    await applyEditsToClause(context, clauseContentRange, edits);

    // // Step 2: Iterate through clauses
    // for (let i = 0; i < structure.length; i++) {
    //   const clauseHeading = structure[i];
    //   const nextClauseHeading = structure[i + 1];

    //   // Only process top-level clauses (e.g., Heading 1)
    //   if (clauseHeading.level === 1) {
    //     // Step 3: Get clause content
    //     const clauseText = clauseText;

    //     // Step 4: Send content to OpenAI API
    //     const edits = await getEditsFromOpenAI(clauseSampleText);

    //     // Step 5: Apply edits back to the document
    //     const clauseRange = clauseHeading.paragraph.getRange("After");
    //     let clauseContentRange;

    //     if (nextClauseHeading) {
    //       clauseContentRange = clauseRange.expandTo(nextClauseHeading.paragraph);
    //     } else {
    //       clauseContentRange = clauseRange.expandTo(context.document.body.getRange("End"));
    //     }

    //     await applyEditsToClause(context, clauseContentRange, edits);
    //   }
    // }

    await context.sync();
    console.log("Document processing complete.");
  });
}

// Step 1: Get document structure - a generic way to get the clauses from all types of constracts with custom headers.
// Step 2: Iterate through clauses
// Step 3: Get clause content
// - the guy that compares word documents.

// Step 4: Send content to OpenAI API and get a JSON with the edits - AI consultants, legal engineers, etc.
// Risk.

// Step 5: Apply edits back to the document.
// - the guy that compares word documents.

// Step 5 with an hardcoded example edits response
// Send a hardcoded clause to OpenAI (here I would need the prompt to get the edits - Snowy)
// Check that the JSON we get can be applied with what I did in step 5 - Snowy
// Automate this process for all the clauses in the document
// Figure out a way to process any document with any structure.

// Step 6: Profit! - Snowy and Daniel.
