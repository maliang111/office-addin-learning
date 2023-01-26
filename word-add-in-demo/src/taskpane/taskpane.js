/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
import {base64Image} from '../base64Image'

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

  // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = insertParagraph;

    document.getElementById("apply-style").onclick = applyStyle;

    document.getElementById("change-font").onclick = changeFont;
    document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;

    document.getElementById("insert-text-outside-range").onclick = insertTextBeforeRange;
    document.getElementById("replace-text").onclick = replaceText;

    document.getElementById("insert-image").onclick = insertImage;
    document.getElementById("insert-html").onclick = insertHTML;
    document.getElementById("insert-table").onclick = insertTable;
  }
});

async function insertTable() {
  await Word.run(async (context) => {

      const secondParagraph = context.document.body.paragraphs.getFirst().getNext();

      const tableData = [
        ["Name", "ID", "Birth City"],
        ["Bob", "434", "Chicago"],
        ["Sue", "719", "Havana"],
      ];
      secondParagraph.insertTable(3, 3, "After", tableData);
      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function insertHTML() {
  await Word.run(async (context) => {

      const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
      blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function insertImage() {
  await Word.run(async (context) => {

      context.document.body.insertInlinePictureFromBase64(base64Image, "End");

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function replaceText() {
  await Word.run(async (context) => {

      const doc = context.document;
      const originalRange = doc.getSelection();
      originalRange.insertText("many", "Replace");

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function insertTextBeforeRange() {
  await Word.run(async (context) => {

      const doc = context.document;
      const originalRange = doc.getSelection();
      // “End”和“After”的区别在于，
      //“End”在现有区域末尾插入新文本，而“After”则是新建包含字符串的区域，并在现有区域后面插入新区域。
      // 同样，“Start”是在现有区域的开头位置插入文本，而“Before”插入的是新区域。
      // “Replace”将现有区域文本替换为第一个参数中的字符串。
      // before不会包含新加的字符串，换成star就会包括插入的字符串
      originalRange.insertText("Office 2019, ", "Before");
      // originalRange.insertText("Office 2019, ", "Start");
      originalRange.load("text");
      await context.sync();

      doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");

      await context.sync();

  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function insertTextIntoRange() {
  await Word.run(async (context) => {

      const doc = context.document;
      const originalRange = doc.getSelection();
      originalRange.insertText(" (C2R)", "End");

      originalRange.load("text");
      await context.sync();
      
      doc.body.insertParagraph("Original range: " + originalRange.text, "End");
      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function changeFont() {
  await Word.run(async (context) => {

    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function applyStyle() {
  await Word.run(async (context) => {

      // TODO1: Queue commands to style text.
      const firstParagraph = context.document.body.paragraphs.getFirst();
      firstParagraph.styleBuiltIn = Word.Style.intenseReference;
      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function insertParagraph() {
  await Word.run(async (context) => {

    // TODO1: Queue commands to insert a paragraph into the document.
    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
        Word.InsertLocation.start);
    await context.sync();
  })
      .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
      });
}