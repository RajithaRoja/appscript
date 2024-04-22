function doGet() {
    let sheetQuestion = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    let lastRow = sheetQuestion.getLastRow();
    let data = sheetQuestion.getRange(1, 1, lastRow, 19).getValues();
    let englishQuestions = [];
    let japaneseQuestions = [];
    let japaneseRankingArray = [];
    let englishRankingArray = [];
    for (let q = 1; q < lastRow - 1; q += 2) {
        let rankingDataEN = data[q][3];
        englishRankingArray.push(rankingDataEN);
    }

    for (let r = 1; r < lastRow - 1; r += 2) {
        let rankingDataJP = data[r][5];
        japaneseRankingArray.push(rankingDataJP);
    }
    for (let i = 1; i < lastRow; i += 2) {
        let questionId = data[i][0];
        let questionNameEN = data[i][4];
        // let questionNameJP = data[i][6].replace(/^\"|\"$/g, "");
        let questionNameJP = data[i][6];
        let questionTypeOneEN = data[i][2];
        let questionTypeOneJP = data[i][2];
        let topicEN = data[i][3];
        let topicJP = data[i][5];
        let questionTextOneEN = data[i][7];
        let questionTextOneJP = data[i][8];
        let optionsOneEN = [];
        let optionsOneJP = [];

        for (let k = 9; k < 14; k++) {
            if (data[i][k]) {
                optionsOneEN.push({ "value": k - 8, "name": data[i][k] });
            }
        }

        for (let n = 14; n < 19; n++) {
            if (data[i][n]) {
                optionsOneJP.push({ "value": n - 13, "name": data[i][n] });
            }
        }

        englishQuestions.push({
            "id": questionId,
            "name": questionNameEN,
            "questionType1": questionTypeOneEN,
            "questionType2": "",
            "topic": topicEN,
            "questionText1": questionTextOneEN,
            "questionText2": "",
            "options1": optionsOneEN,
            "options2": [],
        });

        japaneseQuestions.push({
            "id": questionId,
            "name": questionNameJP,
            "questionType1": questionTypeOneJP,
            "questionType2": "",
            "topic": topicJP,
            "questionText1": questionTextOneJP,
            "questionText2": "",
            "options1": optionsOneJP,
            "options2": []
        });
    }

    for (let j = 2; j < lastRow; j += 2) {
        let questionTextTwoEN = data[j][7];
        let questionTypeTwoEN = data[j][2];
        let optionsTwoEN = [];
        let questionTextTwoJP = data[j][8];
        let questionTypeTwoJP = data[j][2];
        let optionsTwoJP = [];
        for (let k = 9; k < 14; k++) {
            if (data[j][k]) {
                optionsTwoEN.push({ "value": k - 8, "name": data[j][k] });
            }
        }
        for (let m = 14; m < 19; m++) {
            if (data[j][m]) {
                optionsTwoJP.push({ "value": m - 13, "name": data[j][m] });
            }
        }
        let index = Math.floor((j - 1) / 2);
        if (index >= 0 && index < englishQuestions.length && index < japaneseQuestions.length) {
            englishQuestions[index]["questionType2"] = questionTypeTwoEN;
            englishQuestions[index]["questionText2"] = questionTextTwoEN;
            englishQuestions[index]["options2"] = optionsTwoEN;

            japaneseQuestions[index]["questionType2"] = questionTypeTwoJP;
            japaneseQuestions[index]["questionText2"] = questionTextTwoJP;
            japaneseQuestions[index]["options2"] = optionsTwoJP;
        } else {
            console.error("Index out of bounds:", index);
        }
    }

    let response = {
        "english": { "questions": englishQuestions, "rankingArray": englishRankingArray },
        "japanese": { "questions": japaneseQuestions, "rankingArray": japaneseRankingArray }
    };
    return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}

// To save the request in the sheet and create a pdf
function doPost(e) {
    if (!e || !e.parameter || e.parameter.action !== 'alluser') {
        return ContentService.createTextOutput("Error: Invalid parameter. Expected 'action=alluser'.");
    }
    if (!e || !e.postData || !e.postData.contents) {
        return ContentService.createTextOutput("Error: Missing or invalid request body.");
    }
    let jsonString = e.postData.contents;
    try {
        let jsonData = JSON.parse(jsonString);
        if (!jsonData.answers || !Array.isArray(jsonData.answers)) {
            return ContentService.createTextOutput("Error: Invalid JSON data. 'answers' key should contain an array.");
        }
        let resultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Result');
        let lastRow = resultSheet.getLastRow();
        let startId = lastRow === 1 ? 1 : parseInt(resultSheet.getRange(lastRow, 1).getValue()) + 1;
        let rowData = [startId];

        jsonData.answers.forEach(function (answer) {
            rowData.push(
                answer.questionId,
                answer.option1 || null,
                answer.option2 || null,
                answer.isVisited || null,
                answer.delta || null,
                answer.performanceStatus || null
            );
        });
        rowData.push(jsonData.email || null);
        language = jsonData.language;
        option1 = jsonData.option1;
        option2 = jsonData.option2;
        questionId = jsonData.questionId;
        resultSheet.appendRow(rowData);
        let email = jsonData.email;
        let getEmail = "email" + email;
        let pdfFile = generatePDF(jsonData.answers, jsonData.language, jsonData.questionId, email, getEmail, true);
        let titleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('01 Title');
        let subject = (language === "english") ? titleSheet.getRange(2, 1).getValue() : titleSheet.getRange(2, 5).getValue();
        
        GmailApp.sendEmail(
          email,
          subject,
          "Please find attached a thoughtful analysis of your self-assessment report. We appreciate the effort you've put into this. Your dedication is crucial as you continue to cultivate an entrepreneurial mindset and explore innovative opportunities ahead.",
          { attachments: [pdfFile],   
           name: 'Enterpreneur Mindset', 
          'from': 'noreply@entrepreneurmindset.jp'} );

        return ContentService.createTextOutput("Recorded Successfully.").setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
        Logger.log("Error parsing JSON: " + error);
        return ContentService.createTextOutput(error);
    }
}


// generate the pdf
function generatePDF(answers, language, email, getEmail, moveFirstPage) {
    let surveyReportJP = '起業家精神レポート';
    let surveyReportEN = 'Entrepreneurial Mindset Report';
    let templateId = '1Z-WJOOlB1y0V734Q5ie_Fjg25NhUvmF4Vt2WafU9AJY';
    // Make a copy of the template
    let newDoc = DriveApp.getFileById(templateId).makeCopy();
    let doc = DocumentApp.openById(newDoc.getId());
    let templateBody = doc.getBody();
    let sheetQuestion = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");

    // 1st 2 pages title and overview page
    // if (language == 'english') {
    TitleSheetEn(templateBody, email, getEmail, moveFirstPage);
    OverviewpageEn(templateBody);
    templateBody.appendPageBreak();

    addScaleSheetContent(templateBody, language);
    // 4th page disciplines in japanese and english
    addDisciplineSheetContent(templateBody, language);

    // for the common use for ascending and descending the delta values
    let deltaValues = answers.map(function (answer) {
        return { delta: (answer.delta || 0), questionId: answer.questionId || 0 };
    });

    let deltaValuess = answers.map(function (answer) {
        return { delta: (answer.delta || 0), questionId: answer.questionId || 0 };
    });

    deltaValues.sort(function (a, b) {
        return b.delta - a.delta;
    });

    deltaValuess.sort(function (a, b) {
        return a.delta - b.delta;
    });

    // To get the highest and lowest questionIds of delta
    let lowestTwoQuestionIds = [deltaValues[0].questionId, deltaValues[1].questionId];
    let highestTwoQuestionIds = [deltaValuess[0].questionId, deltaValuess[1].questionId];

    // Values of highest and lowest questions
    let valuesForHighestQuestions = highestTwoQuestionIds.map(function (questionId) {
        return { questionId: questionId, value: getValueFromDisciplinesSheet(sheetQuestion, questionId, templateBody) };
    });
    let valuesForLowestQuestions = lowestTwoQuestionIds.map(function (questionId) {
        return { questionId: questionId, value: getValueFromDisciplinesSheet(sheetQuestion, questionId, templateBody) };
    });

    // O5 th strength sheet
    addAssessmentInsight(templateBody, language, valuesForHighestQuestions);
    // 06 th
    addAssessmentLowestInsight(templateBody, language, valuesForLowestQuestions);

    // 07th conclusion
    let highestValue = valuesForHighestQuestions[0].value;
    let lowestValue = valuesForLowestQuestions[0].value;
    conclusion(highestValue, lowestValue, templateBody, language);

    // chart  in 8th page
    Chartdata(answers, templateBody, language);
    // 09 th sheet 
    processHighestQuestions(valuesForHighestQuestions, templateBody);
    templateBody.appendPageBreak();
    // 10th sheet
    processLowestQuestions(valuesForLowestQuestions, templateBody);

    // Final Page 
    let finalPagesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("11 End Page");
    let dataEndPage = finalPagesheet.getDataRange().getValues();
    let link = dataEndPage[1][0];
    addFullColorPage(templateBody, '#f1a33c', link);

    if (moveFirstPage) {
        let firstPage = doc.getBody().getChild(0).copy();
        let lastChild = doc.getBody().getChild(doc.getBody().getNumChildren() - 1);
        lastChild.getParent().appendParagraph(firstPage);
        doc.getBody().getChild(0).removeFromParent();
    } else {
        let firstPage = doc.getBody().getChild(0);
        if (firstPage.getNumChildren() > 1) {
            for (let i = firstPage.getNumChildren() - 1; i >= 0; i--) {
                if (firstPage.getChild(i).getType() === DocumentApp.ElementType.PARAGRAPH) {
                    firstPage.getChild(i).removeFromParent();
                }
            }
        }
    }
    if (moveFirstPage) {
        doc.getBody().getChild(0).removeFromParent();
    }

    let reportTitle = (language == 'english') ? surveyReportEN : surveyReportJP;
    templateBody.replaceText('{{ReportTitle}}', reportTitle);
    doc.saveAndClose();
    let pdfFile = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
    pdfFile.setName('Entrepreneurial Mindset Report.pdf'); // Set PDF name

    DriveApp.getFileById(doc.getId()).setTrashed(true);
    return pdfFile;
}

// 01 Sheet
// English language 1st page
/**
 * The function `TitleSheetEn` generates a title sheet in a Google Document based on data from a
 * specific sheet in a Google Spreadsheet.
 * @param body - The `body` parameter in the `TitleSheetEn` function represents the body of a Google
 * Document where the content will be added. It is used to append paragraphs with specific formatting
 * and content based on the data retrieved from a spreadsheet.
 * @param email - The `TitleSheetEn` function takes several parameters, including `body`, `email`,
 * `getEmail`, and `moveFirstPage`. In the context of the function, the `email` parameter is not
 * directly used. Instead, the `getEmail` parameter is used to construct a paragraph in the Google
 * @param getEmail - The `getEmail` parameter in the `TitleSheetEn` function is used to specify the
 * email address that will be included in the generated document. This email address is concatenated
 * with other values before being appended to the document.
 * @param moveFirstPage - The `moveFirstPage` parameter in the `TitleSheetEn` function is a boolean
 * value that determines whether to move the content to the first page of the document. If
 * `moveFirstPage` is `true`, the content will be moved to the first page. If it is `false`,
 */
function TitleSheetEn(body, email, getEmail, moveFirstPage) {
    let rowFirstNumber;
    let rowsecondNumber;
    let rowthirdNumber;

    if (language == "english") {
        rowFirstNumber = 0;
        rowsecondNumber = 2;
        rowthirdNumber = 3;
    } else {
        rowFirstNumber = 4;
        rowsecondNumber = 6;
        rowthirdNumber = 7;
    }
    if (!moveFirstPage) {
        let nextElement = body.getNextSibling();
        if (nextElement && nextElement.getType() === DocumentApp.ElementType.PARAGRAPH && nextElement.asParagraph().getText().trim() === "") {
            nextElement.removeFromParent();
        }
    }
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('01 Title');
    let data = sheet.getDataRange().getValues();
    body.appendParagraph(data[1][rowFirstNumber]).setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.LEFT).setForegroundColor('#423d37');
    body.appendParagraph("")
    body.appendParagraph(getCurrentDate()).setHeading(DocumentApp.ParagraphHeading.SUBTITLE).setAlignment(DocumentApp.HorizontalAlignment.LEFT).setForegroundColor('#423d37');
    body.appendParagraph("")
    let valuesBeforeColon = getValuesBeforeColon(data[1][rowsecondNumber]);
    body.appendParagraph(valuesBeforeColon + " : " + getEmail).setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.LEFT).setForegroundColor('#423d37');
    body.appendParagraph("")
    body.appendParagraph(data[1][rowthirdNumber]).setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.LEFT).setForegroundColor('#423d37').setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
}

// Call the TitleSheetEn function with email and email1 as arguments
/**
 * The function `getCurrentDate` returns the current date in a formatted string.
 * @returns The `getCurrentDate` function returns the current date in the format "day month year", for
 * example "15 September 2022".
 */
function getCurrentDate() {
    const today = new Date();
    const day = today.getDate();
    const month = today.toLocaleDateString('en-US', { month: 'long' });
    const year = today.getFullYear();
    const formattedDate = `${day} ${month} ${year}`;
    return formattedDate;
}

// 02 overview page 
// English overview page
/**
 * The function `OverviewpageEn` generates a document in Google Docs based on data from a specific
 * sheet, with different formatting based on the language selected.
 * @param body - The function `OverviewpageEn` takes a `body` parameter and uses it to append content
 * to a Google Docs document. The content is retrieved from a specific sheet named "02 Overview" in the
 * active spreadsheet.
 */
function OverviewpageEn(body) {
    let columnNum;
    if (language == 'english') {
        columnNum = 0
    } else {
        columnNum = 1
    }
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("02 Overview");
    let data = sheet.getDataRange().getValues();
    body.appendPageBreak();
    body.appendParagraph(data[1][columnNum]).setHeading(DocumentApp.ParagraphHeading.HEADING4).setAlignment(DocumentApp.HorizontalAlignment.LEFT).setFontSize(16).setForegroundColor('#423d37').setBold(true).setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    body.appendParagraph("");
    body.appendParagraph(data[3][columnNum]).setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.LEFT).setFontSize(11).setForegroundColor('#423d37').setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    if (language == "japanese") {
        body.appendParagraph("");
    }
    boldFirstWord(data[5][columnNum], body);
    if (language == "japanese") {
        body.appendParagraph("");
    }
    boldFirstWord(data[6][columnNum], body);
    if (language == "japanese") {
        body.appendParagraph("");
    }
    body.appendParagraph(data[8][columnNum]).setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.LEFT).setFontSize(11).setForegroundColor('#423d37').setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    body.appendParagraph("");
    body.appendParagraph(data[10][columnNum].replace(/^-/, "1.  ")).setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.LEFT).setIndentFirstLine(50).setFontSize(11).setForegroundColor('#423d37').setIndentStart(72);
    body.appendParagraph(data[11][columnNum].replace(/^-/, "2.  ")).setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.LEFT).setIndentFirstLine(50).setForegroundColor('#423d37');
    body.appendParagraph("");
    body.appendParagraph(data[13][columnNum]).setHeading(DocumentApp.ParagraphHeading.NORMAL).setAlignment(DocumentApp.HorizontalAlignment.LEFT).setForegroundColor('#423d37');
    if (language == "japanese") {
        body.appendParagraph("");
    }
    boldFirstWord(data[15][columnNum], body);
    body.appendParagraph("");
    body.appendParagraph(data[17][columnNum]).setHeading(DocumentApp.ParagraphHeading.NORMAL).setForegroundColor('#423d37').setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
}


// get the values before the colon
/**
 * The function `getValuesBeforeColon` splits an input string by colon and returns an array of values
 * before each colon.
 * @param inputString - I see that you have a JavaScript function `getValuesBeforeColon` that takes an
 * `inputString` as a parameter. This function splits the input string by the colon (':') delimiter and
 * returns an array of values before the colon.
 * @returns The function `getValuesBeforeColon` returns an array containing the values before the colon
 * in the input string.
 */
function getValuesBeforeColon(inputString) {
    let parts = inputString.split(':');
    let valuesBeforeColon = [];
    for (let i = 0; i < parts.length - 1; i++) {
        valuesBeforeColon.push(parts[i].trim());
    }
    return valuesBeforeColon;
}

// bolding first text 
/**
 * The function `boldFirstWord` takes a text input, finds the first word before a colon, makes it bold
 * and sets it as a heading in a Google Docs body.
 * @param text - The `boldFirstWord` function takes two parameters: `text` and `body`. The `text`
 * parameter is the input text that you want to format, and the `body` parameter is the document body
 * where you want to append the formatted text.
 * @param body - The `body` parameter in the `boldFirstWord` function represents the body of a document
 * where the text will be added. It seems like the function is designed to format the text by making
 * the first word bold and changing its color to `#423d37`. If the text contains a colon
 */
function boldFirstWord(text, body) {
    let firstWordEndIndex = text.indexOf(":");
    if (firstWordEndIndex != -1) {
        let firstWord = text.substring(0, firstWordEndIndex);
        let restOfText = text.substring(firstWordEndIndex + 1);
        let combinedText = firstWord + " " + restOfText;
        body.appendParagraph(text)
            .setHeading(DocumentApp.ParagraphHeading.HEADING1)
            .setFontSize(11)
            .setSpacingAfter(10)
            .editAsText()
            .setBold(0, firstWord.length, true).setForegroundColor('#423d37');
    } else {
        body.appendParagraph(text);
    }
}

// 3rd scale page
/**
 * The function `addScaleSheetContent` generates a formatted document content based on data from a
 * Google Sheets document, with options for English or another language.
 * @param body - The `body` parameter in the `addScaleSheetContent` function represents the body of a
 * Google Docs document where the content will be added. This function is designed to populate the
 * document with content retrieved from a specific sheet in a Google Spreadsheet based on the specified
 * language (English or another language). The
 * @param language - The `language` parameter in the `addScaleSheetContent` function determines whether
 * to retrieve content in English or another language from the "03 Scale" sheet in the active
 * spreadsheet. If the language is set to 'english', the function will fetch content from column A,
 * otherwise, it will fetch content
 */
function addScaleSheetContent(body, language) {
    let paragraphStyle = {};
    paragraphStyle[DocumentApp.Attribute.INDENT_END] = 5;
    let titleText;
    let scaleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("03 Scale");
    if (language === 'english') {
        titleText = scaleSheet.getRange("A2").getValue();
    } else {
        titleText = scaleSheet.getRange("B2").getValue();
    }
    let title = body.appendParagraph(titleText);
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1).setFontSize(16).setSpacingAfter(36).setBold(true).setForegroundColor('#423d37');
    let table = body.appendTable();

    let leftColumn = table.appendTableRow().appendTableCell();
    let rightColumn = table.getRow(0).appendTableCell();
    for (let i = 4; i <= 28; i++) {
        let leftValue = (language === "english") ? scaleSheet.getRange("A" + i).getValue() : scaleSheet.getRange("B" + i).getValue();
        if (i == 4) {
            let leftParagraph = leftColumn.appendParagraph(leftValue);
            leftParagraph.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY).setFontSize(12).setIndentStart(5).setIndentFirstLine(5).setAttributes(paragraphStyle).setForegroundColor('#423d37');
            leftParagraph.editAsText().setBold(true);
            leftColumn.appendParagraph('');
        }
        if (i == 5) {
            let leftParagraph = leftColumn.appendParagraph(leftValue);
            leftParagraph.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY).setFontSize(12).setIndentStart(5).setIndentFirstLine(5).setAttributes(paragraphStyle).setForegroundColor('#423d37');
            leftParagraph.editAsText().setBold(false);
        }
        if (i == 15) {
            rightparagraph = rightColumn.appendParagraph(leftValue);
            rightparagraph.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY).setFontSize(12).setIndentStart(5).setIndentFirstLine(5).setAttributes(paragraphStyle).setForegroundColor('#423d37');
            rightparagraph.editAsText().setBold(false);
        }
        if (i == 7) {

            var leftParagraph = leftColumn.appendParagraph('\t   ' + leftValue);
            leftParagraph.editAsText().setBold(true).setForegroundColor('#423d37');
            leftColumn.appendParagraph('');
        }
        if (i == 6) {
            if (language == "english") {
                var leftParagraph = leftColumn.appendParagraph('\t   ' + leftValue);
                leftParagraph.editAsText().setBold(true).setForegroundColor('#423d37');
            }
            else {
                leftColumn.appendParagraph("")
                var leftParagraph = leftColumn.appendParagraph('\t   ' + leftValue);
                leftParagraph.editAsText().setBold(true).setForegroundColor('#423d37');
            }
        }
        if (i == 14) {
            rightparagraph = rightColumn.appendParagraph(leftValue);
            rightparagraph.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY).setFontSize(12).setIndentStart(5).setIndentFirstLine(5).setAttributes(paragraphStyle);
            rightparagraph.editAsText().setBold(true).setForegroundColor('#423d37');
            rightColumn.appendParagraph('');
        }
        if (i == 17) {
            rightparagraph = rightColumn.appendParagraph('\t   ' + leftValue);
            rightparagraph.editAsText().setBold(true).setForegroundColor('#423d37');
            rightColumn.appendParagraph('');

        }
        if (i == 16) {
            if (language == 'english') {
                rightColumn.appendParagraph('');
                rightparagraph = rightColumn.appendParagraph('\t   ' + leftValue);
                rightparagraph.editAsText().setBold(true).setForegroundColor('#423d37');
            } else {
                rightparagraph = rightColumn.appendParagraph('\t   ' + leftValue);
                rightparagraph.editAsText().setBold(true).setForegroundColor('#423d37');
            }
        }
        if (i > 7 && i < 14) {
            let leftParagraph = leftColumn.appendParagraph('\t   ' + leftValue);
            leftParagraph.editAsText().setBold(false).setForegroundColor('#423d37');
        }
        if (i > 17 && i < 23) {
            let rightParagraph = rightColumn.appendParagraph('\t   ' + leftValue);
            rightParagraph.editAsText().setBold(false).setForegroundColor('#423d37');
        }
    }
    table.setBorderWidth(0.5);
    table.setBorderColor('#423d37');
    body.appendPageBreak();
}


// 4th page
/**
 * The function `addDisciplineSheetContent` populates a Google Docs document with content from a
 * specific sheet based on the chosen language.
 * @param body - The `body` parameter in the `addDisciplineSheetContent` function represents the body
 * of a Google Docs document where the content from the "04 Discipline" sheet will be added. This
 * function reads data from the "04 Discipline" sheet in a Google Spreadsheet and appends the content
 * to the provided
 * @param language - The `language` parameter in the `addDisciplineSheetContent` function determines
 * whether the content should be displayed in English or another language. The function checks the
 * language parameter to decide which column of data to use from the "04 Discipline" sheet in the
 * Google Spreadsheet. If the language is set to
 */
function addDisciplineSheetContent(body, language) {
    let disciplineSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("04 Discipline");
    let disciplineContent = disciplineSheet.getDataRange().getValues();
    let disciplineText = "";
    let secondHeading;
    let textDisciplines;
    if (language === 'english') {
        textDisciplines = disciplineSheet.getRange("A2").getValue();
    } else {
        textDisciplines = disciplineSheet.getRange("B2").getValue();
    }
    if (language === 'english') {
        secondHeading = disciplineSheet.getRange("A3").getValue();
    }
    else {
        secondHeading = disciplineSheet.getRange("B3").getValue();
    }
    body.appendParagraph(textDisciplines.toUpperCase()).setAlignment(DocumentApp.HorizontalAlignment.LEFT).setBold(true).setFontSize(16);
    body.appendParagraph(secondHeading).setAlignment(DocumentApp.HorizontalAlignment.LEFT).setBold(false).setFontSize(11);
    body.appendParagraph("");
    for (let j = 4; j < disciplineContent.length; j++) {
        let disciplineValue;
        if (language === 'english') {
            disciplineValue = disciplineContent[j][0];
        } else {
            disciplineValue = disciplineContent[j][1];
        }
        disciplineText += disciplineValue;
        if (language == "japanese" && j === 22) {
            body.appendPageBreak();

        }
        if (j % 3 === 1) {

            body.appendParagraph(disciplineText).setFontSize(12).setForegroundColor('#f1a33c').setBold(true).setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
            ;
            body.appendParagraph("").setBold(false);
            disciplineText = "";


        } else {
            if (j !== disciplineContent.length) {
                body.appendParagraph(disciplineText).setFontSize(12).setForegroundColor('#423d37').setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
                ;
                disciplineText = "";
            }
        }
    }
    body.appendPageBreak();
}

// To get the questionId and the values from the questionSheet 
/**
 * The function `getValueFromDisciplinesSheet` retrieves a value from a specific cell in a Google
 * Sheets document based on the question ID.
 * @param sheetQuestion - `sheetQuestion` is likely a reference to a Google Sheets object representing
 * a specific sheet where data is stored. This object is used to access and retrieve values from the
 * sheet based on the provided `questionId` and other parameters.
 * @param questionId - The `questionId` parameter in the `getValueFromDisciplinesSheet` function
 * represents the ID of a question. It is used to calculate the row number in the spreadsheet where the
 * value associated with that question is located.
 * @param body - The `body` parameter in the `getValueFromDisciplinesSheet` function is typically a
 * document body object where you can append paragraphs or content. In the provided code snippet, there
 * are commented out lines that show how you can append the fetched value, row number, and question ID
 * to the document
 * @returns The function `getValueFromDisciplinesSheet` is returning the value fetched from a specific
 * cell in the `sheetQuestion` based on the provided `questionId`.
 */
function getValueFromDisciplinesSheet(sheetQuestion, questionId, body) {
    try {
        let rowNumber = questionId + 1;
        let columnNumber = 4;
        let value = sheetQuestion.getRange(rowNumber, columnNumber).getValue();
        // body.appendParagraph(value)
        // body.appendParagraph(rowNumber)
        // body.appendParagraph(questionId)
        return value;
    } catch (error) {
        console.error("Error fetching value from '" + rowNumber + "' sheet:", error);
        return null;
    }
}

// 5th sheet
// matching the row in 05 strength
/**
 * The function `findMatchingRows` searches for rows in a dataset that match specific discipline values
 * and returns the matching row indices, while the function `appendMatchingRowsToBody` appends the
 * matching rows to a document body based on language conditions.
 * @param values - The `values` parameter in the `findMatchingRows` function is an array of rows, where
 * each row is an array containing values. The function iterates over these rows to find matching rows
 * based on certain conditions.
 * @param disciplineValue - The `disciplineValue` parameter in the `findMatchingRows` function is an
 * array that contains the discipline values to match against the first and second parts of the row
 * value split by ' - '. It is used to check if both parts of the row value are included in the
 * `disciplineValue
 * @param reversedDisciplineValue - The `reversedDisciplineValue` parameter in the `findMatchingRows`
 * function is used to check for rows where the values in the first and second parts of the row match
 * the values in the `reversedDisciplineValue` array. If both values in a row match the values in the
 * @returns The `findMatchingRows` function returns an object with two properties: `matchingRows` and
 * `matchingReversedRows`, which are arrays containing the row numbers of matching rows based on the
 * conditions specified in the function.
 */
function findMatchingRows(values, disciplineValue, reversedDisciplineValue) {
    let matchingRows = [];
    let matchingReversedRows = [];

    values.forEach(function (row, index) {
        let rowValue = row[0];
        let rowParts = rowValue.split(' - ');

        if (rowParts.length === 2 &&
            disciplineValue.includes(rowParts[0].trim()) &&
            disciplineValue.includes(rowParts[1].trim())) {
            matchingRows.push(index + 1);
        }

        if (rowParts.length === 2 &&
            reversedDisciplineValue.includes(rowParts[0].trim()) &&
            reversedDisciplineValue.includes(rowParts[1].trim())) {
            matchingReversedRows.push(index + 1);
        }
    });
    return { matchingRows: matchingRows, matchingReversedRows: matchingReversedRows };
}

// To find the row in 05th strength sheet
function appendMatchingRowsToBody(body, values, matchingRows, matchingReversedRows) {
    if (matchingRows.length > 0) {
        matchingRows.forEach(function (rowNumber) {
            let row = values[rowNumber - 1];
            if (language === 'english') {
                body.appendParagraph(row[3]).setItalic(false).setFontSize(12).setBold(false).setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

            } else {
                body.appendParagraph(row[7]).setItalic(false).setFontSize(12).setBold(false).setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
            }
        });
    } else if (matchingReversedRows.length > 0) {
        matchingReversedRows.forEach(function (rowNumber) {
            let row = values[rowNumber - 1];
            if (language === 'english') {
                body.appendParagraph(row[3]).setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

            } else {
                body.appendParagraph(row[4]).setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);

            }
        });
    } else {
        body.appendParagraph("");
    }
}

// 5th page heading
/**
 * The function `addAssessmentInsight` generates assessment insights in either English or Japanese
 * language and appends them to a document body with specific formatting.
 * @param body - The `body` parameter in the `addAssessmentInsight` function is likely a reference to
 * the body of a document in Google Apps Script. This parameter is used to append paragraphs and format
 * text within the document.
 * @param language - The `language` parameter in the `addAssessmentInsight` function is used to
 * determine whether the assessment insights should be displayed in English or another language. If the
 * `language` is set to 'english', the assessment insights will be displayed in English. Otherwise, if
 * it is set to another
 * @param valuesForHighestQuestions - The `valuesForHighestQuestions` parameter is an array containing
 * values for the highest questions in the assessment. The function `addAssessmentInsight` uses these
 * values to generate insights based on the assessment results.
 * @returns The function `addAssessmentInsight` is appending assessment insights to a Google Docs body
 * based on the provided language and values for highest questions. It is also finding matching rows in
 * a spreadsheet based on a discipline value and appending them to the body.
 */
function addAssessmentInsight(body, language, valuesForHighestQuestions) {
    let assessment;
    if (language == 'english') {
        assessment = "Assessment Insights";
    } else {
        assessment = "評価に関する洞察";
    }

    body.appendParagraph(assessment.toUpperCase()).setBold(true).setFontSize(16).setItalic(false).setForegroundColor('#423d37').setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    let disciplineValue = valuesForHighestQuestions.map(entry => entry.value).join(' - ');
    let reversedDisciplineValue = disciplineValue.split(' - ').reverse().join(' - ');
    disciplineValue = disciplineValue.toLowerCase();
    disciplineValue = disciplineValue.replace("&", "-").replace(/\b\w/g, function (char) {
        return char.toUpperCase();
    });

    body.appendParagraph("");
    let strengthSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("05 Strengths");
    let range = strengthSheet.getDataRange();
    let values = range.getValues();
    let strengthResult05 = findMatchingRows(values, disciplineValue, reversedDisciplineValue);
    let matchingRows = strengthResult05.matchingRows;
    let matchingReversedRows = strengthResult05.matchingReversedRows;
    appendMatchingRowsToBody(body, values, matchingRows, matchingReversedRows);
}

// 6th page
/**
 * The function `addAssessmentLowestInsight` processes values for the lowest questions and appends
 * matching rows to a document body based on specified conditions.
 * @param body - The `body` parameter in the `addAssessmentLowestInsight` function is likely a
 * reference to the document or content to which you want to append information. It could be a Google
 * Docs body object or similar representation in the context of the code you provided. This function
 * seems to be adding
 * @param language - The `language` parameter in the `addAssessmentLowestInsight` function is used to
 * determine whether to append a page break or an empty paragraph to the `body` based on the language
 * specified. If the language is "english", a page break is appended; otherwise, an empty paragraph
 * @param valuesForLowestQuestions - valuesForLowestQuestions is an array of objects containing values
 * for the lowest questions. Each object in the array has a property called "value".
 * @returns The function `addAssessmentLowestInsight` is returning the matching rows and matching
 * reversed rows from the `strengthSheet` based on the discipline value and reversed discipline value
 * that are calculated within the function.
 */
function addAssessmentLowestInsight(body, language, valuesForLowestQuestions) {
    let disciplineValue = valuesForLowestQuestions.map(entry => entry.value).join(' - ');
    let reversedDisciplineValue = disciplineValue.split(' - ').reverse().join(' - ');
    disciplineValue = disciplineValue.toLowerCase();
    disciplineValue = disciplineValue.replace("&", "-").replace(/\b\w/g, function (char) {
        return char.toUpperCase();
    });
    if (language == "english") {
        body.appendPageBreak();
    } else {
        body.appendParagraph("");
    }
    let strengthSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("06 Weakness");
    let range = strengthSheet.getDataRange();
    let values = range.getValues();
    let strengthResult05 = findMatchingRows(values, disciplineValue, reversedDisciplineValue);
    let matchingRows = strengthResult05.matchingRows;
    let matchingReversedRows = strengthResult05.matchingReversedRows;
    appendMatchingRowsToBody(body, values, matchingRows, matchingReversedRows);
}

// Conclusion Page 7th
/**
 * The function `conclusion` retrieves data from a Google Sheets document and appends it to a Google
 * Docs document based on specified values and language.
 * @param value1 - Value1 is the first value used to find a match in the spreadsheet data.
 * @param value2 - It seems like you were about to provide more information about the `value2`
 * parameter in the `conclusion` function. How can I assist you further with this?
 * @param body - The `body` parameter in the `conclusion` function represents the body of a Google
 * Document where the conclusion text will be added. This function is designed to retrieve data from a
 * Google Spreadsheet and append the conclusion text in either English or Japanese to the provided
 * document body based on the specified language.
 * @param language - The `language` parameter in the `conclusion` function determines whether the
 * conclusion should be added in English or Japanese to the document. If the `language` is set to
 * `'english'`, the conclusion will be added in English, and if it's set to any other value, the
 * conclusion will
 */
function conclusion(value1, value2, body, language) {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("07 Conclusion");
    let range = sheet.getRange("B2:B111");
    let values = range.getValues();
    let range2 = sheet.getRange("C1:C111");
    let values2 = range2.getValues();
    const allMatches = findMatchingIndex(values, values2, value1, value2);
    let resultEn = sheet.getRange("D2:D111").getValues()[allMatches];
    let resultJP = sheet.getRange("H2:H111").getValues()[allMatches];

    let resultEnString = Array.isArray(resultEn) ? resultEn.join(', ') : '';
    let resultJPString = Array.isArray(resultJP) ? resultJP.join(', ') : '';

    if (language == 'english') {
        body.appendPageBreak();

        body.appendParagraph(resultEnString).setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    } else {
        body.appendParagraph("")
        body.appendParagraph(resultJPString).setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
    }
}

// matching the value for the 7th sheet
/**
 * This JavaScript function finds the index of matching elements in two arrays based on specified
 * texts.
 * @param array1 - Array of arrays containing text elements
 * @param array2 - It seems like you were about to provide more information about the parameters, but
 * the message got cut off. Could you please provide more details about the `array2` parameter so that
 * I can assist you further?
 * @param text1 - The `text1` parameter in the `findMatchingIndex` function represents the text that
 * you want to match in the first element of `array1`. The function will iterate through `array1` and
 * `array2` to find a matching pair where the first element of `array1` matches
 * @param text2 - It looks like you forgot to provide the definition for `text2`. Could you please
 * provide more information or clarify what `text2` represents in this context?
 * @returns the index of the matching element in the arrays `array1` and `array2` where the first
 * element in each array matches the provided `text1` and `text2`. If no match is found, it returns -1.
 */
function findMatchingIndex(array1, array2, text1, text2) {
    for (let i = 0; i < array1.length; i++) {
        if (array1[i][0] === text1 && array2[i][0] === text2) {
            return i - 1;
        }
    }
    return -1;
}

/**
 * The function `Chartdata` generates a chart based on provided data and language settings to visualize
 * performance gaps in a document.
 * @param answers - The `answers` parameter is an array of objects containing information about
 * opportunities and scores. Each object in the array has the following properties:
 * @param body - The function `Chartdata` takes in three parameters: `answers`, `body`, and `language`.
 * @param language - The `language` parameter in the `Chartdata` function determines the language
 * setting for the chart and document being generated. The function uses this parameter to customize
 * the content and formatting based on the selected language, which can be either 'english' or another
 * language.
 * @returns The `Chartdata` function is returning a chart image that displays the top 11 question IDs
 * along with their corresponding performance gaps. The chart is a horizontal bar chart with specific
 * styling and formatting options based on the language parameter provided.
 */
function Chartdata(answers, body, language) {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('08 Graphic');
    let data = sheet.getDataRange().getValues();
    let sheetQuestion = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");

    let deltaValues = answers.map(function (answer) {
        return { opportunityScore: (answer.delta || 0) * (answer.option2 || 0), questionId: answer.questionId || 0, delta: answer.delta };
    });

    let opportunityScoreSortValue = deltaValues.sort(function (a, b) {
        return b.opportunityScore - a.opportunityScore;
    });

    let dataRows = opportunityScoreSortValue.slice(0, 11).map(function (value) {
        return [value.questionId, value.opportunityScore]; // Include questionId in the data rows
    });

    let graphHeading, focusedArea, monitorArea, alignedArea;
    if (language === 'english') {
        graphHeading = data[1][0].toUpperCase();
        focusedArea = data[2][0].split(' ')[0];
        monitorArea = data[3][0].split(' ')[0];
        alignedArea = data[4][0].split(' ')[0];
    } else {
        graphHeading = data[1][1];
        focusedArea = data[2][1];
        monitorArea = data[3][1];
        alignedArea = data[4][1];
    }

    let focusedRest = data[2][0].split(' ').slice(1).join(' ');
    let monitorRest = data[3][0].split(' ').slice(1).join(' ');
    let alignedRest = data[4][0].split(' ').slice(1).join(' ');
    body.appendPageBreak();
    body.appendParagraph(graphHeading).setHeading(DocumentApp.ParagraphHeading.HEADING4).setAlignment(DocumentApp.HorizontalAlignment.LEFT).setBold(true).setForegroundColor('#423d37').setFontSize(16);
    body.appendParagraph('');
    body.appendParagraph('');
    if (language === 'english') {
        body.appendParagraph("\t\t    " + focusedArea + "\t\t\t     " + monitorArea + "\t\t\t    " + alignedArea + "\n\t\t" + focusedRest + "\t\t\t" + monitorRest + "\t\t\t" + alignedRest).setFontSize(8).setBold(false);
    } else {
        body.appendParagraph("\t             " + focusedArea + "\t\t             " + monitorArea + "\t\t        " + alignedArea).setFontSize(8).setBold(false);
    }
    let imageUrl = "https://drive.google.com/uc?id=1Ba4Ur-fy6cB5oViDQxxjtetFGFmkpEWD";
    let imageBlobs = UrlFetchApp.fetch(imageUrl).getBlob();
    let image = body.appendImage(imageBlobs);
    image.setWidth(600).setHeight(20);
    let dataTable = Charts.newDataTable()
        .addColumn(Charts.ColumnType.STRING, 'Question ID')
        .addColumn(Charts.ColumnType.NUMBER, 'Performance Gap');

    if (language == "english") {
        dataRows.forEach(function (value) {
            let questionId = String(value[0]);
            let row = parseInt(questionId) + 1;
            let title = sheetQuestion.getRange(row, 4).getValue();
            dataTable.addRow([title, value[1]]);
        });
    } else {
        dataRows.forEach(function (value) {
            let questionId = String(value[0]);
            let row = parseInt(questionId) + 1;
            let title = sheetQuestion.getRange(row, 6).getValue();
            dataTable.addRow([title, value[1]]);
        });
    }
    let chartBuilder = Charts.newBarChart();
    let chart = chartBuilder.setDataTable(dataTable)
        .setXAxisTitle('')
        .setYAxisTitle((language == 'english') ? 'Performance Gap' : 'パフォーマンスのギャップ')
        .setDimensions(650, 430)
        .setOption('orientation', 'horizontal')
        .setOption('hAxis.slantedText', true)
        .setOption('hAxis.slantedTextAngle', 90)
        .setOption('hAxis.textStyle.fontSize', 16)
        .setOption('vAxis.textStyle.fontSize', 16)
        .setOption('bar.groupWidth', '40%')
        .setOption('colors', ['#f1a33c'])
        .setOption('hAxis.minorGridlines', { count: 5 })
        .setOption('hAxis.textStyle.fontSize', 14)
        .setOption('series', {
            0: { targetAxisIndex: 0 },
            1: { targetAxisIndex: 1 }
        })
        .setOption('legend', { position: 'none' })
        .setOption('chartArea', { top: '8%', right: '10%', width: '80%', bottom: '35%' })
        .setOption('backgroundColor', '#f9f6f1')
        .setOption('annotations.highContrast', true)
        .build();
    body.appendImage(chart.getAs('image/png'));
}

// The functions of only 9th page
// To get the value of the strongest disciplines
/**
 * The function `processHighestQuestions` iterates through values, finds matching rows, retrieves
 * corresponding values, and appends them to a document body.
 * @param valuesForHighestQuestions - valuesForHighestQuestions is an array containing values that need
 * to be processed for the highest questions.
 * @param body - The `processHighestQuestions` function takes two parameters:
 * `valuesForHighestQuestions` and `body`.
 */
function processHighestQuestions(valuesForHighestQuestions, body) {
    valuesForHighestQuestions.forEach(function (entry) {
        let valueFromDiscipline = entry.value;
        let matchingRow = findMatchingRow09Sheet(valueFromDiscipline);
        if (matchingRow !== -1) {
            let correspondingValue = getCorrespondingValue(matchingRow);
            body.appendParagraph(correspondingValue).setBold(false).setFontSize(12).setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
            body.appendParagraph("");
        } else {
            body.appendParagraph("");
        }
    });
}


// To match the 09 the sheet value
/**
 * The function `findMatchingRow09Sheet` searches for a specific value in the '09 Q Strengths' sheet
 * and returns the row number where the value is found.
 * @param value - The function `findMatchingRow09Sheet` is designed to search for a specific value in
 * the '09 Q Strengths' sheet of a Google Spreadsheet and return the row number where the value is
 * found. If the value is not found, it returns -1.
 * @returns The function `findMatchingRow09Sheet` is returning the row number (index + 1) where the
 * value matches in the '09 Q Strengths' sheet. If no match is found, it returns -1.
 */
function findMatchingRow09Sheet(value) {
    let strengthsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('09 Q Strengths');
    let data = strengthsSheet.getRange(1, 1, strengthsSheet.getLastRow(), 1).getValues();
    for (let i = 0; i < data.length; i++) {
        if (data[i][0] == value) {
            return i + 1;
        }
    }
    return -1;
}

// To match the value of 9th sheet
/**
 * The function `getCorrespondingValue` retrieves a value from a specific row in a Google Sheets
 * spreadsheet based on the language specified.
 * @param row - The `row` parameter in the `getCorrespondingValue` function represents the row number
 * from which you want to retrieve a corresponding value from the '09 Q Strengths' sheet in a Google
 * Sheets document.
 * @returns The `getCorrespondingValue` function returns the corresponding value from the '09 Q
 * Strengths' sheet based on the row number provided as an argument. The value returned depends on the
 * `language` variable - if `language` is 'english', the value from column 2 is returned, otherwise the
 * value from column 4 is returned.
 */
function getCorrespondingValue(row) {
    let correspondingValue;
    let strengthsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('09 Q Strengths');
    if (language === 'english') {
        correspondingValue = strengthsSheet.getRange(row, 2).getValue();
    } else {
        correspondingValue = strengthsSheet.getRange(row, 4).getValue();
    }
    return correspondingValue;
}

// 10th sheet
// To get the value of the weekest disciplines
/**
 * The function `processLowestQuestions` iterates through values and finds corresponding values in a
 * spreadsheet, appending them to a document body or indicating if no match is found.
 * @param valuesForLowestQuestions - The `valuesForLowestQuestions` parameter is an array containing
 * objects with information about the lowest questions. Each object in the array has a `value` property
 * that represents a value from a discipline and a `questionId` property that identifies the question.
 * @param body - The `processLowestQuestions` function takes two parameters:
 */
function processLowestQuestions(valuesForLowestQuestions, body) {
    valuesForLowestQuestions.forEach(function (entry) {
        let valueFromDiscipline = entry.value;
        let matchingRow = findMatchingRow10Sheet(valueFromDiscipline);
        if (matchingRow !== -1) {
            let correspondingValue = getCorrespondingValueWeek(matchingRow);
            body.appendParagraph(correspondingValue).setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
            body.appendParagraph("");
        } else {
            body.appendParagraph(`No match found in '09 Q Strengths' for Lowest Question ${entry.questionId}`);
        }
    });
}

// To find the matching data from the sheet 10
/**
 * The function `findMatchingRow10Sheet` searches for a specific value in the '10 Q Weakness' sheet and
 * returns the row number where the value is found.
 * @param value - The function `findMatchingRow10Sheet` is designed to search for a specific value in
 * the '10 Q Weakness' sheet of a Google Spreadsheet and return the row number where the value is
 * found. The `value` parameter represents the value you want to search for in the sheet.
 * @returns The function `findMatchingRow10Sheet` is returning the row number (index + 1) where the
 * value matches in the '10 Q Weakness' sheet. If no match is found, it returns -1.
 */
function findMatchingRow10Sheet(value) {
    let weekestSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('10 Q Weakness');
    let data = weekestSheet.getRange(1, 1, weekestSheet.getLastRow(), 1).getValues();
    for (let i = 0; i < data.length; i++) {
        if (data[i][0] == value) {
            return i + 1;
        }
    }
    return -1;
}

// To get the weekdiscipline data from sheet 10
/**
 * The function `getCorrespondingValueWeek` retrieves a value from a specific row in a spreadsheet
 * based on the language selected.
 * @param row - The `row` parameter in the `getCorrespondingValueWeek` function represents the row
 * number from which you want to retrieve a corresponding value from the spreadsheet.
 * @returns The function `getCorrespondingValueWeek(row)` returns the corresponding value from the '10
 * Q Weakness' sheet based on the row number provided as an argument. The value returned depends on the
 * language variable - if the language is 'english', it retrieves the value from column 2, otherwise
 * from column 4 of the sheet.
 */
function getCorrespondingValueWeek(row) {
    let correspondingValue;
    let weekestSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('10 Q Weakness');
    if (language === 'english') {
        correspondingValue = weekestSheet.getRange(row, 2).getValue();
    } else {
        correspondingValue = weekestSheet.getRange(row, 4).getValue();
    }
    return correspondingValue;
}

// Final Page
/**
 * The function `addFullColorPage` adds a full-color page with specified text to a Google Docs
 * document.
 * @param body - The `body` parameter in the `addFullColorPage` function represents the body of a
 * Google Docs document where the content will be added. It is a reference to the body of the document
 * where you want to insert a new full-color page with specified text.
 * @param color - The `color` parameter in the `addFullColorPage` function is used to specify the color
 * of the border of the table that will be added to the document page. It should be a valid color value
 * that can be recognized by Google Apps Script, such as '#RRGGBB' for
 * @param text - The `addFullColorPage` function takes three parameters:
 */
function addFullColorPage(body, color, text) {
    body.appendPageBreak();
    var table = body.appendTable([['']]);
    table.setBorderWidth(1);
    table.setBorderColor(color);
    var cell = table.getCell(0, 0);
    cell.setWidth(450);
    cell.clear();
    cell.appendParagraph("\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");
    var paragraph = cell.appendParagraph(text);
    paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    paragraph.setFontSize(20);
    paragraph.setForegroundColor('#ffffff');
    paragraph.setLineSpacing(1.5);
}

