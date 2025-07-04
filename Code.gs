function appendSessionsToDoc() {
  const sheetId = '1y8klx-mnV6QEKnwIeLbkwBqkY2v8i5smgY9dIuIXxgc'; // Spreadsheet ID
  const docId = '1Hkos9mxsiVzZQGew3W4p976z8UVoQ-EDiUvb3_fms3w'; // Target Google Doc ID

  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName('SessionsYYC');
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Fields to skip if empty
  const skipIfEmpty = [
    'Co-Presenter 1', 'Co-presenter email 1', 'Co-presenter linkedin 1',
    'Co-Presenter 2', 'Co-presenter email 2', 'Co-presenter linkedin 2',
    'Co-Presenter 3', 'Co-presenter email 3', 'Co-presenter linkedin 3',
    'Co-Presenter 4', 'Co-presenter email 4', 'Co-presenter linkedin 4'
  ];

  // Append each row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    body.appendParagraph('────────────────────────────').setBold(true);
    body.appendParagraph(`📌 Session ${i}`).setHeading(DocumentApp.ParagraphHeading.HEADING2);

    headers.forEach((field, idx) => {
      const value = row[idx];
      const valueStr = value.toString().trim();

      if (skipIfEmpty.includes(field) && valueStr === '') return;

      body.appendParagraph(field).setBold(true);
      body.appendParagraph(valueStr).setBold(false);
      body.appendParagraph('');
    });

    body.appendPageBreak(); // Add a page break after each session
  }

  doc.saveAndClose();
}

function updateFormWithSessionTitlesConfirmYYC() {
  const formUrl = 'https://docs.google.com/forms/d/1IpH0stJpXEABjBfySE_UjYgdu6rymq-Lo21jp_g23IY/edit';
  const sheet = SpreadsheetApp.openById('1y8klx-mnV6QEKnwIeLbkwBqkY2v8i5smgY9dIuIXxgc').getSheetByName('SessionsYYC');
  const form = FormApp.openByUrl(formUrl);

  const workshopTitles = sheet.getRange('E2:E' + sheet.getLastRow())
    .getValues()
    .flat()
    .filter(title => title && title.toString().trim() !== '');

  const rows = [
    'Interest',
    'Originality & Depth',
    'Quality & Clarity',
    'Experience & Presence'
  ];

  const columns = ['1', '2', '3', '4', '5'];

  workshopTitles.forEach(title => {
    const questionTitle = `Rate: ${title}`;
    const existingItem = form.getItems(FormApp.ItemType.GRID).find(item => item.getTitle() === questionTitle);

    if (!existingItem) {
      const gridItem = form.addGridItem().setTitle(questionTitle);
      gridItem.setRows(rows).setColumns(columns);
    }
  });
}

function updateFormWithSessionTitlesConfirmYEG() {
  const formUrl = 'https://docs.google.com/forms/d/1sQFhSWA0WQG96AEfw76SNSlQxJ5kuRn55up0PNcDzSo/edit';
  const sheet = SpreadsheetApp.openById('1y8klx-mnV6QEKnwIeLbkwBqkY2v8i5smgY9dIuIXxgc').getSheetByName('SessionsYEG');
  const form = FormApp.openByUrl(formUrl);

  const workshopTitles = sheet.getRange('E2:E' + sheet.getLastRow())
    .getValues()
    .flat()
    .filter(title => title && title.toString().trim() !== '');

  const rows = [
    'Interest',
    'Originality & Depth',
    'Quality & Clarity',
    'Experience & Presence'
  ];

  const columns = ['1', '2', '3', '4', '5'];

  workshopTitles.forEach(title => {
    const questionTitle = `Rate: ${title}`;
    const existingItem = form.getItems(FormApp.ItemType.GRID).find(item => item.getTitle() === questionTitle);

    if (!existingItem) {
      const gridItem = form.addGridItem().setTitle(questionTitle);
      gridItem.setRows(rows).setColumns(columns);
    }
  });
}