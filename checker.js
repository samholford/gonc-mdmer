const dropArea = document.getElementById('mdmFileDropArea');
const nhiPattern = /(?<![A-Z0-9])([A-Z]{3}\d{2}[A-Z0-9]{2})(?![A-Z0-9])/g; // Uses negative lookbehind and lookahead to exclude if a character immediately contacts the NHI pattern

// Prevent default drag behaviours
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropArea.addEventListener(eventName, e => e.preventDefault(), false);
    dropArea.addEventListener(eventName, e => e.stopPropagation(), false);
});

dropArea.addEventListener('drop', async function(event) {
    const files = event.dataTransfer.files;
    const processingPromises = [];
    const mdmFiles = [];
    let fileInputTextVersion = 0;
    let mdmListText = '';
    
    resetErrors();

    for (const file of files) {
        const arrayBuffer = await file.arrayBuffer();
        const processingPromise = mammoth.extractRawText({arrayBuffer: arrayBuffer})
        .then(function(result){
            const rawText = result.value; // The raw text
            //console.log(result.messages);
            let fileName = file.name.replace('.docx', '');
            if (fileName.indexOf('MDM') != -1) { // MDM list
              let listNumber = parseInt(fileName.substring(0,2));
              if (listNumber > fileInputTextVersion) { // More recent version found
                fileInputTextVersion = listNumber;
                mdmListText = rawText;
              }
            } else {  // Patient
              let nhi = rawText.match(nhiPattern)[0];
              fileName = fileName.replace(nhiPattern, '');
              if (fileName.substring(0,1) == '#') { // Incomplete template
                fileName = fileName.substring(1);
                mdmFiles.push({
                  name: fileName,
                  nhi: nhi,
                  status: 'incomplete'
                });
              } else {
                mdmFiles.push({
                  name: fileName,
                  nhi: nhi,
                  status: 'complete'
                });
              }
            }
      })
      .catch(function(error) {
          console.error(error);
      }); 
      
      processingPromises.push(processingPromise); // Collect promises for await process below
    }
    
    await Promise.all(processingPromises); // Wait for all mammoth.extractRawText() promises to finish before continuing
    const mdmList = parseList(mdmListText);
    let rows = compare(mdmList, mdmFiles);
    printTable(rows);
});

function parseList(list) {
  let cleanString = list.replace(/NDHB|BOPDHB|CMDHB|ADHB|Waikato|LDHB|WDHB|Pvt|DHB/g, '');
  cleanString = cleanString.replace(/^[\s\S]*?\bProblem\b[\s\r\n]*\bHistology\b[\s\r\n]*\bRadiology\b[\s\r\n]*/i, '\n\n\n\n'); // Replace preamble with quad line breaks
  const keepAfterLastQuadrupleLineBreak = text => {
    const parts = text.split(/(?:\r?\n){4}/g).map(p => p.trim()).filter(Boolean);
    return parts.length > 0 ? parts[parts.length - 1] : text;
  };
  
  // Split list by NHIs resulting in ['raw text', 'NHI', 'raw text', 'NHI']
  let matches = cleanString.split(nhiPattern);
  let parsedList = [];
  if (!matches) {
    showError('Failed to parse the MDM list');
    return;
  }

  // Loop through array. i+1 so that we ignore the last bit of text after the last NHI
  for (let i = 0; i+1 < matches.length; i++) {
    let line = matches[i]; // Raw text between NHIs
    let nhi = matches[i+1].replace(/\s+/g, ''); // NHI with any white space removed    
    
    let name = keepAfterLastQuadrupleLineBreak(line);
    let commaIndex = name.indexOf(',');
    name = (commaIndex != -1) ? name.substring(0, commaIndex) : name; // Remove text after the comma
    name = name.trim();
    
    parsedList.push({
      name: name,
      nhi: nhi
    });
    i++; // Increment index again to skip NHI
  }
  
  return parsedList;
}

// Compares the MDM list against reports and adds completion status. 
// Identifies if listed reports are missing or if there are reports present that are not listed.
function compare(mdmList, mdmFiles) {
  let matched = [], rows = [];
  for (let i = 0; i <mdmList.length; i++) {
    let template = 'none';
    for (let j = 0; j<mdmFiles.length; j++) {
      if (mdmList[i].nhi == mdmFiles[j].nhi) { // NHI match
        template = mdmFiles[j].status; // 'incomplete' or 'complete'
        matched.push(mdmList[i].nhi);
        rows.push({
          name: mdmFiles[j].name,
          nhi: mdmFiles[j].nhi,
          listed: true,
          status: mdmFiles[j].status
        });
        break;
      }
    }
    if (template == 'none') {
      rows.push({
        name: mdmList[i].name,
        nhi: mdmList[i].nhi,
        listed: true,
        status: 'none'
      });
    }
  }
  for (let i = 0; i<mdmFiles.length; i++) {
    if (matched.indexOf(mdmFiles[i].nhi) == -1) {
      rows.push({
        name: mdmFiles[i].name,
        nhi: mdmFiles[i].nhi,
        listed: false,
        status: mdmFiles[i].status
      });
    }
  }
  return rows;
}

function printTable(rows) {
  let i;
  let ready = true;
  let table = document.getElementById('checker-output').getElementsByTagName('tbody')[0];
  
  table.innerHTML = ''; // Clear table in DOM

  function addCell(row, content, index) {
    let newCell = row.insertCell(index); // Insert a cell in the row at index 0

    if (index <2) { // Left align the first two columns
      newCell.classList.add('left-align');
    }
    
    switch(content) {
      case 'star':
        newCell.innerHTML = '<i class="nes-icon star"></i>';
        break;
      case 'halfstar':
        newCell.innerHTML = '<i class="nes-icon star is-half"></i>';
        break;
      case 'nostar':
        newCell.innerHTML = '<i class="nes-icon star is-empty"></i>';
        break;
      default:
        let newText = document.createTextNode(content);
        newCell.appendChild(newText); // Append a text node to the cell
    }
  }

  for (let i = 0; i<rows.length; i++) {
    let newRow = table.insertRow(-1); // Insert a row in the table at row index 0

    addCell(newRow, i + 1 + '.', 0);
    addCell(newRow, rows[i].name, 1);
    addCell(newRow, rows[i].nhi, 2);

    if (rows[i].listed) { // Listed for MDM star
      addCell(newRow, 'star', 3);
    } else {
      addCell(newRow, 'nostar', 3);
      ready = false;
    }

    switch(rows[i].status) { // Templated MDM star
      case 'none':
        addCell(newRow, 'nostar', 4);
        addCell(newRow, 'Template missing', 5);
        ready = false;
        break;
      case 'incomplete':
        addCell(newRow, 'halfstar', 4);
        ready = false;
        if (rows[i].listed) { // Only if listed for MDM, otherwise status is 'Not listed for MDM'
          addCell(newRow, 'Template incomplete', 5);
        }
        break;
      case 'complete':
        addCell(newRow, 'star', 4);
        break;
    }

    // Apply status colours
    if (rows[i].listed && rows[i].status == 'complete') {
      newRow.classList.add('row-green');
      addCell(newRow, 'Done', 5);
    } else if (!rows[i].listed) {
      ready = false;
      newRow.classList.add('row-red');
      addCell(newRow, 'Not listed for MDM', 5);
    } else if (rows[i].status == 'none') {
      ready = false;
      newRow.classList.add('row-red');
    } else {
      ready = false;
      newRow.classList.add('row-orange');
    }
  }
  
  if (document.getElementById('checker-output').getElementsByTagName('td').length > 0) { // Show table
    document.getElementById('checker-output').classList.remove('hidden');
    if (ready) {
      document.getElementById('checker-section').classList.remove('hidden');
    }
  } else { // Hide table
    document.getElementById('checker-section').classList.add('hidden');
    document.getElementById('checker-output').classList.add('hidden');
  }
}