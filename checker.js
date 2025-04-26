const dropArea = document.getElementById('mdmFileDropArea');
const nhiPattern = /([A-Z]{3}\d{2}[A-Z0-9]{2})/g;
let mdmList = [], mdmFiles = [];
let listInputText = '';
let mdmListText, rawText, fileInputTextVersion, filesToProcess;

// Prevent default drag behaviours
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropArea.addEventListener(eventName, e => e.preventDefault(), false);
    dropArea.addEventListener(eventName, e => e.stopPropagation(), false);
});

dropArea.addEventListener('drop', async function(event) {
    const files = event.dataTransfer.files;
    const processingPromises = [];
    
    resetErrors();
    filesToProcess = event.dataTransfer.files.length;
    fileInputTextVersion = 0;
    rawText = '';
    mdmFiles = [];

    for (const file of files) {
        const arrayBuffer = await file.arrayBuffer();
        const processingPromise = mammoth.extractRawText({arrayBuffer: arrayBuffer})
        .then(function(result){
            rawText = result.value; // The raw text
            var messages = result.messages;
            
            let fileName = file.name.replace('.docx', '');
            if (fileName.indexOf('MDM') != -1) {
              // MDM list
              let listNumber = parseInt(fileName.substring(0,2));
              if (listNumber > fileInputTextVersion) {
                // More recent version found
                fileInputTextVersion = listNumber;
                mdmListText = rawText;
              }
            } else {
              // Patient
              //let nhi = fileName.match(nhiPattern)[0]; // Gets NHI from file name instead of docx content
              let nhi = rawText.match(nhiPattern)[0];
              fileName = fileName.replace(nhiPattern, '');
              // Check for # in filename indicated incomplete
              if (fileName.substring(0,1) == '#') {
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
      
      // Collect promises
      processingPromises.push(processingPromise);  
    }
    
    // Wait for all mammoth.extractRawText() promises to finish before continuing
    await Promise.all(processingPromises);
    
    mdmList = parseList(rawText);
    let rows = compare();
    printTable(rows);
    
});

function parseList(list) {
  var cleanString = list.replace(/NDHB|BOPDHB|CMDHB|ADHB|Waikato|LDHB|WDHB|Pvt|DHB/g, '');
  cleanString = cleanString.replace(/^[\s\S]*?\bProblem\b[\s\r\n]*\bHistology\b[\s\r\n]*\bRadiology\b[\s\r\n]*/i, '\n\n\n\n'); // Replace preamble with quad line breaks
  
  const keepAfterLastQuadrupleLineBreak = text => {
    const parts = text.split(/(?:\r?\n){4}/g).map(p => p.trim()).filter(Boolean);
    return parts.length > 0 ? parts[parts.length - 1] : text;
  };
  
  // Split list by NHIs
  // Note split keeps the NHI as every second element e.g. ['raw text', 'NHI', 'raw text', 'NHI']
  var matches = cleanString.split(nhiPattern);
  var parsedList = [];
  if (!matches) {
    showError('Failed to parse the MDM list');
    return;
  }

  // Loop through array. i+1 so that we ignore the last bit of text after the last NHI
  for (var i = 0; i+1 < matches.length; i++) {
    // Get all text between NHIs
    var line = matches[i];
   
    // Get NHI
    var nhi = matches[i+1].replace(/\s+/g, ''); // Remove any white space
    
    // Split line by quadruple line break
    //var name = line.split(/\r?\n\r?\n\r?\n\r?\n/).slice(1).join('');
    
    var name = keepAfterLastQuadrupleLineBreak(line);
    
    // Find comma
    var commaIndex = name.indexOf(",");

    // Remove text after the comma
    if (commaIndex != -1) {
      name = name.substring(0, commaIndex);
    }

    name = name.trim();
    
    parsedList.push({
      name: name,
      nhi: nhi
    });
    // Increment index again so skip NHI
    i++;
  }
  return parsedList;
}

function compare() {
  let matched = [], rows = [];
  for (let i = 0; i <mdmList.length; i++) {
    let template = 'none';
    for (let j = 0; j<mdmFiles.length; j++) {
      if (mdmList[i].nhi == mdmFiles[j].nhi) {
        // NHI match
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
  let table = document.getElementById("checker-output").getElementsByTagName("tbody")[0];
  
  // Clear table in DOM
  table.innerHTML = '';

  function addCell(row, content, index) {
    // Insert a cell in the row at index 0
    let newCell = row.insertCell(index);

    // Left align the first two columns
    if (index <2) {
      newCell.classList.add("left-align");
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
        // Append a text node to the cell
        newCell.appendChild(newText);
    }
  }

  for (i = 0; i<rows.length; i++) {
    // Insert a row in the table at row index 0
    let newRow = table.insertRow(-1);

    addCell(newRow, i + 1 + '.', 0);
    addCell(newRow, rows[i].name, 1);
    addCell(newRow, rows[i].nhi, 2);

    // Listed for MDM star
    if (rows[i].listed) {
      addCell(newRow, 'star', 3);
    } else {
      addCell(newRow, 'nostar', 3);
      ready = false;
    }

    // Templated MDM star
    switch(rows[i].status) {
      case 'none':
        addCell(newRow, 'nostar', 4);
        addCell(newRow, 'Template missing', 5);
        ready = false;
        break;
      case 'incomplete':
        addCell(newRow, 'halfstar', 4);
        ready = false;
        if (rows[i].listed) {
          // Only add this status if listed for MDM, otherwise status will become 'Not listed for MDM' instead
          addCell(newRow, 'Template incomplete', 5);
        }
        break;
      case 'complete':
        addCell(newRow, 'star', 4);
        break;
    }

    // Apply status colours
    if (rows[i].listed && rows[i].status == 'complete') {
      newRow.classList.add("row-green");
      addCell(newRow, 'Done', 5);
    } else if (!rows[i].listed) {
      ready = false;
      newRow.classList.add("row-red");
      addCell(newRow, 'Not listed for MDM', 5);
    } else if (rows[i].status == 'none') {
      ready = false;
      newRow.classList.add("row-red");
    } else {
      ready = false;
      newRow.classList.add("row-orange");
    }
  }
  
  // Show the table
  if (document.getElementById("checker-output").getElementsByTagName("td").length > 0) {
    document.getElementById("checker-output").classList.remove("hidden");
    if (ready) {
      document.getElementById("checker-section").classList.remove("hidden");
    }
  } else {
    document.getElementById("checker-section").classList.add("hidden");
    document.getElementById("checker-output").classList.add("hidden");
  }
}