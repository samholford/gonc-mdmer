

(function () {
  if (!window.FileReader || !window.ArrayBuffer) {
    showError('This browser does not support the List Checker');
  }

  const dropArea = document.getElementById('mdmFileDropArea');
  var mdmList = [], mdmFiles = [];
  var listInputText = '';
  var fileInputText, fileInputTextVersion, filesToProcess;

  // EVENT HANDLING
  document.getElementById("mdmListInput").addEventListener("keyup", readListFromInput, false);
  document.getElementById("mdmListInput").addEventListener("change", readListFromInput, false);

  dropArea.addEventListener('dragover', function(event) {
    event.stopPropagation();
    event.preventDefault();
    // Style the drag-and-drop as a "copy file" operation.
    event.dataTransfer.dropEffect = 'copy';
  });

  dropArea.addEventListener('drop', function(event) {
    event.stopPropagation();
    event.preventDefault();
    resetErrors();
    filesToProcess = event.dataTransfer.files.length;
    fileInputTextVersion = 0;
    fileInputText = '';
    readFiles(event, callback);
  });

  function callback(result) {
    filesToProcess--;
    if (filesToProcess == 0) {
      if (listInputText == '' && fileInputTextVersion > 0) {
        // MDM list not supplied by input box and an MDM list was dragged in
        var list = readListFromFile(fileInputText);
        mdmList = parseList(list);
      }
      var rows = compare();
      printTable(rows);
    }
  }

  function readListFromInput() {
    var list = document.getElementById("mdmListInput").value.trim();
    if (list.length > 50 && list != listInputText) {
      listInputText = list;
      // Clear input
      document.getElementById("mdmListInput").value = '';
      resetErrors();
      mdmList = parseList(list);
      var rows = compare();
      printTable(rows);
    }
  }

  function readListFromFile(content) {
    var list = '';

    function stripXML(xml) {
      var result = '';
      for (var j = 0; j < xml.length; j++) {
        var stripped = xml[j].replace(/<w:t>/, "");
        stripped = stripped.replace(/<\/w:t>/, "");
        result += stripped + " ";
      }
      // Split by NHI
      var patientParts = result.split(/([A-Z]{3}\d{4})/);

      // Find comma
      var commaIndex = patientParts[0].indexOf(",");

      // Remove text after the comma
      if (commaIndex != -1) {
        patientParts[0] = patientParts[0].substring(0, commaIndex);
      }

      patientParts[0] = patientParts[0].trim();

      return patientParts[0] + " " + patientParts[1];
    }

    // remove xml:space="preserve" from text
    content = content.replace(/ xml:space="preserve"/g, "");

    var patients = content.split('w:val="1"');

    // For each patient
    for (var i = 1; i < patients.length; i++) {
      // Match everything between '<w:t>' tags
      var tags = patients[i].match(/(<w:t>)([\s\S]*?)(<\/w:t>)/g);

      if (tags) {
        // Strip tags and append
        list += stripXML(tags);
      }

    }
    return list;
  }

  function parseList(list) {
    var cleanString = list.replace(/NDHB|BOPDHB|CMDHB|ADHB|Waikato|LDHB|WDHB|Pvt|DHB/g, "");
    // Split list by NHIs
    // Note split keeps the NHI as every second element e.g. ['raw text', 'NHI', 'raw text', 'NHI']
    var matches = cleanString.split(/([A-Z]{3}\d{4})/);
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
      var nhi = matches[i+1];

      // Split line by number (DD.)
      var lineSplit = line.split(/\d{1,2}\./);

      // Keep last bit (text after last number)
      var name = (lineSplit.length > 1) ? lineSplit[lineSplit.length-1] : line;

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
    var matched = [], rows = [];
    for (var i = 0; i <mdmList.length; i++) {
      var template = 'none';
      for (var j = 0; j<mdmFiles.length; j++) {
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
    for (var i = 0; i<mdmFiles.length; i++) {
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
    var i;
    var ready = true;
    var table = document.getElementById("checker-output").getElementsByTagName("tbody")[0];

    // Clear table in DOM
    table.innerHTML = "";

    function addCell(row, content, index) {
      // Insert a cell in the row at index 0
      var newCell = row.insertCell(index);

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
          var newText = document.createTextNode(content);
          // Append a text node to the cell
          newCell.appendChild(newText);
      }
    }

    for (i = 0; i<rows.length; i++) {
      // Insert a row in the table at row index 0
      var newRow = table.insertRow(-1);

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


  function readFiles(event, callback) {
    // Reset
    mdmFiles = [];

    var files = event.dataTransfer.files;
    for (var i = 0, len = files.length; i < len; i++) {
      var f = files[i];
      var reader = new FileReader();

      // Closure to capture the file information.
      reader.onload = (function(theFile) {
        return function(e) {
          // theFile.name
          try {
            // read the content of the file with PizZip
            var zip = new PizZip(e.target.result);
            var property;

            // Cycle through each file contained with the docx container
            for (property in zip.files) {
              var zipEntry = zip.files[property];
              // Docx is a zip file of many xml files, we only want 'word/document.xml'
              // Also, exclude MDM lists (contain MDM in file name)
              if (zipEntry.name == 'word/document.xml') {
                var text = zipEntry.asText();
                var fileName = theFile.name.replace(".docx", "");
                if (theFile.name.indexOf("MDM") != -1) {
                  // MDM list
                  var listNumber = parseInt(fileName.substring(0,2));
                  if (listNumber > fileInputTextVersion) {
                    // More recent version found
                    fileInputTextVersion = listNumber;
                    fileInputText = text;
                  }
                } else {
                  // Patient
                  var nhi = text.match(/[A-Z]{3}[0-9]{4}/g)[0];
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
              }
            }
          } catch(e) {
            showError('Error reading ' + theFile.name + ' : ' + e.message);
          }

          //Called at end of each asynchronous read of file
          // https://stackoverflow.com/questions/30312894/filereader-and-callbacks
          callback('foo');
        }
      })(f);

      // read the file !
      // readAsArrayBuffer and readAsBinaryString both produce valid content for PizZip.
      reader.readAsArrayBuffer(f);
      // reader.readAsBinaryString(f);
    }
  }
})();
