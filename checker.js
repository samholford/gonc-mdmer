

(function () {
  if (!window.FileReader || !window.ArrayBuffer) {
    showError('This browser does not support the List Checker');
  }

  const dropArea = document.getElementById('mdmFileDropArea');
  var mdmList = [], mdmFiles = [], rows = [];
  var rawText = '';
  var filesToProcess = 0;

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
    readFiles(event, callback);
  });

  function callback(result) {
    filesToProcess--;
    if (filesToProcess == 0) {
      compare();
      printTable();
    }
  }

  document.getElementById("mdmListInput").addEventListener("keyup", readList, false);
  document.getElementById("mdmListInput").addEventListener("change", readList, false);

  function readList() {
    var listText = document.getElementById("mdmListInput").value.trim();
    if (listText.length > 50 && listText != rawText) {
      rawText = listText;
      // Clear input
      document.getElementById("mdmListInput").value = '';
      resetErrors();
      parseList();
      compare();
      printTable(rows);
    }
  }

  function parseList() {
    // Reset
    mdmList = [];

    // Split list by NHIs
    // Note split keeps the NHI as every second element e.g. ['raw text', 'NHI', 'raw text', 'NHI']
    var matches = rawText.split(/([A-Z]{3}\d{4})/);

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

      mdmList.push([name, nhi]);
      // Increment index again so skip NHI
      i++;
    }
  }

  function compare() {
    var i, name, nhi, template, matched =[];
    //Reset
    rows = [];
    for (i = 0; i <mdmList.length; i++) {
      var j;
      var listed = true;
      name = mdmList[i][0];
      nhi = mdmList[i][1];
      template = 'none';
      for (j= 0; j<mdmFiles.length; j++) {
        if (mdmList[i][1] == mdmFiles[j][1]) {
          // NHI match
          name = mdmFiles[j][0];
          nhi = mdmFiles[j][1];
          template = mdmFiles[j][2]; // 'complete' or 'incomplete'
          //files.splice(j,1); // remove from array
          matched.push(nhi);
          rows.push([name, nhi, listed, template]);
          break;
        }
      }
      if (template == 'none') {
        rows.push([name, nhi, listed, template]);
      }
    }
    for (i= 0; i<mdmFiles.length; i++) {
      name = mdmFiles[i][0];
      nhi = mdmFiles[i][1];
      template = mdmFiles[i][2]; // 'complete' or 'incomplete'
      if (matched.indexOf(nhi) == -1) {
        rows.push([name, nhi, false, template]);
      }
    }
  }

  function printTable() {
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
          var newText  = document.createTextNode(content);
          // Append a text node to the cell
          newCell.appendChild(newText);
      }
    }

    for (i=0; i<rows.length; i++) {
      // Insert a row in the table at row index 0
      var newRow = table.insertRow(-1);

      addCell(newRow, i + 1 + '.', 0);
      addCell(newRow, rows[i][0], 1);
      addCell(newRow, rows[i][1], 2);

      // Listed for MDM star
      if (rows[i][2] == true) {
        addCell(newRow, 'star', 3);
      } else {
        addCell(newRow, 'nostar', 3);
        ready = false;
      }

      // Templated MDM star
      switch(rows[i][3]) {
        case 'none':
          addCell(newRow, 'nostar', 4);
          addCell(newRow, 'Template missing', 5);
          ready = false;
          break;
        case 'incomplete':
          addCell(newRow, 'halfstar', 4);
          ready = false;
          if (rows[i][2]) {
            // Only add this status if listed for MDM, otherwise status will become 'Not listed for MDM' instead
            addCell(newRow, 'Template incomplete', 5);
          }
          break;
        case 'complete':
          addCell(newRow, 'star', 4);
          break;
      }

      // Apply status colours
      if (rows[i][2] == true && rows[i][3] == 'complete') {
        newRow.classList.add("row-green");
        addCell(newRow, 'Done', 5);
      } else if (rows[i][2] == false) {
        ready = false;
        newRow.classList.add("row-red");
        addCell(newRow, 'Not listed for MDM', 5);
        if (rows[i][4] == 'complete') {
          addCell(newRow, 'star', 4);
        }
      } else if (rows[i][3] == 'none') {
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

            // Cycle through each file contained with the docx container
            $.each(zip.files, function (index, zipEntry) {
              // Docx is a zip file of many xml files, we only want 'word/document.xml'
              // Also, exclude MDM lists (contain MDM in file name)
              if (zipEntry.name == 'word/document.xml' && theFile.name.indexOf("MDM") == -1) {
                var rawText = zipEntry.asText();
                var nhi = rawText.match(/[A-Z]{3}[0-9]{4}/g)[0];
                var fileName = theFile.name.replace(".docx", "");
                // Check for # in filename indicated incomplete
                if (fileName.substring(0,1) == '#') {
                  mdmFiles.push([fileName.substring(1), nhi, 'incomplete']);
                } else {
                  mdmFiles.push([fileName, nhi, 'complete']);
                }
              }
            });

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
