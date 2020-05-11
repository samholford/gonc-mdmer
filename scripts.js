// If a 2-digit year is used, assume less than 100 years old
moment.parseTwoDigitYear = function(yearString) {
  return parseInt(yearString) + (parseInt(yearString) > 20 ? 1900 : 2000);
};

function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback);
}
function generate() {
  const DEBUG = false;

  // Declare referral data
  var raw = document.getElementById("referralRawText").value;
  var Referral = {
    referralDate: "",
    nhi: "",
    dob: "",
    patientName: "",
    address: "",
    phone: "",
    referrerName: "",
    referrerDHB: "",
    referrerEmail: "",
    gp: "",
    age: "",
    history: "",
    comorbidities: "",
    markers: "",
    bmi: "",
    ecog: "",
    ethnicity: "",
    hasRad: false,
    radiology: [],
    hasOp: false,
    operation: [],
    hasHisto: false,
    histology: [],
    question: ""
  };

  // Text extraction function
  function getText(start, end, fuzzyStart, date) {
    fuzzyStart = fuzzyStart || false;
    date = date || false;
    var startPos = raw.toLowerCase().search(start.toLowerCase());
    var notProvided = "Not provided";

    function include(arr, obj) {
      return arr.indexOf(obj) != -1;
    }

    if (include(["Date", "Type", "Procedure", "Location", "Surgeon"], start)) {
      if (start == "Surgeon" || start == "Location") {
        notProvided = "";
      } else {
        notProvided = start + " not provided";
      }
    }

    if (startPos == -1) {
      // Not found
      return notProvided;
    }

    startPos += start.length;

    raw = raw.substring(startPos); // Delete prior text

    if (fuzzyStart) {
      // Start from the first colon to allow variations on label
      raw = raw.substring(raw.indexOf(":") + 1);
    } else if (raw.substring(0, 1) == ":") {
      // Remove leading colon if there
      raw = raw.substring(1);
    }

    if (typeof end == "string") {
      end = end.toLowerCase();
    }
    var endPos = raw.toLowerCase().search(end);
    if (endPos == -1) {
      endPos = raw.toLowerCase().search(/\n/g) || []; // End at the next new line
    }
    if (endPos == -1) {
      return "ERROR!"; // Can't find the end
    }

    var result = raw.substring(0, endPos);

    // Remove new lines, \t, bullets
    // Bullets https://stackoverflow.com/questions/18266529/how-to-write-regex-for-bullet-space-digit-and-dot/18266778
    result = result
      .replace(/\r?\n|\r|\t/g, " ") // new lines
      .replace(/[\u2022\u2023\u25E6\u2043\u2219]/g, "") // bullets
      .trim();

    if (result) {
      if (date) {
        var date = moment.utc(result, [
          "DD-MM-YYYY",
          "DD/MM/YYYY",
          "DD-MMMM-YYYY",
          "DD/MMMM/YYYY",
          "DD-MMM-YYYY",
          "DD/MMM/YYYY",
          "Do-MMMM-YYYY",
          "Do/MMMM/YYYY",
          "DD-MM-YY",
          "DD/MM/YY",
          "DD-M-YY",
          "DD/M/YY",
          "D-M-YY",
          "D/M/YY"
        ]);
        if (date.isValid()) {
          return date.format("DD/MM/YYYY");
        }
      }
      if (notProvided != "") {
        // Check if needing special output
        return result;
      } else {
        // Surgeon or location type, gets different output
        if (start == "Surgeon") {
          return "(" + result + ")"; // Add brackets
        } else {
          return ", " + result; // Add leading comma
        }
      }
    } else {
      return notProvided; // Blank result
    }
  }


  // Link to the template docx - see https://docxtemplater.readthedocs.io/en/latest/tag_types.html
  loadFile("https://samholford.github.io/gonc-mdmer/newPatientTemplate.docx", function(
    error,
    content
  ) {
    if (error) {
      throw error;
    }

    // The error object contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
    function replaceErrors(key, value) {
      if (value instanceof Error) {
        return Object.getOwnPropertyNames(value).reduce(function(error, key) {
          error[key] = value[key];
          return error;
        }, {});
      }
      return value;
    }

    function errorHandler(error) {
      console.log(JSON.stringify({ error: error }, replaceErrors));

      if (error.properties && error.properties.errors instanceof Array) {
        const errorMessages = error.properties.errors
          .map(function(error) {
            return error.properties.explanation;
          })
          .join("\n");
        console.log("errorMessages", errorMessages);
        document.getElementById("errors").innerHTML = errorMessages;
        document.getElementById("error-container").style.display = "block";
        // errorMessages is a humanly readable message looking like this :
        // 'The tag beginning with "foobar" is unopened'
      }
      throw error;
    }

    var zip = new PizZip(content);
    var doc;
    try {
      doc = new window.docxtemplater(zip);
    } catch (error) {
      // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
      errorHandler(error);
    }


    // Extract referral content
    Referral["referralDate"] = getText("Date of Referral", "NHI Number", false, true);
    Referral["nhi"] = getText("NHI Number", "Patient Name");
    Referral["patientName"] = getText("Patient Name", "DOB");
    Referral["dob"] = getText("DOB", "Address", false, true);
    Referral["address"] = getText("Address", "Phone/");
    Referral["phone"] = getText("Mobile", /nb|high/); // /\r?\n|\r/

    Referral["referrerName"] = getText("Consultant name", "Hospital/DHB", true);
    Referral["referrerDHB"] = getText("Hospital/DHB", "Email address", true);
    Referral["referrerEmail"] = getText("Email address", "GP name", true);

    Referral["gp"] = getText("GP name and address", "HISTORY", true);

    Referral["age"] = getText("Age", "Brief history", true);
    Referral["history"] = getText("Brief History", /(co morbidities|co-morbidities)/, true);
    Referral["comorbidities"] = getText("morbidities", "Tumour", true);
    Referral["markers"] = getText("markers", "BMI", true);
    Referral["bmi"] = getText("BMI", "ECOG", true);
    Referral["ecog"] = getText("ECOG", "Ethnicity", true);
    Referral["ethnicity"] = getText("Ethnicity", "RADIOLOGY", true);

    var split = "WHAT IS THE QUESTION"; // Default end
    var holder = raw;
    var i, len, loop, obj;

    if (raw.search("RADIOLOGY") > -1) {
      Referral["hasRad"] = true;
    }
    if (raw.search("OPERATION") > -1) {
      Referral["hasOp"] = true;
    }
    if (raw.search("HISTOLOGY") > -1) {
      Referral["hasHisto"] = true;
    }

    if (Referral["hasRad"]) {
      // Determine the end of the loop
      if (Referral["hasOp"]) {
        split = "OPERATION";
      } else if (Referral["hasHisto"]) {
        split = "HISTOLOGY";
      }

      // Make loop hold just the radiology text
      loop = raw
        .split("RADIOLOGY")[1]
        .split(split)[0]
        .trim();
      loop = loop.substring(loop.indexOf("Type"));
      // Make loop an array of each imaging
      loop = loop.split("Type");

      len = loop.length;
      for (i = 1; i < len; i++) {
        raw = "Type" + loop[i] + "LOOPEND";
        obj = {
          radType: getText("Type", "Date", true),
          radDate: getText("Date", "Location", true, true),
          radDHB: getText("Location", "Key findings", true),
          radFindings: getText("Findings", "LOOPEND", true)
        };
        if (
          obj["radType"] != "Type not provided" ||
          obj["radFindings"] != "Not provided"
        ) {
          Referral["radiology"].push(obj);
        }
      }
      Referral["hasRad"] = !(Referral["radiology"].length == 0); // Remove radiology section if array empty
      raw = holder; // Restore raw
    }

    if (Referral["hasOp"]) {
      // Determine the end of the loop
      if (Referral["hasHisto"]) {
        split = "HISTOLOGY";
      }

      // Make loop hold just the operation text
      loop = raw
        .split("OPERATION")[1]
        .split(split)[0]
        .trim();
      loop = loop.substring(loop.indexOf("Date"));
      // Make loop an array of each operation
      loop = loop.split("Date");

      //Referral['hasOp'] = false // Will check input data in loop below
      len = loop.length;
      for (i = 1; i < len; i++) {
        raw = "Date" + loop[i] + "LOOPEND";
        obj = {
          opDate: getText("Date", "Surgeon", true, true),
          opSurgeon: getText("Surgeon", "Procedure", true),
          opType: getText("Procedure", "Findings", true),
          opFindings: getText("Findings", "LOOPEND", true)
        };
        if (
          obj["opType"] != "Procedure not provided" ||
          obj["opFindings"] != "Not provided"
        ) {
          Referral["operation"].push(obj);
        }
      }
      Referral["hasOp"] = !(Referral["operation"].length == 0); // Remove operation section if array empty
      raw = holder; // Restore raw
    }

    if (Referral["hasHisto"]) {
      // Make loop hold just the histology text
      loop = raw
        .split("HISTOLOGY")[1]
        .split("WHAT IS THE QUESTION")[0]
        .trim();
      //loop = loop.substring(loop.indexOf('type'))
      // Make loop an array of each specimen
      if (loop.search("Specimen type") > 0) {
        loop = loop.split("Specimen type");
      } else {
        loop = loop.split("Histology type");
      }

      len = loop.length;
      for (i = 1; i < len; i++) {
        raw = "type" + loop[i] + "LOOPEND";
        obj = {
          histoType: getText("Type", "Date", true),
          histoDate: getText("Date", "Location", true, true),
          histoDHB: getText("Location", "Key findings", true),
          histoFindings: getText("Findings", "LOOPEND", true)
        };
        if (
          obj["histoType"] != "Type not provided" ||
          obj["histoFindings"] != "Not provided"
        ) {
          Referral["histology"].push(obj);
        }
      }
      Referral["hasHisto"] = !(Referral["histology"].length == 0); // Remove histology section if array empty
      raw = holder; // Restore raw
    }

    Referral["question"] = getText("FOR THE MDM?", /is the patient|what does the patient know/);


    if (DEBUG) {
      console.log(Referral);
    }

    // set the template variables from the Referral
    doc.setData(Referral);


    try {
      // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
      doc.render();
    } catch (error) {
      // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
      errorHandler(error);
    }

    var firstName = Referral["patientName"].split(/\s(.+)/)[0].toLowerCase(); //everything before the first space
    var lastName = Referral["patientName"].split(/\s(.+)/)[1];
    firstName = firstName[0].toUpperCase() + firstName.slice(1); // Convert firstname is sentence case
    var fileName = lastName.toUpperCase() + " " + firstName + ".docx";

    var out = doc.getZip().generate({
      type: "blob",
      compression: "DEFLATE",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    }); //Output the document using Data-URI

    document.getElementById("message").innerHTML = 'You just created<br />' + fileName;

    if (!DEBUG) {
      saveAs(out, fileName);
    }
  });
}

document.getElementById("submitReferral").addEventListener("click", function() {
  document.getElementById("errors").innerHTML = "";
  document.getElementById("error-container").style.display = "none";
  generate();
  document.getElementById("referralRawText").placeholder = 'Paste another one!'
  document.getElementById("referralRawText").value = '';
  document.getElementById("submitReferral").disabled = true;
  document.getElementById("submitReferral").classList.add("is-disabled")
});

function toggleButton() {
  var name = document.getElementById("referralRawText").value;
  if (name.length > 3) {
    document.getElementById("submitReferral").disabled = false;
    document.getElementById("submitReferral").classList.remove("is-disabled")    
  } else {
    document.getElementById("submitReferral").disabled = true;
    document.getElementById("submitReferral").classList.add("is-disabled")
  }
}

document.getElementById("referralRawText").addEventListener("keyup", toggleButton, false);
document.getElementById("referralRawText").addEventListener("change", toggleButton, false);