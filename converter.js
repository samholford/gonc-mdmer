function showError(msg) {
  const node = document.createElement('li');
  const textnode = document.createTextNode(msg);
  node.appendChild(textnode);
  document.getElementById('errors').appendChild(node);
  document.getElementById('error-container').classList.remove('hidden');
}

function resetErrors() {
  document.getElementById('errors').innHTML = '';
  document.getElementById('error-container').classList.add('hidden');
}

// If a 2-digit year is used, assume no more than 1 year in future
moment.parseTwoDigitYear = function(yearString) {
  const currentYear = moment().get('year') - 2000;
  return parseInt(yearString) + (parseInt(yearString) > ( currentYear + 1) ? 1900 : 2000);
};

function loadFile(url, callback) {
  PizZipUtils.getBinaryContent(url, callback);
}

function generate() {
  const DEBUG = false;
  resetErrors();

  let raw = document.getElementById('referralRawText').value;
  const Referral = {
    referralDate: '',
    nhi: '',
    dob: '',
    patientName: '',
    address: '',
    phone: '',
    referrerName: '',
    referrerDHB: '',
    referrerEmail: '',
    gp: '',
    age: '',
    history: '',
    comorbidities: '',
    markers: '',
    bmi: '',
    ecog: '',
    ethnicity: '',
    hasRad: false,
    radiology: [],
    hasOp: false,
    operation: [],
    hasHisto: false,
    histology: [],
    question: ''
  };

  function getText(start, end, fuzzyStart, date) {
    fuzzyStart = fuzzyStart || false;
    date = date || false;
    let startPos = raw.toLowerCase().search(start.toLowerCase());
    let notProvided = 'Not provided';

    function include(arr, obj) {
      return arr.indexOf(obj) != -1;
    }

    if (include(['Date', 'Type', 'Procedure', 'Location', 'Surgeon', 'Name of the lab'], start)) {
      if (start == 'Surgeon' || start == 'Location' || start == 'Name of the lab') {
        notProvided = '';
      } else {
        notProvided = start + ' not provided';
      }
    }

    if (startPos == -1) {
      return notProvided;
    }

    startPos += start.length;
    raw = raw.substring(startPos); // Delete prior text

    if (fuzzyStart) { // Start from the first colon to allow variations on label
      raw = raw.substring(raw.indexOf(':') + 1);
    } else if (raw.substring(0, 1) == ':') { // Remove leading colon if there
      raw = raw.substring(1);
    }

    if (typeof end == 'string') {
      end = end.toLowerCase();
    }
    let endPos = raw.toLowerCase().search(end);
    if (endPos == -1) {
      endPos = raw.toLowerCase().search(/\n/g) || []; // End at the next new line
    }
    if (endPos == -1) { // Can't find the end
      return 'ERROR!';
    }

    let result = raw.substring(0, endPos);

    // Remove new lines, \t, bullets (https://stackoverflow.com/questions/18266529/how-to-write-regex-for-bullet-space-digit-and-dot/18266778)
    result = result
      .replace(/\r?\n|\r|\t/g, ' ') // new lines
      .replace(/[\u2022\u2023\u25E6\u2043\u2219]/g, '') // bullets
      .trim();

    if (result) {
      if (date) {
        let date = moment.utc(result, [
          'DD-MM-YYYY',
          'DD/MM/YYYY',
          'DD-MMMM-YYYY',
          'DD/MMMM/YYYY',
          'DD-MMM-YYYY',
          'DD/MMM/YYYY',
          'Do-MMMM-YYYY',
          'Do/MMMM/YYYY',
          'DD-MM-YY',
          'DD/MM/YY',
          'DD-M-YY',
          'DD/M/YY',
          'D-M-YY',
          'D/M/YY'
        ]);
        if (date.isValid()) {
          return date.format('DD/MM/YYYY');
        }
      }
      if (notProvided != '') { // Check if needing special output
        return result;
      } else { // Surgeon or location type, gets different output
        if (start == 'Surgeon') {
          return '(' + result + ')'; // Add brackets
        } else {
          return ', ' + result; // Add leading comma
        }
      }
    } else {
      return notProvided; // Blank result
    }
  }

  // Link to the template docx (cannot be loaded locally due to security restrictions) - see https://docxtemplater.com/docs/tag-types/
  loadFile('https://https://goncmdm.operatingsystems.nz/newPatientTemplate.docx', function(
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
          .join('\n');
        console.log('errorMessages', errorMessages);
        showError(errorMessages);
      }
      throw error;
    }

    let zip = new PizZip(content);
    let doc;
    try {
      doc = new window.docxtemplater(zip);
    } catch (error) { // Catch compilation errors eg misplaced tags
      errorHandler(error);
    }
    
    const isMidlands = (raw.search('MIDLANDS REFERRAL INFORMATION') > -1) ? true : false;
    
    if (isMidlands) { // Remove header from Midlands referrals
      raw = raw.replace(/Referral - Super regional Gynaecology MDM \(Auckland\)/ig, '');
      // Remove footer text from Midlands referrals allowing for random spaces in the NHI and the word Generated
      raw = raw.replace(/N\s*H\s*I\s*:\s*[A-Z]\s*[A-Z]\s*[A-Z]\s*\d\s*\d\s*[A-Z0-9]\s*[A-Z0-9]\s*Date G\s*enerated: \d{1,2} \w{3} \d{4} Page number: \d+ of \d+/gi, '');
    }

    // Extract referral content
    Referral['referralDate'] = getText('Date of Referral', 'NHI Number', false, true);
    Referral['nhi'] = getText('NHI Number', 'Patient Name');
    Referral['patientName'] = getText('Patient Name', 'DOB');
    Referral['dob'] = getText('DOB', 'Address', false, true);
    Referral['address'] = getText('Address', 'Phone/');
    Referral['phone'] = getText('Mobile', /nb|high/); // /\r?\n|\r/
    Referral['referrerName'] = getText('Consultant name', 'Hospital/DHB', true);
    Referral['referrerDHB'] = getText('Hospital/DHB', 'Email address', true);
    Referral['referrerEmail'] = getText('Email address', 'GP name', true);
    Referral['gp'] = getText('GP name and address', 'HISTORY', true);
    Referral['age'] = getText('Age', 'Brief history', true);
    
    if (isMidlands) {
      Referral['history'] = getText('Brief History', 'MIDLANDS REFERRAL', true);
      Referral['menopause'] = getText('Menopausal Status', 'Gravidity');      
      Referral['gravidity'] = getText('Gravidity', 'Parity');
      Referral['parity'] = getText('Parity', 'Abortus');
      Referral['abortus'] = getText('Abortus', 'Smoking Status');
      Referral['pregnancies'] = (isNaN(parseInt(Referral['gravidity']))) ? '' : 'G' + parseInt(Referral['gravidity']);
      Referral['pregnancies'] += (isNaN(parseInt(Referral['parity']))) ? '' : 'P' + parseInt(Referral['parity']);
      Referral['pregnancies'] += (isNaN(parseInt(Referral['abortus']))) ? '' : 'A' + parseInt(Referral['abortus']);
      Referral['smoking'] = getText('Smoking Status', 'Alcohol History');
      Referral['alcohol'] = getText('Alcohol History', 'Family History');
      Referral['familyHx'] = getText('Family History', 'Frailty/G8 Score');
      // Frailty, psychosocial, and preferences are inconsistently ordered in the referral so must be combined.
      Referral['social'] = getText('Frailty/G8 Score', 'Co-morbidities');
    } else {
      Referral['history'] = getText('Brief History', /(co morbidities|co-morbidities)/, true);
    }
    
    Referral['comorbidities'] = getText('morbidities', 'Tumour', true);
    Referral['markers'] = getText('markers', 'BMI', true);
    Referral['bmi'] = getText('BMI', 'ECOG', true);
    Referral['ecog'] = getText('ECOG', 'Ethnicity', true);
    Referral['ethnicity'] = getText('Ethnicity', 'RADIOLOGY', true);

    // Array that is transformed into bullets in the docx templates
    Referral['bullets'] = [
      { 'label': 'Age', 'value': Referral['age'] },
      { 'label': 'Brief history', 'value': Referral['history'] },
      { 'label': 'Co-morbidities', 'value': Referral['comorbidities'] },
      { 'label': 'Tumour markers', 'value': Referral['markers'] },
      { 'label': 'BMI', 'value': Referral['bmi'] },
      { 'label': 'ECOG', 'value': Referral['ecog'] },
      { 'label': 'Ethnicity', 'value': Referral['ethnicity'] }
    ];
    
    if (isMidlands) { // Add populated Midlands referral parameters to the bulleted list
      if (Referral['menopause'] != 'Not provided') {
        Referral['bullets'].push({ 'label': 'Menopause status', 'value': Referral['menopause'] });
      }
      if (Referral['pregnancies'] != '') {
        Referral['bullets'].push({ 'label': 'Pregnancies', 'value': Referral['pregnancies'] });
      }
      if (Referral['smoking'] != 'Not provided') {
        Referral['bullets'].push({ 'label': 'Smoking status', 'value': Referral['smoking'] });
      }
      if (Referral['alcohol'] != 'Not provided') {
        Referral['bullets'].push({ 'label': 'Alcohol history', 'value': Referral['alcohol'] });
      }
      if (Referral['familyHx'] != 'Not provided') {
        Referral['bullets'].push({ 'label': 'Family history', 'value': Referral['familyHx'] });
      }
      Referral['social'] = Referral['social'].replace(/(Frailty\/G8 Score:|Psychosocial or High needs patient consideration:|Patient Preferences and Other Factors:)/g, ''); // Remove social labels
      if (Referral['social'].trim() != '') {
        Referral['bullets'].push({ 'label': 'Social', 'value': Referral['social'] });
      }
    } 

    let split = 'WHAT IS THE QUESTION'; // Default end
    let holder = raw;
    let i, len, loop, obj;

    if (raw.search('RADIOLOGY') > -1) {
      Referral['hasRad'] = true;
    }
    if (raw.search('OPERATION') > -1) {
      Referral['hasOp'] = true;
    }
    if (raw.search('HISTOLOGY') > -1) {
      Referral['hasHisto'] = true;
    }

    if (Referral['hasRad']) {
      // Determine the end of the loop
      if (Referral['hasOp']) {
        split = 'OPERATION';
      } else if (Referral['hasHisto']) {
        split = 'HISTOLOGY';
      }

      // Make loop hold just the radiology text
      loop = raw
        .split('RADIOLOGY')[1]
        .split(split)[0]
        .trim();
      loop = loop.substring(loop.indexOf('Type'));
      // Make loop an array of each imaging
      loop = loop.split('Type');

      len = loop.length;
      for (i = 1; i < len; i++) {
        raw = 'Type' + loop[i] + 'LOOPEND';
        obj = {
          radType: getText('Type', 'Date', true),
          radDate: getText('Date', 'Location', true, true),
          radDHB: getText('Location', 'Key findings', true),
          radFindings: getText('Findings', 'LOOPEND', true)
        };
        
        // Remove preamble before 'Findings:' and sign off
        obj['radFindings'] = obj['radFindings'].replace(/^.*?Findings:/i, '').replace(/(Reported by|This final report has been electronically|Electronically signed by).*$/, '');
        
        if (
          obj['radType'] != 'Type not provided' ||
          obj['radFindings'] != 'Not provided'
        ) {
          Referral['radiology'].push(obj);
        }
      }
      Referral['hasRad'] = !(Referral['radiology'].length == 0); // Remove radiology section if array empty
      raw = holder; // Restore raw
    }

    if (Referral['hasOp']) {
      // Determine the end of the loop
      if (Referral['hasHisto']) {
        split = 'HISTOLOGY';
      }

      // Make loop hold just the operation text
      loop = raw
        .split('OPERATION')[1]
        .split(split)[0]
        .trim();
      loop = loop.substring(loop.indexOf('Date'));
      // Make loop an array of each operation
      loop = loop.split('Date');

      len = loop.length;
      for (i = 1; i < len; i++) {
        raw = 'Date' + loop[i] + 'LOOPEND';
        obj = {
          opDate: getText('Date', 'Surgeon', true, true),
          opSurgeon: getText('Surgeon', 'Procedure', true),
          opType: getText('Procedure', 'Findings', true),
          opFindings: getText('Findings', 'LOOPEND', true)
        };
        if (
          obj['opType'] != 'Procedure not provided' ||
          obj['opFindings'] != 'Not provided'
        ) {
          Referral['operation'].push(obj);
        }
      }
      Referral['hasOp'] = !(Referral['operation'].length == 0); // Remove operation section if array empty
      raw = holder; // Restore raw
    }

    if (Referral['hasHisto']) {
      // Make loop hold just the histology text
      loop = raw
        .split('HISTOLOGY')[1]
        .split(/WHAT IS THE QUESTION/i)[0]
        .trim();

      // Make loop an array of each specimen
      if (loop.search('Specimen type') > 0) {
        loop = loop.split('Specimen type');
      } else {
        loop = loop.split('Histology type');
      }

      len = loop.length;
      for (i = 1; i < len; i++) {
        raw = 'type' + loop[i] + 'LOOPEND';
        obj = {
          histoType: getText('Type', 'Date', true),
          histoDate: getText('Date', 'Location', true, true),
          histoDHB: getText('Name of the lab', 'Key findings', true),
          histoFindings: getText('Findings', 'LOOPEND', true)
        };
        if (
          obj['histoType'] != 'Type not provided' ||
          obj['histoFindings'] != 'Not provided'
        ) {
          Referral['histology'].push(obj);
        }
      }
      Referral['hasHisto'] = !(Referral['histology'].length == 0); // Remove histology section if array empty
      raw = holder; // Restore raw
    }

    Referral['question'] = getText('FOR THE MDM?', /is the patient|what does the patient know/);

    if (DEBUG) {
      console.log(Referral);
    }

    try {       
      doc.render(Referral); // Render the .docx from template
    } catch (error) {
      errorHandler(error);
    }

    let firstName = Referral['patientName'].split(/\s(.+)/)[0].toLowerCase(); // Everything before the first space
    let lastName = Referral['patientName'].split(/\s(.+)/)[1];
    firstName = firstName[0].toUpperCase() + firstName.slice(1); // Convert firstname is sentence case
    let fileName = lastName.toUpperCase() + ' ' + firstName + '.docx';

    let out = doc.getZip().generate({
      type: 'blob',
      compression: 'DEFLATE',
      mimeType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    }); //Output the document using Data-URI

    document.getElementById('message').innerHTML = 'You just created<br />' + fileName;

    if (!DEBUG) {
      saveAs(out, fileName);
    }
  });
}

document.getElementById('submitReferral').addEventListener('click', function() {
  document.getElementById('errors').innerHTML = '';
  document.getElementById('error-container').style.display = 'none';
  generate();
  document.getElementById('referralRawText').placeholder = 'Let\'s do another one!'
  document.getElementById('referralRawText').value = '';
  document.getElementById('submitReferral').disabled = true;
  document.getElementById('submitReferral').classList.add('is-disabled')
});

function toggleButton() {
  let name = document.getElementById('referralRawText').value;
  if (name.length > 50) {
    document.getElementById('submitReferral').disabled = false;
    document.getElementById('submitReferral').classList.remove('is-disabled')
  } else {
    document.getElementById('submitReferral').disabled = true;
    document.getElementById('submitReferral').classList.add('is-disabled')
  }
}

document.getElementById('referralRawText').addEventListener('keyup', toggleButton, false);
document.getElementById('referralRawText').addEventListener('change', toggleButton, false);

// Drag and drop referral .docx to automatically parse
const referralRawText = document.getElementById('referralRawText');

// Prevent default drag behaviours
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    referralRawText.addEventListener(eventName, e => e.preventDefault(), false);
    referralRawText.addEventListener(eventName, e => e.stopPropagation(), false);
});

referralRawText.addEventListener('drop', async function(event) {
    const file = event.dataTransfer.files[0];
    
    if (!file.name.includes('.docx')) {
      showError('Incompatible file type');
    } else {
      const arrayBuffer = await file.arrayBuffer();
      const processingPromise = mammoth.extractRawText({arrayBuffer: arrayBuffer})
      .then(function(result){
        referralRawText.value = result.value;
        toggleButton();
        document.getElementById('submitReferral').click();
      })
      .catch(function(error) {
          showError(`Error processing ${file.name}:`, error);
          console.error(error);
      });
    }
});
