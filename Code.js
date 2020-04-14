function onOpen()
{
  DocumentApp.getUi()
      .createMenu('Macro')
      .addItem('Add Title Page Information', 'opentitlePageDialog')
      .addItem('Add Tests', 'openTestDialog')
      .addToUi();
}

function opentitlePageDialog()
{
  var html = HtmlService.createHtmlOutputFromFile('TitlePage'); // Create an HTML output using 'TitlePage.html'.
  html.setTitle("Add Title Page Information"); // Set the title of the sidebar.
  DocumentApp.getUi() 
      .showSidebar(html); // Display the HTML file as a sidebar.
}

function openTestDialog()
{
  var html = HtmlService.createHtmlOutputFromFile('Tests'); 
  html.setTitle("Add Tests"); 
  DocumentApp.getUi() 
      .showSidebar(html); 
}

function populateGeneralInformation(form)
{
  var marker = ['{Test Plan Name}', '{Name}', '{Date}', '{Jira #}', '{Part name}', '{Part number}', '{Project name}', '{testSummary}'];
  var fieldName = [form.testPlanName, form.fullName, form.date, form.jiraTicketNumber, form.partName, form.partNumber, form.projectName, form.testSummary]; // I learned that arrays can store a class with a property, like form.fullName. Arrays can store any data type!
  var body = DocumentApp.getActiveDocument().getBody();
  var footer = DocumentApp.getActiveDocument().getFooter().getParent().getChild(4); // Because the first page has a different footer, DocumentApp.getActiveDocument().getFooter(); would check only the first page.
  for (i=0; i<marker.length; i++)
  {
    if (fieldName[i] != "")
    {
      body.replaceText(marker[i], fieldName[i]); // First argument is the location in the document, second argument is the matching text field on the sidebar.
      footer.replaceText(marker[i], fieldName[i]); 
    }
  }
  nameFile(form.testPlanName, form.partNumber, form.partName, form.date);
  storeData(form.date);
}

function populateTestSpecificInformation(form)
{
  var marker = [];
  var fieldName = [];
  var body = DocumentApp.getActiveDocument().getBody();

  switch (form.hiddenTestName)
  { 
    case "1": /* "coldSoakVariables" */
      marker = ['{tooling}', '{notes}', '{dwellTemp}', '{dwellTime}', '{minVoltage}', '{maxVoltage}'];
      fieldName = [form.tooling, form.notes, form.dwellTemp, form.dwellTime, form.minVoltage, form.maxVoltage];
      break;
    case "2": /*"hotSoakVariables" */
      marker = ['{tooling}', '{notes}', '{dwellTemp}', '{dwellTime}', '{minVoltage}', '{maxVoltage}', '{steadyTime}', '{voltageRange}'];
      fieldName = [form.tooling, form.notes, form.dwellTemp, form.dwellTime, form.minVoltage, form.maxVoltage, form.steadyTime, form.voltageRange];
      break;
    case "3": /*"hotStorageVariables"*/
      marker = ['{tooling}', '{notes}', '{dwellTemp}', '{dwellTime}'];
      fieldName = [form.tooling, form.notes, form.dwellTemp, form.dwellTime];
      break;
    case "4": /*"coldStorageVariables"*/
      marker = ['{tooling}', '{notes}', '{dwellTemp}', '{dwellTime}'];
      fieldName = [form.tooling, form.notes, form.dwellTemp, form.dwellTime];
      break;
    case "5": /*"thermalCycleVariables"*/
      marker = ['{tooling}', '{notes}', '{minTemp}', '{maxTemp}', '{dwellTime}', '{cycles}'];
      fieldName = [form.tooling, form.notes, form.minTemp, form.maxTemp, form.dwellTime, form.cycles]; 
      break;
    case "6": /*"thermalShockVariables"*/
      marker = ['{tooling}', '{notes}', '{minTemp}', '{maxTemp}', '{dwellTime}', '{shockEvents}'];
      fieldName = [form.tooling, form.notes, form.minTemp, form.maxTemp, form.dwellTime, form.shockEvents];  
      break;
    case "7": /*"enduranceVibrationVariables"*/ 
      marker = ['{notes}'];
      fieldName = [form.notes];
      break;
    case "8": /* "chemicalExposureVariables" */
      marker = ['{notes}'];
      fieldName = [form.notes];
      break;
    case "9": /* "functionalVibrationVariables" */
      marker = ['{notes}'];
      fieldName = [form.notes];
      break;
    case "10": /* "istaDropVariables" */
      marker = ['{tooling}', '{notes}'];
      fieldName = [form.tooling, form.notes];
      break;
    case "11": /* "transitDropVariables" */
      marker = ['{tooling}', '{notes}'];
      fieldName = [form.tooling, form.notes];
      break;
    case "12": /* humidityToleranceVariables*/
    marker = ['{notes}', '{cycles}', '{cycleTime}', '{dwellTime}'];
    fieldName = [form.notes, form.cycles, form.cycleTime, form.dwellTime];
    break;
    case "13": /* corrosionVariables*/
    marker = ['{notes}', '{days}', '{cycles}'];
    fieldName = [form.notes, form.days, form.cycles];
    break;
  }
    for (i=0; i<marker.length; ++i)
    {
      if (fieldName[i] != "")
      {
        body.replaceText(marker[i], fieldName[i]); // First argument is the location in the document. Second argument is the matching text field on the sidebar.
      }
    }
}

function addNewTest(form)
{
  var selectedTest = form.tests;
  switch (selectedTest)
  {
    case "thermalCycle":
      appendTest('https://docs.google.com/document/d/1pdqgD7eWGZfip1K795DE0Pal0qG5NsNTxAe-CfJyXkQ/edit');
      break;
    case "chemicalExposure":
      appendTest('https://docs.google.com/document/d/1BMM0rXBD-rB-9tWraeDT6lnCEUSx3timBVt86inHNbs/edit');
      break;
    case "coldSoak":
      appendTest('https://docs.google.com/document/d/1dsrzhk2AqhhJIS47mieGwTNUmeFL2l9tB4xWlJkEJ5c/edit');
      break;
    case "hotSoak":
      appendTest('https://docs.google.com/document/d/1_MVYJUq45_QSDnc6KTuh1Fkula_Lp8--K_u30FLsC5s/edit');
      break;
    case "coldStorage":
      appendTest('https://docs.google.com/document/d/1OklUjVlpdcEsk4V0dS-NHteJxE2J1wpF7LWyfisrHrY/edit');
      break;
    case "hotStorage":
      appendTest('https://docs.google.com/document/d/1oC4i4jLXp2xG-5yAV8D80vmewMoralx14RWngX4NM3w/edit');
      break;
    case "enduranceVibration":
      appendTest('https://docs.google.com/document/d/1X4U0m360d6m9b4cdz9GGlRUsFOo53nmyHgL6bhqECSA/edit');
      break;
    case "functionalVibration":
      appendTest('https://docs.google.com/document/d/1gPzlT8K64NxwVMyoyGHMndo2NtuQcqqnLGrISwMIIbI/edit#');
      break;
    case "istaDrop":
      appendTest('https://docs.google.com/document/d/1qL0A0cFN5nbHcxq0lR9vMlKd1ckAvrletOPKzMPJ0HI/edit');
      break;
    case "thermalShock":
      appendTest('https://docs.google.com/document/d/1mHgwVLi0PeGopI6vLiTdW1vamGXcJ6Av-RTre7wxxec/edit');
      break;
    case "transitDrop":
      appendTest('https://docs.google.com/document/d/1LvmJm6EVaJkp0D2ZrqGiyUD8oDWLpco9nU1ohgh6ems/edit');
      break;
    case "userDefined":
      appendTest('https://docs.google.com/document/d/1xRd1Cn_6zIMeH-tmcAxBZgPSzO605cfQaviyoGXuUUk/edit');
      break;
    case "humidityTolerance":
      appendTest('https://docs.google.com/document/d/1qYYNoMIdfjbaSpwEVouJIODRPGIe44xwdiOX4D7fQF0/edit#')
      break;
    case "corrosion":
      appendTest('https://docs.google.com/document/d/15o-8TEO7IkKZgzghOEFed5YeUzcE2VdsEcPXQ09wewo/edit#')
      break;          
  }
}

function appendTest(testID)
{
  var thisDoc = DocumentApp.getActiveDocument();
  var thisBody = thisDoc.getBody();
  var templateDoc = DocumentApp.openByUrl(testID); // Pass in id of doc to be used as a template.
  var templateBody = templateDoc.getBody();
  var sizeOfItems = {};
  sizeOfItems[DocumentApp.Attribute.FONT_SIZE] = 10;

  for (var i = 0; i < templateBody.getNumChildren(); i++)
  { // Run through the elements of the template doc's Body.
    switch (templateBody.getChild(i).getType())
    { // Deal with the various types of Elements we will encounter and append.
      case DocumentApp.ElementType.PARAGRAPH:
        thisBody.appendParagraph(templateBody.getChild(i).copy());
        break;
      case DocumentApp.ElementType.LIST_ITEM:
        var typeOfList = templateBody.getChild(i).getGlyphType(); // Determine the type of the list item. Returns an object.
        var typeOfListAsString = JSON.stringify(typeOfList); // Convert the type from an object to a string that will be used in the if statement.
        if (typeOfListAsString = "NUMBER")
        {
          var item = thisBody.appendListItem(templateBody.getChild(i).copy());
          item.setAttributes(sizeOfItems);
          item.setGlyphType(DocumentApp.GlyphType.NUMBER);
        } else if (typeOfListAsString = "BULLET")
        {
          thisBody.appendListItem(templateBody.getChild(i).copy()).setGlyphType(DocumentApp.GlyphType.BULLET);
        } else {
          thisBody.appendListItem(templateBody.getChild(i).copy()).setGlyphType(DocumentApp.GlyphType.LATIN_LOWER);
        }
        break;
      case DocumentApp.ElementType.TABLE:
        thisBody.appendTable(templateBody.getChild(i).copy());
        break;
      case DocumentApp.ElementType.INLINE_IMAGE:
        thisBody.appendImage(templateBody.getChild(i).copy());
        break;
    }
  }
  var doc = DocumentApp.getActiveDocument();
  var paragraphs = doc.getBody().getParagraphs();
  var position = doc.newPosition(paragraphs[paragraphs.length - 1], 0);
  doc.setCursor(position);
  thisBody.appendPageBreak(); // Add a page break once the entire test has been copied.
}

function nameFile(type, number, name, date)
{
  DocumentApp.getActiveDocument().setName(type + " - " + number + " - " + name + " - " + date);
}

function storeData() 
{
  var database = FirebaseApp.getDatabaseByUrl("https://ctct-environmental-test-plans.firebaseio.com/", "xTw3pwH5Gt8lZk4t9FgA2hpTtblfz0J7azfnM2sD");
  var email = Session.getActiveUser().getEmail(); // Get the user's email.
  var date = new Date();
  var docName = DocumentApp.getActiveDocument().getName();
  var storageLocation = '';

  date = (date.getMonth() + 1) + '/' + date.getDate() + '/' + date.getFullYear();
  var userInfo = {
    email: email,
    date: date,
    docName: docName
  }
  email = email.split("@", 1);
  storageLocation = email + Math.round(Math.random() * 10000).toString();
  database.setData(storageLocation, userInfo);
}