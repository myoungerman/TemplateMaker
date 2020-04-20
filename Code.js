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
  html.setTitle('Add Title Page Information'); // Set the title of the sidebar.
  DocumentApp.getUi() 
      .showSidebar(html); // Display the HTML file as a sidebar.
}

function openTestDialog()
{
  var html = HtmlService.createHtmlOutputFromFile('Tests'); 
  html.setTitle('Add Tests'); 
  DocumentApp.getUi() 
      .showSidebar(html); 
}

function populateGeneralInformation(form)
{
  const marker = ['{Test Plan Name}', '{Name}', '{Date}', '{Jira #}', '{Part name}', '{Part number}', '{Project name}', '{testSummary}'];
  const fieldName = [form.testPlanName, form.fullName, form.date, form.jiraTicketNumber, form.partName, form.partNumber, form.projectName, form.testSummary];
  const body = DocumentApp.getActiveDocument().getBody();
  const footer = DocumentApp.getActiveDocument().getFooter().getParent().getChild(4); // Because the first page has a different footer, DocumentApp.getActiveDocument().getFooter(); would check only the first page.
  for (i=0; i<marker.length; i++)
  {
    if (fieldName[i] != '')
    {
      body.replaceText(marker[i], fieldName[i]); // First argument is the location in the document, second argument is the matching text field on the sidebar.
      footer.replaceText(marker[i], fieldName[i]); 
    }
  }
  nameFile(form.testPlanName, form.partNumber, form.partName, form.date);
  storeData(form.date);
}

function populateTestSpecificInformation(test)
{
  let marker = [];
  const body = DocumentApp.getActiveDocument().getBody();
  let limit = 0;

  switch (test.name)
  { 
    case 'coldSoakVariables':
      marker = ['{tooling}', '{notes}', '{dwellTemp}', '{dwellTime}', '{minVoltage}', '{maxVoltage}'];
      break;
    case 'hotSoakVariables':
      marker = ['{tooling}', '{notes}', '{dwellTemp}', '{dwellTime}', '{minVoltage}', '{maxVoltage}', '{steadyTime}', '{voltageRange}'];
      break;
    case 'hotStorageVariables':
      marker = ['{tooling}', '{notes}', '{dwellTemp}', '{dwellTime}'];
      break;
    case 'coldStorageVariables':
      marker = ['{tooling}', '{notes}', '{dwellTemp}', '{dwellTime}'];
      break;
    case 'thermalCycleVariables':
      marker = ['{tooling}', '{notes}', '{minTemp}', '{maxTemp}', '{dwellTime}', '{cycles}']; 
      break;
    case 'thermalShockVariables':
      marker = ['{tooling}', '{notes}', '{minTemp}', '{maxTemp}', '{dwellTime}', '{shockEvents}'];
      break;
    case 'enduranceVibrationVariables':
      marker = ['{notes}'];
      break;
    case 'chemicalExposureVariables':
      marker = ['{notes}'];
      break;
    case 'functionalVibrationVariables':
      marker = ['{notes}'];
      break;
    case 'istaDropVariables':
      marker = ['{tooling}', '{notes}'];
      break;
    case 'transitDropVariables':
      marker = ['{tooling}', '{notes}'];
      break;
    case 'humidityToleranceVariables':
      marker = ['{notes}', '{cycles}', '{cycleTime}', '{dwellTime}'];
    break;
    case 'corrosionVariables':
      marker = ['{notes}', '{days}', '{cycles}'];
    break;
  }
    limit = marker.length;
    for (let i = 0; i < limit; i++)
    {
      if (test.variables[i] != '')
      {
        body.replaceText(marker[i], test.variables[i]); // First argument is the location in the document. Second argument is the matching text field on the sidebar.
      }
    }
}

function addNewTest(testName)
{
  switch (testName)
  {
    case 'thermalCycle':
      appendTest('https://docs.google.com/document/d/1pdqgD7eWGZfip1K795DE0Pal0qG5NsNTxAe-CfJyXkQ/edit');
      break;
    case 'chemicalExposure':
      appendTest('https://docs.google.com/document/d/1BMM0rXBD-rB-9tWraeDT6lnCEUSx3timBVt86inHNbs/edit');
      break;
    case 'coldSoak':
      appendTest('https://docs.google.com/document/d/1dsrzhk2AqhhJIS47mieGwTNUmeFL2l9tB4xWlJkEJ5c/edit');
      break;
    case 'hotSoak':
      appendTest('https://docs.google.com/document/d/1_MVYJUq45_QSDnc6KTuh1Fkula_Lp8--K_u30FLsC5s/edit');
      break;
    case 'coldStorage':
      appendTest('https://docs.google.com/document/d/1OklUjVlpdcEsk4V0dS-NHteJxE2J1wpF7LWyfisrHrY/edit');
      break;
    case 'hotStorage':
      appendTest('https://docs.google.com/document/d/1oC4i4jLXp2xG-5yAV8D80vmewMoralx14RWngX4NM3w/edit');
      break;
    case 'enduranceVibration':
      appendTest('https://docs.google.com/document/d/1X4U0m360d6m9b4cdz9GGlRUsFOo53nmyHgL6bhqECSA/edit');
      break;
    case 'functionalVibration':
      appendTest('https://docs.google.com/document/d/1gPzlT8K64NxwVMyoyGHMndo2NtuQcqqnLGrISwMIIbI/edit#');
      break;
    case 'istaDrop':
      appendTest('https://docs.google.com/document/d/1qL0A0cFN5nbHcxq0lR9vMlKd1ckAvrletOPKzMPJ0HI/edit');
      break;
    case 'thermalShock':
      appendTest('https://docs.google.com/document/d/1mHgwVLi0PeGopI6vLiTdW1vamGXcJ6Av-RTre7wxxec/edit');
      break;
    case 'transitDrop':
      appendTest('https://docs.google.com/document/d/1LvmJm6EVaJkp0D2ZrqGiyUD8oDWLpco9nU1ohgh6ems/edit');
      break;
    case 'userDefined':
      appendTest('https://docs.google.com/document/d/1xRd1Cn_6zIMeH-tmcAxBZgPSzO605cfQaviyoGXuUUk/edit');
      break;
    case 'humidityTolerance':
      appendTest('https://docs.google.com/document/d/1qYYNoMIdfjbaSpwEVouJIODRPGIe44xwdiOX4D7fQF0/edit#')
      break;
    case 'corrosion':
      appendTest('https://docs.google.com/document/d/15o-8TEO7IkKZgzghOEFed5YeUzcE2VdsEcPXQ09wewo/edit#')
      break;          
  }
}

function appendTest(testID)
{
  const thisDoc = DocumentApp.getActiveDocument();
  const thisBody = thisDoc.getBody();
  let templateDoc = DocumentApp.openByUrl(testID); // Pass in id of doc to be used as a template.
  const templateBody = templateDoc.getBody();
  const elType = DocumentApp.ElementType;
  const itemType = DocumentApp.GlyphType;
  let sizeOfItems = {};
  sizeOfItems[DocumentApp.Attribute.FONT_SIZE] = 10;

  for (let i = 0; i < templateBody.getNumChildren(); i++) { // Run through the elements of the template doc's Body.
    switch (templateBody.getChild(i).getType()) { // Deal with the various types of Elements we will encounter and append.
      case elType.PARAGRAPH:
        thisBody.appendParagraph(templateBody.getChild(i).copy());
        break;
      case elType.LIST_ITEM:
        let typeOfList = templateBody.getChild(i).getGlyphType(); // Determine the type of the list item. Returns an object.
        let typeOfListAsString = typeOfList.toString(); // Convert the type from an object to a string that will be used in the if statement.
        if (typeOfListAsString === 'NUMBER')
        {
          let item = thisBody.appendListItem(templateBody.getChild(i).copy());
          item.setAttributes(sizeOfItems);
          item.setGlyphType(itemType.NUMBER);
        } else if (typeOfListAsString === 'BULLET')
        {
          thisBody.appendListItem(templateBody.getChild(i).copy()).setGlyphType(itemType.BULLET);
        } else {
          thisBody.appendListItem(templateBody.getChild(i).copy()).setGlyphType(itemType.LATIN_LOWER);
        }
        break;
      case elType.TABLE:
        thisBody.appendTable(templateBody.getChild(i).copy());
        break;
      case elType.INLINE_IMAGE:
        thisBody.appendImage(templateBody.getChild(i).copy());
        break;
    }
  }
  let paragraphs = thisDoc.getBody().getParagraphs();
  let position = thisDoc.newPosition(paragraphs[paragraphs.length - 1], 0);
  thisDoc.setCursor(position);
  thisBody.appendPageBreak();
}

function nameFile(type, number, name, date)
{
  DocumentApp.getActiveDocument().setName(type + ' - ' + number + ' - ' + name + ' - ' + date);
}

function storeData() 
{
  const database = FirebaseApp.getDatabaseByUrl('https://ctct-environmental-test-plans.firebaseio.com/', 'xTw3pwH5Gt8lZk4t9FgA2hpTtblfz0J7azfnM2sD');
  let email = Session.getActiveUser().getEmail(); // Get the user's email.
  const date = new Date();
  let docName = DocumentApp.getActiveDocument().getName();
  let storageLocation = '';

  let day = (date.getMonth() + 1) + '/' + date.getDate() + '/' + date.getFullYear();
  const userInfo = {
    email: email,
    date: day,
    docName: docName
  }
  email = email.split('@', 1);
  storageLocation = email + day + (date.getTime()).toString();
  database.setData(storageLocation, userInfo);
}