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
  var html = HtmlService.createHtmlOutputFromFile('TitlePage'); // Create an HTML output using 'TitlePage.html'
  html.setTitle("Add Title Page Information"); // Set the title of the sidebar
  DocumentApp.getUi() 
      .showSidebar(html); // Display the HTML file as a sidebar
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
  var marker = ['{Test Plan Name}', '{Name}', '{Date}', '{Jira #}', '{Part name}', '{Part number}', '{Project name}'];
  var fieldName = [form.testPlanName, form.fullName, form.date, form.jiraTicketNumber, form.partName, form.partNumber, form.projectName]; // I learned that arrays can store a class with a property, like form.fullName. Arrays can store any data type!
  
  var body = DocumentApp.getActiveDocument().getBody();
  var footer = DocumentApp.getActiveDocument().getFooter();
  for (i=0; i<marker.length; i++)
  {
    if (fieldName[i] != "")
    {
      body.replaceText(marker[i], fieldName[i]); // First argument is the location in the document, second argument is the matching text field on the sidebar.
      //footer.replaceText(marker[i], fieldName[i]); // TO DO: The footer text isn't being replaced for some reason.
      DocumentApp.getUi().alert(marker[i],footer[i]);
    }
  }
}

function addNewTest(form) {
  var selectedTest = form.tests;
  switch (selectedTest) {
    case "thermalCycle":
      appendTest('https://docs.google.com/document/d/1WZLUd_iRxE5uwvMsFIBV-DMvpGlWa8-tbdAg0QIYTo0/edit');
      break;
    case "chemicalExposure":
      appendTest('https://docs.google.com/document/d/1Tkp-byLQ9MUcBLqyPPTP568x8rad4QcgQ9cEkV1CFk8/edit');
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
      appendTest('https://docs.google.com/document/d/1c9-j9FwiH_H53BMk_n4kJtOwY2Kz6cFEwomtkUup67E/edit');
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
  }
}

function appendTest(testID) {
  var thisDoc = DocumentApp.getActiveDocument();
  var thisBody = thisDoc.getBody();
  var templateDoc = DocumentApp.openByUrl(testID); // Pass in id of doc to be used as a template.

  var templateBody = templateDoc.getBody();

  for (var i = 0; i < templateBody.getNumChildren(); i++) { // run through the elements of the template doc's Body.
    switch (templateBody.getChild(i).getType()) { // Deal with the various types of Elements we will encounter and append.
      case DocumentApp.ElementType.PARAGRAPH:
        thisBody.appendParagraph(templateBody.getChild(i).copy());
        break;
      case DocumentApp.ElementType.LIST_ITEM:
        thisBody.appendListItem(templateBody.getChild(i).copy());
        break;
      case DocumentApp.ElementType.TABLE:
        thisBody.appendTable(templateBody.getChild(i).copy());
        break;
    }
  }
  return thisDoc;
}
