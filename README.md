# Template Generator
## Background
This is a Google Apps script I wrote that generates a test plan from user input for Trimble's lab. The user picks what tests to add, then they specify the test variables. When the user submits title page information, the document automatically names itself and stores some general document details in Firebase Database to determine how often the template is used. 

## Provide Title Page Information
1. In the toolbar, click Macro, and then click Add Title Page Information. If this is your first time running a script in this document, you'll be prompted to authorize the script. A sidebar will appear.
2. Fill out every text field in the sidebar, and then click Submit. The title page fields and the file name will update using your input.
![Adding title page information](https://github.com/dorrzun/TemplateMaker/blob/master/Animations/Adding%20title%20page%20information.gif)

## Add a Test and Specify Test Variables
1. In the toolbar, click Macro, and then click Add Tests. A sidebar will appear.
2. Select a test from the drop-down, and then click Add Test. The sidebar will prompt you for additional information, and the test will be appended to your document.
3. Fill out every text field in the sidebar (write 'None' if there are no tooling or notes), and then click Submit Information for This Test. The test fields will update using your input.
![Adding a test and test information](https://github.com/dorrzun/TemplateMaker/blob/master/Animations/Adding%20a%20test%20and%20test%20information.gif)
