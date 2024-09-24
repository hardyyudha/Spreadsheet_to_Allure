function convertSheetToAllureJSON() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeSheetName = sheet.getSheetName();
    const range = sheet.getDataRange();
    const values = range.getValues();
    const mergedRanges = range.getMergedRanges();

    // Define an array to store the test cases
    let testCases = {};

    // Helper function to get merged cell values
    function getMergedValue(row, col) {
        for (let mergedRange of mergedRanges) {
            const startRow = mergedRange.getRow() - 1;
            const startCol = mergedRange.getColumn() - 1;
            const numRows = mergedRange.getNumRows();
            const numCols = mergedRange.getNumColumns();

            if (row >= startRow && row < startRow + numRows && col >= startCol && col < startCol + numCols) {
                return mergedRange.getValue();
            }
        }
        return values[row][col];
    }

    // Iterate through each row of the sheet
    for (let i = 1; i < values.length; i++) {
        const testScenario = getMergedValue(i, 0);
        const idTest = getMergedValue(i, 1);
        const testCase = getMergedValue(i, 2);
        const testStep = getMergedValue(i, 3);
        const expectedResult = getMergedValue(i, 4);
        const actualResult = getMergedValue(i, 5);
        const testStatus = getMergedValue(i, 6);
        const testEvidence = getMergedValue(i, 7);

        // Skip empty rows
        if (!testScenario || !idTest || !testCase) continue;

        // Initialize the test case entry if it doesn't exist
        if (!testCases[idTest]) {
            testCases[idTest] = {
                "name": testCase,
                "status": testStatus || "unknown",
                "steps": [],
                "start": new Date().getTime(), // Add start time (could be adjusted for your needs)
                "stop": new Date().getTime() + 10000, // Add stop time (could be adjusted),
                "attachments": testEvidence ? [{
                    name: `${testEvidence}`,
                    type: '', // You can specify the type if needed
                    source: ''
                }] : []
            };
        }

        // Add the test step to the test case steps array
        if (testStep) {
            testCases[idTest].steps.push({
                "name": testStep,
                "status": expectedResult === actualResult ? "passed" : "failed"
            });
        }
    }

    // Generate the individual JSON files for each test case
    for (const idTest in testCases) {
        const testCaseObject = testCases[idTest];

        // Convert the test case object to JSON
        const json = JSON.stringify(testCaseObject, null, 2);

        // Save the JSON file to Google Drive
        const fileName = `${activeSheetName}-${idTest}-result.json`;
        const folderId = 'your_drive_folder_id'; // Replace with your Google Drive folder ID
        const folder = DriveApp.getFolderById(folderId);
        folder.createFile(fileName, json, MimeType.PLAIN_TEXT);

        // Log the generated JSON (for debugging)
        Logger.log(json);
    }
}
