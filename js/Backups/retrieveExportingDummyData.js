// Import necessary modules
const smartsheet = require('smartsheet');

// Smartsheet API token and workspace ID
const smartsheetToken = "vwM6AHwhUlFPs5pLk5WIqzt82FSqWkoOeIqyp";
const sheetID = "6561699532328836";

// Initialize Smartsheet client
const smartsheetClient = smartsheet.createClient({
    accessToken: smartsheetToken
});

// Column IDs
const columnIds = {
    EmployeeName: 1223039130750852,
    LeadStatus: 5726638758121348,
    TimeIn: 3474838944436100,
    TimeOut: 7978438571806596,
    Incentive: 660089177329540,
    Notes: 5163688804700036
};

// Dummy data
const dummyData = {
    EmployeeName: 'Ezequiel HAHAHAHAHA',
    LeadStatus: 'Yes',
    TimeIn: '09:00',
    TimeOut: '17:30',
    Incentive: 'Bike',
    Notes: 'Hello world'
};



// Function to add row to Smartsheet
function addRowToSheet(data, columnIds) {
    const rowToAdd = {
        toBottom: true,
        cells: []
    };

    // Construct cells based on column IDs
    Object.keys(data).forEach(key => {
        if (columnIds[key]) {
            rowToAdd.cells.push({
                columnId: columnIds[key],
                value: data[key]
            });
        } else {
            console.warn(`Column ID not found for '${key}'. Skipping...`);
        }
    });

    // Add row to sheet
    smartsheetClient.sheets.addRow({
            sheetId: sheetID,
            body: rowToAdd
        })
        .then(response => {
            console.log("Row added successfully:", response.result);
        })
        .catch(error => {
            console.error("Error adding row:", error);
        });
}

// Add the dummy data row to Smartsheet
addRowToSheet(dummyData, columnIds);
