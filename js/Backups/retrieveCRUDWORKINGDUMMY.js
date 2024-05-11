const axios = require('axios');

// Smartsheet API token and sheet ID
const smartsheetToken = "vRF9IrGXepLhGkB5CwehrGRbt7oOy5oEmP4Rd";
// Official Token: "vwM6AHwhUlFPs5pLk5WIqzt82FSqWkoOeIqyp"
// Backup Token: vRF9IrGXepLhGkB5CwehrGRbt7oOy5oEmP4Rd
const sheetID = "6561699532328836";

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
    EmployeeName: 'Ezequiel Hello',
    LeadStatus: 'Yes',
    TimeIn: '09:00',
    TimeOut: '17:30',
    Incentive: 'Bike',
    Notes: 'Hello Mundo Nuevo'
};

// Function to clear the contents of all rows in the sheet
async function clearAllRows() {
    try {
        const response = await axios.get(`https://api.smartsheet.com/2.0/sheets/${sheetID}`, {
            headers: {
                'Authorization': `Bearer ${smartsheetToken}`
            }
        });
        const rows = response.data.rows;
        const requests = rows.map(row => ({
            method: 'PUT',
            url: `https://api.smartsheet.com/2.0/sheets/${sheetID}/rows/${row.id}`,
            headers: {
                'Authorization': `Bearer ${smartsheetToken}`,
                'Content-Type': 'application/json'
            },
            data: {
                cells: row.cells.map(cell => ({
                    columnId: cell.columnId,
                    value: null
                }))
            }
        }));
        await axios.all(requests.map(req => axios(req)));
        console.log("Rows cleared successfully.");
    } catch (error) {
        console.error("Error clearing rows:", error.response.data);
    }
}

// Function to add row to Smartsheet
async function addRowToSheet(data, columnIds) {
    const rowToAdd = {
        toTop: true, // Add the new row to the top of the sheet
        cells: Object.keys(data).map(key => ({
            columnId: columnIds[key],
            value: data[key]
        }))
    };

    try {
        await axios.post(`https://api.smartsheet.com/2.0/sheets/${sheetID}/rows`, rowToAdd, {
            headers: {
                'Authorization': `Bearer ${smartsheetToken}`
            }
        });
        console.log("Row added successfully.");
    } catch (error) {
        console.error("Error adding row:", error.response.data);
    }
}

// Clear all rows from the sheet first
clearAllRows()
    .then(() => {
        // Then add the dummy data row to Smartsheet after clearing rows
        addRowToSheet(dummyData, columnIds);
    });
