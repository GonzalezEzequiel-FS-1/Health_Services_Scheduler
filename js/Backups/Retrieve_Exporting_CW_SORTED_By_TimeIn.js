const axios = require('axios');
const mysql = require('mysql');

// Database connection configuration
const dbConfig = {
    host: '23.117.213.174',
    port: 3369,
    user: 'Zeke',
    password: 'TrashPanda6000',
    database: 'Health_Services'
};

// Smartsheet API token and sheet ID
const smartsheetToken = "vRF9IrGXepLhGkB5CwehrGRbt7oOy5oEmP4Rd";
// Official Token: "vwM6AHwhUlFPs5pLk5WIqzt82FSqWkoOeIqyp"
// Backup Token: vRF9IrGXepLhGkB5CwehrGRbt7oOy5oEmP4Rd
const sheetID = "6561699532328836";

// Column IDs
const columnIds = {
    Employee: 1223039130750852,
    LeadStatus: 5726638758121348,
    TimeIn: 3474838944436100,
    TimeOut: 7978438571806596,
    Incentive: 660089177329540,
    Notes: 5163688804700036
};

// Create a connection to the database
const connection = mysql.createConnection(dbConfig);

// Connect to the database
connection.connect((err) => {
    if (err) {
        console.error('Error connecting to the database:', err);
        return;
    }
    console.log('Connected to the database successfully!');
});

// Function to clear the contents of all rows in the sheet
async function clearAllRows() {
    try {
        console.log("Clearing all rows from the Smartsheet...");
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

// Function to retrieve sorted data from the database
function retrieveSortedDataFromDatabase() {
    return new Promise((resolve, reject) => {
        console.log("Retrieving sorted data from the database...");
        const query = "SELECT EmployeeName, LeadStatus, TimeIn, TimeOut, Incentive, Notes FROM CityWalkSchedule ORDER BY TimeIn DESC";
        connection.query(query, (err, results) => {
            if (err) {
                reject(err);
                return;
            }
            console.log("Retrieved sorted data from the database successfully:", results);
            resolve(results);
        });
    });
}


// Function to add row to Smartsheet
async function addRowToSheet(columnIds) {
    try {
        // Retrieve sorted data from the database
        const dataFromDB = await retrieveSortedDataFromDatabase();
        
        // Iterate over each row of data retrieved from the database
        for (const row of dataFromDB) {
            // Assign values from the database to variables
            const employeeData = row.EmployeeName;
            const leadStatusData = row.LeadStatus;
            const timeInData = row.TimeIn;
            const timeOutData = row.TimeOut;
            const incentiveData = row.Incentive;
            const notesData = row.Notes;

            // Construct row object for Smartsheet
            const rowToAdd = {
                toTop: true,
                cells: [
                    { columnId: columnIds.Employee, value: employeeData },
                    { columnId: columnIds.LeadStatus, value: leadStatusData },
                    { columnId: columnIds.TimeIn, value: timeInData },
                    { columnId: columnIds.TimeOut, value: timeOutData },
                    { columnId: columnIds.Incentive, value: incentiveData },
                    { columnId: columnIds.Notes, value: notesData }
                ]
            };

            // Add row to Smartsheet
            console.log("Adding row to Smartsheet:", rowToAdd);
            await axios.post(`https://api.smartsheet.com/2.0/sheets/${sheetID}/rows`, rowToAdd, {
                headers: {
                    'Authorization': `Bearer ${smartsheetToken}`
                }
            });
            console.log("Row added successfully.");
        }

        console.log("All rows added successfully.");
    } catch (error) {
        if (error.response) {
            console.error("Error adding rows:", error.response.data);
        } else {
            console.error("Error adding rows:", error.message);
        }
    }
}



// Clear all rows from the sheet first
clearAllRows()
    .then(() => {
        // Then add the data from the database to Smartsheet after clearing rows
        addRowToSheet(columnIds);
    })
    .catch(err => {
        console.error("Error:", err);
    })
    .finally(() => {
        // Close the connection when done
        connection.end();
    });
