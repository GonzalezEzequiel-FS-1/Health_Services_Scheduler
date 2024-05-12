const idFinder=()=>{
//Use This script to get the Column Names and IDs given the Sheet Id results logged to the console in JSON
const token = 'vwM6AHwhUlFPs5pLk5WIqzt82FSqWkoOeIqyp';

// Function to make a GET request to Smartsheet API
async function fetchSheet(sheetId) {
    const url = `https://api.smartsheet.com/2.0/sheets/${sheetId}`;
    const response = await fetch(url, {
        method: 'GET',
        headers: {
            'Authorization': `Bearer ${token}`,
            'Content-Type': 'application/json'
        }
    });
    return response.json();
}

// Array of sheet IDs to fetch data from
const sheetIds = [
  1724277866844036
];

// Function to fetch row IDs and column titles for each sheet
async function fetchSheetData() {
    const sheetData = [];
    for (const sheetId of sheetIds) {
        const response = await fetchSheet(sheetId);
        const rowData = {
            sheetName: response.name,
            sheetId: sheetId,
            columns: response.columns.map(column => ({
                columnName: column.title,
                columnId: column.id
            }))
        };
        sheetData.push(rowData);
    }
    return sheetData;
}

// Fetch data for each sheet and log the results in JSON format
fetchSheetData()
    .then(data => {
        console.log(JSON.stringify(data, null, 2));
    })
    .catch(error => console.error('Error fetching sheet data:', error));

  }
  idFinder();
