const XLSX = require('xlsx');
const moment = require('moment');
const mysql = require('mysql');
const fs = require('fs');
const path = require('path');

// Database connection configuration
const dbConfig = {
    host: '23.117.213.174',
    port: 3369,
    user: 'Zeke',
    password: 'TrashPanda6000',
    database: 'Health_Services'
};
// Create a MySQL connection pool
const pool = mysql.createPool(dbConfig);

// Function to clear data from a table
function clearTable(tableName, callback) {
    const sql = `DELETE FROM ${tableName}`;

    pool.query(sql, (error, results, fields) => {
        if (error) {
            console.error(`Error clearing data from ${tableName}: ${error.message}`);
            if (callback) {
                callback(error);
            }
        } else {
            console.log(`Cleared data from ${tableName}`);
            if (callback) {
                callback(null);
            }
        }
    });
}


// Function to insert employees into the database
function insertEmployees(employees, tableName, callback) {
    const columns = ['EmployeeName', 'LeadStatus', 'TimeIn', 'TimeOut', 'Incentive', 'Notes'];
    const values = employees.map(employee => [employee.Name, employee.StatusLead ? 'Yes' : 'No', employee.TimeIn, employee.TimeOut, '', employee.Notes]);
    const sql = `INSERT INTO ${tableName} (${columns.join(', ')}) VALUES ?`;

    pool.query(sql, [values], (error, results, fields) => {
        if (error) {
            console.error(`Error inserting employees into ${tableName}: ${error.message}`);
            callback(error);
        } else {
            console.log(`Inserted ${results.affectedRows} employees into ${tableName}`);
            callback(null);
        }
    });
}

// Function to get the latest file in a folder based on the date in the file name
function getLatestFile(folderPath) {
    const files = fs.readdirSync(folderPath);
    let latestFileName = '';
    let latestDate = new Date(0); // Initialize with a very old date
    
    files.forEach(file => {
        const match = file.match(/_(\d{4}-\d{2}-\d{2})T/); // Extract the date from the file name
        if (match) {
            const fileDate = new Date(match[1]); // Convert the extracted date to a Date object
            if (fileDate > latestDate) {
                latestDate = fileDate;
                latestFileName = file;
            }
        }
    });
    
    return latestFileName;
}


// Clear data from respective tables before inserting new employees
clearTable('universal', (error) => {
    if (!error) {
        clearTable('islands', (error) => {
            if (!error) {
                clearTable('volcano', (error) => {
                    if (!error) {
                        clearTable('citywalk', (error) => {
                            if (!error) {
                                clearTable('notParkBased', (error) => {
                                    if (!error) {
                                        // Get the current directory of the script
                                        const currentDir = __dirname;

                                        // Specify the folder path where your XLSX files are located (media folder)
                                        const folderPath = path.join(currentDir, '../media');

                                        // Get the latest XLSX file in the media folder
                                        const latestFile = getLatestFile(folderPath);

                                        // Load the latest XLSX file
                                        const workbook = XLSX.readFile(path.join(folderPath, latestFile));

                                        // Choose the first sheet in the workbook
                                        const sheetName = workbook.SheetNames[0];
                                        const sheet = workbook.Sheets[sheetName];

                                        // Rest of the code for processing Excel data and inserting employees goes here
                                    } else {
                                        console.error("Failed to clear data from 'notParkBased'");
                                    }
                                });
                            } else {
                                console.error("Failed to clear data from 'citywalk'");
                            }
                        });
                    } else {
                        console.error("Failed to clear data from 'volcano'");
                    }
                });
            } else {
                console.error("Failed to clear data from 'islands'");
            }
        });
    } else {
        console.error("Failed to clear data from 'universal'");
    }
});


// Get the current directory of the script
const currentDir = __dirname;

// Specify the folder path where your XLSX files are located (media folder)
const folderPath = path.join(currentDir, '../media');

// Get the latest XLSX file in the media folder
const latestFile = getLatestFile(folderPath);

// Load the latest XLSX file
const workbook = XLSX.readFile(path.join(folderPath, latestFile));

// Choose the first sheet in the workbook
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// Find the range of rows containing "Employee" and "Totals"
const range = XLSX.utils.decode_range(sheet['!ref']);
let startRow = -1;
let endRow = -1;
for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 0 })];
    if (cell && cell.v && cell.v.toString().trim() === 'Employee') {
        startRow = rowNum + 1; // Start from the row after "Employee"
    }
    if (cell && cell.v && cell.v.toString().trim() === 'Totals') {
        endRow = rowNum; // Stop before "Totals"
        break;
    }
}

// Create arrays to store employees based on work location
const UNIVERSAL = [];
const ISLANDS = [];
const VolcanoBay = [];
const CityWalk = [];
const NotParkBased = [];

// Regular expression pattern to extract time and location
const scheduleRegex = /(\d{1,2}:\d{2}(?:AM|PM))\s*-\s*(\d{1,2}:\d{2}(?:AM|PM))\s*(.*)/i;

// Function to calculate total shift duration in minutes
function calculateTotalShift(employee) {
    // Convert time in and time out to minutes
    const timeInParts = employee.TimeIn.split(':').map(part => parseInt(part));
    const timeOutParts = employee.TimeOut.split(':').map(part => parseInt(part));
    let timeInMinutes = timeInParts[0] * 60 + timeInParts[1];
    let timeOutMinutes = timeOutParts[0] * 60 + timeOutParts[1];
    
    // If TimeOut is before TimeIn (e.g., TimeIn: 8:00AM, TimeOut: 4:30PM),
    // adjust TimeOut to represent the next day
    if (timeOutMinutes < timeInMinutes) {
        timeOutMinutes += 24 * 60; // Add 24 hours to represent next day
    }
    
    // Calculate total shift duration in minutes
    const totalShiftMinutes = timeOutMinutes - timeInMinutes;
    return totalShiftMinutes;
}

// Iterate through the rows to extract employee information and arrange them based on work location
for (let rowNum = startRow; rowNum < endRow; rowNum++) {
    const employee = {};
    // Extract data from columns A, B, D, J, L, and N
    employee.Name = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 0 })]?.v || ''; // Column A (Name)
    employee.Job = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 1 })]?.v || ''; // Column B (Job)
    const scheduleText = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 3 })]?.v || ''; // Column D (Schedule)
    employee.Notes = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 13 })]?.v || ''; // Column N (Notes)

    // Extracting location from the schedule text
    const locationMatch = scheduleText.match(/(USF|CW|VB|IOA)/i);
    employee.Location = locationMatch ? locationMatch[1].toUpperCase() : ''; // Extracted location code

    // Extracting time in and time out from the schedule text
    const timeParts = scheduleText.match(/(\d{1,2}:\d{2}(?:AM|PM))\s*-\s*(\d{1,2}:\d{2}(?:AM|PM))/i);
    if (timeParts && timeParts.length === 3) {
        employee.TimeIn = moment(timeParts[1], 'h:mm A').format('HH:mm'); // Convert time to 24-hour format
        employee.TimeOut = moment(timeParts[2], 'h:mm A').format('HH:mm'); // Convert time to 24-hour format

        // Calculate shift duration in minutes
        const timeInMinutes = getMinutesFromTimeString(employee.TimeIn);
        const timeOutMinutes = getMinutesFromTimeString(employee.TimeOut);
        const shiftDuration = timeOutMinutes - timeInMinutes;

        // Add 30 minutes to TimeOut if shift duration is longer than 6 hours
        if (shiftDuration > 360) { // 360 minutes = 6 hours
            employee.TimeOut = moment(employee.TimeOut, 'HH:mm').add(30, 'minutes').format('HH:mm');
        }
    } else {
        // If unable to extract, set defaults
        employee.TimeIn = '';
        employee.TimeOut = '';
    }

    // Function to convert time string to minutes
    function getMinutesFromTimeString(timeString) {
        const [hours, minutes] = timeString.split(':'); // Split time into hours and minutes
        return parseInt(hours) * 60 + parseInt(minutes); // Convert hours and minutes to total minutes
    }

    // Check if the job title is Lead-Paramedic and set the lead status accordingly
    employee.StatusLead = (employee.Job === 'Lead-Paramedic');

    // Calculate total shift duration in minutes and add it to the employee object
    employee.totalShift = calculateTotalShift(employee);
    
    // Exclude employees with totalShift === 30
    if (employee.totalShift > 30) {
        // Push the employee to the corresponding array
        switch (employee.Location) {
            case 'USF':
                UNIVERSAL.push(employee);
                break;
            case 'IOA':
                ISLANDS.push(employee);
                break;
            case 'VB':
                VolcanoBay.push(employee);
                break;
            case 'CW':
                CityWalk.push(employee);
                break;
            default:
                NotParkBased.push(employee);
                break;
        }
    }
}




// Function to insert employees into the database
function insertEmployees(employees, tableName) {
    const columns = ['EmployeeName', 'LeadStatus', 'TimeIn', 'TimeOut', 'Incentive', 'Notes'];
    const values = employees.map(employee => [employee.Name, employee.StatusLead ? 'Yes' : 'No', employee.TimeIn, employee.TimeOut, '', employee.Notes]);
    const sql = `INSERT INTO ${tableName} (${columns.join(', ')}) VALUES ?`;

    pool.query(sql, [values], (error, results, fields) => {
        if (error) {
            console.error(`Error inserting employees into ${tableName}: ${error.message}`);
        } else {
            console.log(`Inserted ${results.affectedRows} employees into ${tableName}`);
        }
    });
}

// Clear data from respective tables before inserting new employees
// clearTable('universal');
// clearTable('volcano');
// clearTable('citywalk');
// clearTable('islands');
// clearTable('notParkBased');

// Wait for 3 seconds before inserting employees
setTimeout(() => {
    // Insert employees into respective tables
    insertEmployees(UNIVERSAL, 'universal');
    insertEmployees(VolcanoBay, 'volcano');
    insertEmployees(CityWalk, 'citywalk');
    insertEmployees(ISLANDS, 'islands');
    insertEmployees(NotParkBased, 'notParkBased');

    // Call function to close the database connection after 6 seconds
    setTimeout(closeConnection, 4000); // 4000 milliseconds = 4 seconds
}, 3000); // 3000 milliseconds = 3 seconds

// Function to close the database connection
function closeConnection() {
    pool.end((error) => {
        if (error) {
            console.error(`Error closing the database connection: ${error.message}`);
        } else {
            console.log('Database connection closed.');
        }
    });
}


// Print the arrays of employees based on work location to the console
console.log('Employees at UNIVERSAL:');
console.log(UNIVERSAL);
console.log('Employees at ISLANDS:');
console.log(ISLANDS);
console.log('Employees at VolcanoBay:');
console.log(VolcanoBay);
console.log('Employees at CityWalk:');
console.log(CityWalk);
console.log('Employees at Not Park Based:');
console.log(NotParkBased);
console.log(`Retrieved latest file: ${latestFile}`); // Log the name of the retrieved file