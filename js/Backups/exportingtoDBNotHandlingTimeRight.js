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

// Function to get the latest file in a folder
function getLatestFile(folderPath) {
    const files = fs.readdirSync(folderPath);
    const stats = files.map(file => ({
        file,
        mtime: fs.statSync(path.join(folderPath, file)).mtime
    }));
    const latestFile = stats.reduce((prev, current) => {
        return (current.mtime > prev.mtime) ? current : prev;
    });
    return latestFile.file;
}

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
    employee.TimeIn = timeParts[1];
    employee.TimeOut = timeParts[2];

    // Calculate shift duration in minutes
    const timeInMinutes = getMinutesFromTimeString(employee.TimeIn);
    const timeOutMinutes = getMinutesFromTimeString(employee.TimeOut);
    const shiftDuration = timeOutMinutes - timeInMinutes;

    // Add 30 minutes to TimeOut if shift duration is longer than 6 hours
    if (shiftDuration > 360) { // 360 minutes = 6 hours
        const [hours, minutes, period] = employee.TimeOut.split(/:|(?=[AP]M)/); // Split time and period (AM/PM)
        let newHours = parseInt(hours);
        let newMinutes = parseInt(minutes) + 30;
        if (newMinutes >= 60) {
            newMinutes -= 60;
            newHours = (newHours + 1) % 12 || 12; // Handle rollover and edge cases
        }
        employee.TimeOut = `${newHours}:${newMinutes.toString().padStart(2, '0')}${period}`;
    }
} else {
    // If unable to extract, set defaults
    employee.TimeIn = '';
    employee.TimeOut = '';
}

// Function to convert time string to minutes
function getMinutesFromTimeString(timeString) {
    const [hours, minutes, period] = timeString.split(/:|(?=[AP]M)/); // Split time and period (AM/PM)
    let totalMinutes = parseInt(hours) * 60 + parseInt(minutes);
    if (period.toUpperCase() === 'PM') {
        totalMinutes += 12 * 60; // Add 12 hours if PM
    }
    return totalMinutes;
}
    

    // Check if the job title is Lead-Paramedic and set the lead status accordingly
    if (employee.Job === 'Lead-Paramedic') {
        employee.StatusLead = true;
    } else {
        employee.StatusLead = false;
    }
    
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

// Insert employees into respective tables
insertEmployees(UNIVERSAL, 'universal');
insertEmployees(ISLANDS, 'islands');
insertEmployees(VolcanoBay, 'volcano');
insertEmployees(CityWalk, 'citywalk');
insertEmployees(NotParkBased, 'notParkBased');


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