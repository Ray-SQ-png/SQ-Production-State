// Function to load Excel file and update data and interface
async function loadExcelFile() {
    try {
        const response = await fetch('test.xlsx'); // Assuming test.xlsx is in the same directory
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const tableBody = document.querySelector('#production-table tbody');
        tableBody.innerHTML = '';
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        jsonData.forEach((rowData, rowIndex) => {
            if (rowIndex === 0) return;
            const row = document.createElement('tr');
            rowData.forEach((cellData, cellIndex) => {
                const cell = document.createElement(cellIndex === 0 ? 'th' : 'td');
                cell.textContent = cellData;
                row.appendChild(cell);
            });
            tableBody.appendChild(row);
        });
        displayTotalWorkers(jsonData);
        updateWorkingHoursBar(); // Update working hours bar when loading Excel file
    } catch (error) {
        console.error('Error loading Excel file:', error);
    }
}

// Function to update data and interface every 5 minutes
setInterval(() => {
    loadExcelFile();
}, 300000); // Update every 5 minutes (300,000 milliseconds)

// Initial setup
loadExcelFile();

// Function to handle file input change event
document.getElementById('file-input').addEventListener('change', function(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const tableBody = document.querySelector('#production-table tbody');
            tableBody.innerHTML = '';
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            jsonData.forEach((rowData, rowIndex) => {
                if (rowIndex === 0) return;
                const row = document.createElement('tr');
                rowData.forEach((cellData, cellIndex) => {
                    const cell = document.createElement(cellIndex === 0 ? 'th' : 'td');
                    cell.textContent = cellData;
                    row.appendChild(cell);
                });
                tableBody.appendChild(row);
            });
            displayTotalWorkers(jsonData);
            updateWorkingHoursBar(); // Update working hours bar when loading Excel file
        };
        reader.readAsArrayBuffer(file);
    }
});

// Function to update working hours bar
function updateWorkingHoursBar() {
    const now = new Date();
    const hours = now.getHours();
    const minutes = now.getMinutes();
    const bar = document.querySelector('#working-hours-bar');
    const hourMarker = document.querySelector('#current-hour');

    // Set working hours based on the month
    const isJulyOrAugust = now.getMonth() === 6 || now.getMonth() === 7;
    const startHour = isJulyOrAugust ? 7 : 8;
    const endHour = isJulyOrAugust ? 15 : 17;

    // Rest time from 12:00 PM to 12:30 PM
    if (hours < startHour || hours >= endHour || (hours === 12 && minutes >= 0 && minutes <= 29)) {
        bar.style.backgroundColor = 'red';
    } else {
        bar.style.backgroundColor = 'green';
    }

    const totalHours = endHour - startHour;
    const currentHourProgress = (hours - startHour + minutes / 60) / totalHours * 100;

    bar.style.width = `${currentHourProgress}%`;

    // Position hour marker inside the bar
    const barWidth = bar.offsetWidth;
    const hourPosition = currentHourProgress / 100 * barWidth;
    hourMarker.style.left = `${hourPosition}px`;

    // Display current hour inside the bar
    hourMarker.textContent = `${hours < 10 ? '0' + hours : hours}:${minutes < 10 ? '0' + minutes : minutes}`;

    // Display the date
    const dateSection = document.querySelector('#date-section');
    dateSection.textContent = now.toDateString();
}

// Function to display total number of workers
function displayTotalWorkers(data) {
    const totalWorkers = data.reduce((total, rowData, rowIndex) => {
        // Skip header row
        if (rowIndex === 0) return total;
        return total + rowData[1]; // Assuming the second column contains the number of workers
    }, 0);

    const totalWorkersElement = document.querySelector('#total-workers');
    totalWorkersElement.textContent = `Total Workers: ${totalWorkers}`;
}