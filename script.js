const entries = [];
let chartDuration, chartCalories, chartWater, chartSleep;

// Calorie rates per minute for each exercise
const calorieRates = {
    'Running': 10,
    'Walking': 5,
    'Yoga': 4,
    'Gym': 8,
    'Cycling': 7
};

// Function to calculate and update calories
function calculateCalories() {
    const exercise = document.getElementById('exercise').value;
    const duration = Number(document.getElementById('duration').value || 0);
    
    if (exercise && duration > 0) {
        const rate = calorieRates[exercise];
        const calories = duration * rate;
        document.getElementById('calories').value = calories;
    } else {
        document.getElementById('calories').value = '';
    }
}

// Function to calculate recovery rate based on sleep hours
// Recovery % = (sleep/8)*100, capped at 100%
function calculateRecoveryRate() {
    const sleep = Number(document.getElementById('sleep').value || 0);
    
    if (sleep > 0) {
        // Optimal value: 8 hours sleep = 100% recovery
        const recovery = Math.min((sleep / 8) * 100, 100);
        document.getElementById('recoveryRate').value = recovery.toFixed(1);
    } else {
        document.getElementById('recoveryRate').value = '';
    }
}

// Add event listeners for automatic calculation
document.getElementById('exercise').addEventListener('change', calculateCalories);
document.getElementById('duration').addEventListener('input', calculateCalories);
document.getElementById('sleep').addEventListener('input', calculateRecoveryRate);

function buildChart(canvasId, label, data, color) {
    const ctx = document.getElementById(canvasId).getContext('2d');

    return new Chart(ctx, {
        type: 'line',
        data: {
            labels: entries.map(e => e.date),
            datasets: [{
                label,
                data,
                borderColor: color,
                backgroundColor: color.replace('1)', '0.2)'),
                fill: false,
                tension: 0.1
            }]
        },
        options: {
            responsive: false,
            maintainAspectRatio: false,
            scales: {
                y: { beginAtZero: true }
            },
            plugins: {
                legend: { display: true, position: 'top' }
            }
        }
    });
}

function renderChart() {
    const labels = entries.map(entry => entry.date);
    const durationData = entries.map(entry => entry.duration);
    const caloriesData = entries.map(entry => entry.calories);
    const waterData = entries.map(entry => entry.water);
    const sleepData = entries.map(entry => entry.sleep);

    if (chartDuration) chartDuration.destroy();
    if (chartCalories) chartCalories.destroy();
    if (chartWater) chartWater.destroy();
    if (chartSleep) chartSleep.destroy();

    chartDuration = buildChart('chartDuration', 'Duration (min)', durationData, 'rgba(75, 192, 192, 1)');
    chartCalories = buildChart('chartCalories', 'Calories', caloriesData, 'rgba(255, 99, 132, 1)');
    chartWater = buildChart('chartWater', 'Water (L)', waterData, 'rgba(54, 162, 235, 1)');
    chartSleep = buildChart('chartSleep', 'Sleep (hrs)', sleepData, 'rgba(255, 206, 86, 1)');
}

let excelFileName = null;

// Function to create Excel Workbook with data and chart images
async function createExcelWorkbook() {
    if (entries.length === 0) {
        alert('No data to export. Please add some entries first.');
        return null;
    }

    try {
        const ExcelJS = window.ExcelJS;
        const workbook = new ExcelJS.Workbook();

        // ============ SHEET 1: FITNESS DATA ============
        const wsData = workbook.addWorksheet('Fitness Data');
        
        // Add header row
        const headerRow = wsData.addRow(['Date', 'Exercise', 'Duration (min)', 'Calories', 'Water (L)', 'Sleep (hrs)', 'Recovery %']);
        headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4CAF50' } };
        headerRow.alignment = { horizontal: 'center', vertical: 'center' };

        // Add data rows
        entries.forEach(entry => {
            const row = wsData.addRow([
                entry.date,
                entry.exercise,
                entry.duration,
                entry.calories,
                entry.water,
                entry.sleep,
                entry.recoveryRate ? entry.recoveryRate.toFixed(1) : 0
            ]);
            row.alignment = { horizontal: 'center' };
        });

        // Set column widths
        wsData.columns = [
            { width: 12 },
            { width: 12 },
            { width: 15 },
            { width: 12 },
            { width: 12 },
            { width: 12 },
            { width: 12 }
        ];

        // ============ SHEET 2: CHARTS (with chart images) ============
        const wsCharts = workbook.addWorksheet('Charts');
        wsCharts.pageSetup = { paperSize: 9, orientation: 'landscape' };

        // Get canvas elements and convert to images
        const durationCanvas = document.getElementById('chartDuration');
        const caloriesCanvas = document.getElementById('chartCalories');
        const waterCanvas = document.getElementById('chartWater');
        const sleepCanvas = document.getElementById('chartSleep');

        try {
            // Add title
            const titleRow = wsCharts.addRow(['FITNESS METRICS CHARTS']);
            titleRow.font = { bold: true, size: 14, color: { argb: 'FF4CAF50' } };
            titleRow.alignment = { horizontal: 'center' };

            // Add some spacing
            wsCharts.addRow([]);
            wsCharts.addRow([]);

            // Convert canvas to image and add to worksheet
            if (durationCanvas) {
                const durationImage = durationCanvas.toDataURL('image/png');
                const imageId1 = workbook.addImage({
                    base64: durationImage,
                    extension: 'png',
                });
                wsCharts.addImage(imageId1, 'A4:H12');
            }

            // Add spacing
            wsCharts.addRow([]);
            wsCharts.addRow([]);
            wsCharts.addRow([]);
            wsCharts.addRow([]);
            wsCharts.addRow([]);

            if (caloriesCanvas) {
                const caloriesImage = caloriesCanvas.toDataURL('image/png');
                const imageId2 = workbook.addImage({
                    base64: caloriesImage,
                    extension: 'png',
                });
                wsCharts.addImage(imageId2, 'I4:P12');
            }

            // Add spacing for second row
            wsCharts.addRow([]);
            wsCharts.addRow([]);
            wsCharts.addRow([]);
            wsCharts.addRow([]);
            wsCharts.addRow([]);

            if (waterCanvas) {
                const waterImage = waterCanvas.toDataURL('image/png');
                const imageId3 = workbook.addImage({
                    base64: waterImage,
                    extension: 'png',
                });
                wsCharts.addImage(imageId3, 'A20:H28');
            }

            if (sleepCanvas) {
                const sleepImage = sleepCanvas.toDataURL('image/png');
                const imageId4 = workbook.addImage({
                    base64: sleepImage,
                    extension: 'png',
                });
                wsCharts.addImage(imageId4, 'I20:P28');
            }

        } catch (imgError) {
            console.warn('Could not add chart images:', imgError);
            // Continue even if images fail
        }

        // ============ SHEET 3: SUMMARY ============
        const wsSummary = workbook.addWorksheet('Summary');
        
        const titleRow2 = wsSummary.addRow(['HEALTH METRICS SUMMARY']);
        titleRow2.font = { bold: true, size: 14, color: { argb: 'FF4CAF50' } };
        titleRow2.alignment = { horizontal: 'center' };
        
        wsSummary.addRow([]); // Empty row
        
        const headerSummaryRow = wsSummary.addRow(['Metric', 'Total', 'Average', 'Min', 'Max']);
        headerSummaryRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
        headerSummaryRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4CAF50' } };
        headerSummaryRow.alignment = { horizontal: 'center' };

        // Duration stats
        const durTotal = entries.reduce((sum, e) => sum + e.duration, 0);
        const durAvg = (durTotal / entries.length).toFixed(2);
        const durMin = Math.min(...entries.map(e => e.duration));
        const durMax = Math.max(...entries.map(e => e.duration));
        const row1 = wsSummary.addRow(['Duration (min)', durTotal, parseFloat(durAvg), durMin, durMax]);
        row1.alignment = { horizontal: 'center' };

        // Calories stats
        const calTotal = entries.reduce((sum, e) => sum + e.calories, 0);
        const calAvg = (calTotal / entries.length).toFixed(2);
        const calMin = Math.min(...entries.map(e => e.calories));
        const calMax = Math.max(...entries.map(e => e.calories));
        const row2 = wsSummary.addRow(['Calories', calTotal, parseFloat(calAvg), calMin, calMax]);
        row2.alignment = { horizontal: 'center' };

        // Water stats
        const watTotal = entries.reduce((sum, e) => sum + e.water, 0);
        const watAvg = (watTotal / entries.length).toFixed(2);
        const watMin = Math.min(...entries.map(e => e.water));
        const watMax = Math.max(...entries.map(e => e.water));
        const row3 = wsSummary.addRow(['Water (L)', watTotal, parseFloat(watAvg), watMin, watMax]);
        row3.alignment = { horizontal: 'center' };

        // Sleep stats
        const sleepTotal = entries.reduce((sum, e) => sum + e.sleep, 0);
        const sleepAvg = (sleepTotal / entries.length).toFixed(2);
        const sleepMin = Math.min(...entries.map(e => e.sleep));
        const sleepMax = Math.max(...entries.map(e => e.sleep));
        const row4 = wsSummary.addRow(['Sleep (hrs)', sleepTotal, parseFloat(sleepAvg), sleepMin, sleepMax]);
        row4.alignment = { horizontal: 'center' };

        // Recovery stats
        const recAvg = (entries.reduce((sum, e) => sum + e.recoveryRate, 0) / entries.length).toFixed(2);
        const row5 = wsSummary.addRow(['Average Recovery %', '', parseFloat(recAvg), '', '']);
        row5.alignment = { horizontal: 'center' };

        wsSummary.columns = [{ width: 20 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 }];

        return workbook;

    } catch (error) {
        console.error('Error creating workbook:', error);
        alert('Error creating Excel file: ' + error.message);
        return null;
    }
}

// Export to Excel - Create new file
async function exportToExcel() {
    const wb = await createExcelWorkbook();
    if (!wb) return;

    try {
        const buffer = await wb.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        excelFileName = 'fitness-tracker-data.xlsx';
        link.download = excelFileName;
        link.click();
        URL.revokeObjectURL(url);
        alert('✅ Excel file exported successfully as "fitness-tracker-data.xlsx"\n\nYou can now use the "Update Excel" button to update this file with new data.');
    } catch (error) {
        console.error('Export error:', error);
        alert('Error exporting to Excel: ' + error.message);
    }
}

// Update Excel - Update existing file with new data
async function updateExcel() {
    if (!excelFileName) {
        alert('⚠️ Please export to Excel first using "Export to Excel" button');
        return;
    }

    const wb = await createExcelWorkbook();
    if (!wb) return;

    try {
        const buffer = await wb.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = excelFileName;
        link.click();
        URL.revokeObjectURL(url);
        alert('✅ Excel file updated successfully!\n\nYour downloads folder should ask to replace "fitness-tracker-data.xlsx"');
    } catch (error) {
        console.error('Update error:', error);
        alert('Error updating Excel: ' + error.message);
    }
}

// Add event listeners for export and update buttons
document.getElementById('exportBtn').addEventListener('click', exportToExcel);
document.getElementById('updateBtn').addEventListener('click', updateExcel);

function addEntryToTable(entry) {
    const table = document.getElementById('dataTable');
    const row = table.insertRow();

    row.insertCell(0).innerText = entry.date;
    row.insertCell(1).innerText = entry.exercise;
    row.insertCell(2).innerText = entry.duration;
    row.insertCell(3).innerText = entry.calories;
    row.insertCell(4).innerText = entry.water;
    row.insertCell(5).innerText = entry.sleep;
    row.insertCell(6).innerText = entry.recoveryRate.toFixed(1) + '%';
}

function populateSampleData(total = 100) {
    const exercises = Object.keys(calorieRates);
    const startDate = new Date('2026-01-01');

    for (let i = 1; i <= total; i++) {
        const exercise = exercises[(i - 1) % exercises.length];
        const duration = 20 + (i % 41); // 20 to 60 min
        const calories = duration * calorieRates[exercise];
        const water = Number((1.5 + ((i % 16) * 0.1)).toFixed(1)); // 1.5 to 3.0 L
        const sleep = Number((5.5 + ((i % 11) * 0.3)).toFixed(1)); // 5.5 to 8.5 hrs
        const recoveryRate = Number(Math.min((sleep / 8) * 100, 100).toFixed(1));

        const currentDate = new Date(startDate);
        currentDate.setDate(startDate.getDate() + i - 1);
        const date = currentDate.toISOString().split('T')[0];

        const entry = { date, exercise, duration, calories, water, sleep, recoveryRate };
        entries.push(entry);
        addEntryToTable(entry);
    }

    renderChart();
}

document.getElementById("fitnessForm").addEventListener("submit", function(e) {
    e.preventDefault();

    const date = document.getElementById("date").value;
    const exercise = document.getElementById("exercise").value;
    const duration = Number(document.getElementById("duration").value || 0);
    const calories = Number(document.getElementById("calories").value || 0);
    const water = Number(document.getElementById("water").value || 0);
    const sleep = Number(document.getElementById("sleep").value || 0);
    const recoveryRate = Number(document.getElementById("recoveryRate").value || 0);

    const entry = { date, exercise, duration, calories, water, sleep, recoveryRate };
    entries.push(entry);
    addEntryToTable(entry);

    document.getElementById("fitnessForm").reset();

    renderChart();
});

// initial chart setup with 100 records
populateSampleData(100);