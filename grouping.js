const xlsx = require('xlsx');
const _ = require('lodash');
const fs = require('fs');
const { ChartJSNodeCanvas } = require('chartjs-node-canvas');

// Load the Excel file
const workbook = xlsx.readFile('data.xlsx');
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convert sheet to JSON
const jsonData = xlsx.utils.sheet_to_json(worksheet);

// Function to sanitize "Göreviniz nedir?" column to allow only alphabetic characters (including Turkish alphabet) and spaces
const sanitizeJobTitle = (jobTitle) => {
    return jobTitle.replace(/[^a-zA-ZçğıİöşüÇĞİÖŞ\s]/g, '').trim(); // Allows Turkish characters and spaces
};

// Grouping function
const groupedData = _.groupBy(jsonData, row => 
    `${row['Kendinizi ne olarak tanımlarsınız?']} | ${row['Tecrübe yılınız ?']} | ${sanitizeJobTitle(row['Göreviniz nedir?'])} | ${row['Şirket  kadar büyük?']}`
);

// Sort each group by 'Maaş aralığınız?'
Object.keys(groupedData).forEach(key => {
    groupedData[key] = _.sortBy(groupedData[key], row => row['Maaş aralığınız?']);
});

// Generate charts for each group
const width = 800;
const height = 600;
const chartJSNodeCanvas = new ChartJSNodeCanvas({ width, height });

// Function to sanitize filenames (reduce length and replace invalid characters)
const sanitizeFilename = (name) => {
    const sanitized = name.replace(/[\\/:*?"<>|]/g, ''); // Replace invalid filename characters
    return sanitized.length > 100 ? sanitized.slice(0, 100) : sanitized; // Truncate to 100 characters
};

// Generate chart and save
const generateChart = async (groupKey, data) => {
    const labels = data.map(row => row['Maaş aralığınız?']);
    const values = data.map((_, index) => index + 1);

    const configuration = {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Maaş Dağılımı',
                data: values,
                backgroundColor: 'rgba(75, 192, 192, 0.2)',
                borderColor: 'rgba(75, 192, 192, 1)',
                borderWidth: 1
            }]
        },
        options: {
            responsive: false,
            scales: {
                y: { beginAtZero: true }
            }
        }
    };

    const imageBuffer = await chartJSNodeCanvas.renderToBuffer(configuration);

    // Extract individual grouping parts to create directories
    const [identity, experience, sanitizedJob, companySize] = groupKey.split(' | ');
    const directoryPath = `charts_grouped/${sanitizeFilename(identity)}/${sanitizeFilename(companySize)}/${sanitizeFilename(experience)}`;

    // Ensure directories exist
    fs.mkdirSync(directoryPath, { recursive: true });

    // Sanitize the file name and avoid path overflow
    const safeFilename = sanitizeFilename(groupKey);
    fs.writeFileSync(`${directoryPath}/${safeFilename}.png`, imageBuffer);
};

// Ensure root charts directory exists
fs.mkdirSync('charts_grouped', { recursive: true });

(async () => {
    for (const [key, value] of Object.entries(groupedData)) {
        await generateChart(key, value);
    }
    console.log('Charts generated and saved in the charts folder.');
})();

// Save grouped data to a JSON file
fs.writeFileSync('groupedData.json', JSON.stringify(groupedData, null, 2), 'utf-8');

console.log('Data has been grouped, sorted, and saved to groupedData.json');
