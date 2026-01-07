const xlsx = require('xlsx');

// Create a workbook with 100 sample patients
// Columns: ID, FullName, BirthDate, Address
const data = [];
for (let i = 1; i <= 100; i++) {
    // Random birth year between 1990 and 2024
    const year = 1990 + Math.floor(Math.random() * 35);
    const month = Math.floor(Math.random() * 12) + 1;
    const day = Math.floor(Math.random() * 28) + 1;

    // Format nicely
    const birthDate = `${day}.${month}.${year}`;

    data.push({
        ID: i,
        "Ism Familiya": `Patient Test ${i}`,
        "Tug'ilgan sanasi": birthDate,
        "Manzil": `Street ${i}, City`
    });
}

const wb = xlsx.utils.book_new();
const ws = xlsx.utils.json_to_sheet(data);
xlsx.utils.book_append_sheet(wb, ws, "Patients");

xlsx.writeFile(wb, "sample_patients.xlsx");
console.log("sample_patients.xlsx created");
