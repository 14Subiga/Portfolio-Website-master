const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const fs = require('fs');

const app = express();
const upload = multer();

app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));

app.post('/submit', upload.none(), (req, res) => {
    const { name, email, message } = req.body;

    let workbook;
    const filePath = './contacts.xlsx';

    if (fs.existsSync(filePath)) {
        workbook = XLSX.readFile(filePath);
    } else {
        workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([]), 'Contacts');
    }

    const worksheet = workbook.Sheets['Contacts'];
    const data = XLSX.utils.sheet_to_json(worksheet);

    data.push({ Name: name, Email: email, Message: message });

    const newWorksheet = XLSX.utils.json_to_sheet(data);
    workbook.Sheets['Contacts'] = newWorksheet;
    XLSX.writeFile(workbook, filePath);

    res.json({ message: 'Contact details saved successfully!' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
