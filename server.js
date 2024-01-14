const express = require('express');
const bodyParser = require('body-parser');
const excel = require('exceljs');
const app = express();
const PORT = 3000;

app.use(bodyParser.json());

app.get('/',(req, res) => {
    res.sendFile(__dirname + '/index.html');
})

app.post('/search', (req, res) => {
    const searchName = req.body.name.toLowerCase();
    const workbook = new excel.Workbook();
    workbook.xlsx.readFile("./data1.xlsx").then(() => {
        const worksheet = workbook.getWorksheet(1);
        let result = null;

        worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            const name = row.getCell(1).value.toLowerCase();
            if (name.includes(searchName)) {
                result = {
                    name: row.getCell(1).value,
                    type: row.getCell(2).value,
                    thickness: row.getCell(3).value,
                    db2: row.getCell(4).value,
                    db1: row.getCell(5).value,
                    db3: row.getCell(6).value,
                    gap: row.getCell(7).value,
                    distance: row.getCell(8).value,
                };
            }
        });

        res.json(result);
   });
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
