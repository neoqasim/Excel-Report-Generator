const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const ExcelJS = require('exceljs');
const app = express();
const port = 5000;

app.use(cors());
app.use(bodyParser.json());

app.post('/download', async (req, res) => {
  const { reports, rowStyles } = req.body;

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Reports');

  worksheet.columns = [
    { header: 'Date of Service', key: 'dos', width: 16, style: { numFmt: 'MM/DD/YYYY' } },
    { header: 'Name', key: 'fname', width: 18 },
    // { header: 'Last Name', key: 'lname', width: 18 },
    { header: 'DOB', key: 'dob', width: 16, style: { numFmt: 'MM/DD/YYYY' } },
    { header: 'Insurance Name', key: 'insname', width: 25 },
    { header: 'Provider Name', key: 'repName', width: 21 },
    { header: 'Claim Status', key: 'status', width: 22 },
    { header: 'Reason for Denial', key: 'denialReason', width: 25 },
    { header: 'Claim Number', key: 'cNumber', width: 27 },
    { header: 'Check Number', key: 'chNumber', width: 24 },
    { header: 'Reference Number', key: 'refNumber', width: 24 },
    // { header: 'Received Date', key: 'rdate', width: 16, style: { numFmt: 'MM/DD/YYYY' } },
    // { header: 'Processed Date', key: 'pdate', width: 16, style: { numFmt: 'MM/DD/YYYY' } },
    { header: 'Notes', key: 'notes', width: 30 },
  ];

  // Apply styles to the header row
  worksheet.getRow(1).eachCell({ includeEmpty: true }, cell => {
    cell.alignment = { wrapText: true };
    cell.font = { 
      bold: true, 
      size: 15, 
      color: { argb: '000000' } 
    };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'C4ECFD' } 
    };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  });

  // Apply styles to data rows
  reports.forEach((report, index) => {
    const row = worksheet.addRow(report);

    row.eachCell({ includeEmpty: true }, cell => {
      cell.alignment = { wrapText: true };
      cell.font = {
        size: 12,
        color: { argb: '000000' },
        bold: false,
      };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF' }
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
  });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=reports.xlsx');

  await workbook.xlsx.write(res);
  res.end();
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
