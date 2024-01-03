// netlify-functions/place-order.js
const express = require('express');
const bodyParser = require('body-parser');
const excel = require('exceljs');
const fs = require('fs');
const cors = require('cors');

const app = express();
const excelFilePath = 'orders.xlsx';

// Middleware to enable CORS
app.use(cors());

// Middleware to parse JSON
app.use(bodyParser.json());

// Create or load the Excel workbook and worksheet
let workbook = new excel.Workbook();
let worksheet;

// Check if the Excel file exists
if (fs.existsSync(excelFilePath)) {
  workbook.xlsx.readFile(excelFilePath)
    .then(() => {
      worksheet = workbook.getWorksheet(1);
      console.log('Existing Excel file loaded.');
    })
    .catch(error => {
      console.error('Error loading Excel file:', error);
    });
} else {
  worksheet = workbook.addWorksheet('Orders');
  worksheet.addRow(['Book ID', 'Name', 'Phone', 'Department', 'Academic Year', 'Payment Method']);
  console.log('New Excel file created.');
}

// Endpoint to handle the order placement
app.post('/netlify/functions/place-order', (req, res) => {
  const orderDetails = req.body;

  worksheet.addRow([
    orderDetails.bookId,
    orderDetails.name,
    orderDetails.phone,
    orderDetails.department,
    orderDetails.academicYear,
    orderDetails.paymentMethod
  ]);

  workbook.xlsx.writeFile(excelFilePath)
    .then(() => {
      console.log('Order details written to Excel file.');
      res.json({ status: 'success', message: 'Order placed successfully.' });
    })
    .catch(error => {
      console.error('Error writing to Excel file:', error);
      res.status(500).json({ status: 'error', message: 'Failed to place the order.' });
    });
});

module.exports = app;
