const excel = require('exceljs');
const fs = require('fs');

const excelFilePath = 'orders.xlsx';
let workbook = new excel.Workbook();
let worksheet;

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

exports.handler = async function (event, context) {
    try {
        const orderDetails = JSON.parse(event.body);

        worksheet.addRow([
            orderDetails.bookId,
            orderDetails.name,
            orderDetails.phone,
            orderDetails.department,
            orderDetails.academicYear,
            orderDetails.paymentMethod
        ]);

        await workbook.xlsx.writeFile(excelFilePath);

        return {
            statusCode: 200,
            body: JSON.stringify({ status: 'success', message: 'Order placed successfully.' })
        };
    } catch (error) {
        console.error('Error placing order:', error);

        return {
            statusCode: 500,
            body: JSON.stringify({ status: 'error', message: 'Failed to place the order.' })
        };
    }
};
