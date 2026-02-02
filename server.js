const express = require('express');
const sgMail = require('@sendgrid/mail');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

sgMail.setApiKey(process.env.SENDGRID_API_KEY);

// =====================================================
// SEND MANIFEST API
// =====================================================
app.post('/send-manifest', async (req, res) => {

    if (req.body.ping) return res.json({ success: true });

    const { operator, courier, awbs, count } = req.body;
    const todayDate = new Date().toLocaleDateString('en-IN');

    res.json({ success: true });

    try {

        // =====================================================
        // EXCEL GENERATION
        // =====================================================
        const excelData = [["SL No.", "Date", "Courier", "AWB"]];

        awbs.forEach((a, i) => {
            excelData.push([i + 1, todayDate, courier, a]);
        });

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(excelData);
        XLSX.utils.book_append_sheet(wb, ws, "Manifest");

        const excelBuffer = XLSX.write(wb, {
            type: 'buffer',
            bookType: 'xlsx'
        });

        // =====================================================
        // ULTIMATE PRO PDF GENERATION
        // =====================================================
        const doc = new PDFDocument({
            size: 'A4',
            layout: 'landscape',
            margin: 30
        });

        let pdfBuffers = [];
        doc.on('data', pdfBuffers.push.bind(pdfBuffers));

        // HEADER
        doc.fontSize(20).font('Helvetica-Bold')
        .text('MEDIKA BAZAAR', { align: 'center' });

        doc.moveDown(0.5);

        doc.fontSize(30).font('Helvetica-Bold')
        .text(courier.toUpperCase(), { align: 'center' });

        doc.moveDown(0.5);

        doc.fontSize(16).font('Helvetica-Bold')
        .text(`TOTAL AWB : ${awbs.length}`, { align: 'center' });

        doc.moveDown(1);

        // AUTO FONT
        let fontSize = 11;
        if (awbs.length > 120) fontSize = 10;
        if (awbs.length > 180) fontSize = 9;
        if (awbs.length > 250) fontSize = 8;

        doc.fontSize(fontSize);

        // TABLE CONFIG
        const pageWidth =
            doc.page.width -
            doc.page.margins.left -
            doc.page.margins.right;

        const startX = doc.page.margins.left;
        const tableTop = doc.y;
        const rowHeight = fontSize + 10;

        const slWidth = pageWidth * 0.15;
        const awbWidth = pageWidth * 0.85;

        const tableHeight = rowHeight * (awbs.length + 1);

        // OUTER BORDER
        doc.lineWidth(2)
        .rect(startX, tableTop, slWidth + awbWidth, tableHeight)
        .stroke();

        doc.lineWidth(1);

        // HEADER ROW
        doc.rect(startX, tableTop, slWidth, rowHeight)
        .fillAndStroke('#d9d9d9', '#000');

        doc.rect(startX + slWidth, tableTop, awbWidth, rowHeight)
        .fillAndStroke('#d9d9d9', '#000');

        doc.fillColor('black').font('Helvetica-Bold');

        doc.text("SL No", startX, tableTop + 6, {
            width: slWidth,
            align: 'center'
        });

        doc.text("AWB Number", startX + slWidth, tableTop + 6, {
            width: awbWidth,
            align: 'center'
        });

        doc.font('Helvetica');

        // ROWS
        let y = tableTop + rowHeight;

        awbs.forEach((awb, i) => {

            if (i % 2 === 0) {
                doc.rect(startX, y, slWidth + awbWidth, rowHeight)
                .fill('#f5f5f5')
                .fillColor('black');
            }

            doc.rect(startX, y, slWidth, rowHeight).stroke();
            doc.rect(startX + slWidth, y, awbWidth, rowHeight).stroke();

            doc.text(i + 1, startX, y + 6, {
                width: slWidth,
                align: 'center'
            });

            doc.text(awb, startX + slWidth + 5, y + 6, {
                width: awbWidth - 10,
                align: 'left'
            });

            y += rowHeight;

            if (y > doc.page.height - 70) return;
        });

        // FOOTER
        const printTime = new Date().toLocaleString('en-IN');

        doc.fontSize(9).font('Helvetica')
        .text(
            `Printed: ${printTime} | System: Medika Logistics Portal`,
            doc.page.margins.left,
            doc.page.height - 40,
            { align: 'center', width: pageWidth }
        );

        doc.end();

        const pdfBuffer = await new Promise(resolve => {
            doc.on('end', () => resolve(Buffer.concat(pdfBuffers)));
        });

        // =====================================================
        // EMAIL SEND
        // =====================================================
        const msg = {
            to: [
                'rajeshrak413@outlook.com'
            ],
            cc: [],

            from: {
                email: 'Rajeshrak413@gmail.com',
                name: 'Medika Logistics Portal'
            },

            subject: `Outbound Manifest - ${courier} - ${todayDate}`,

            text: `Outbound Manifest

Courier: ${courier}
Operator: ${operator}
Total AWB: ${count}

Excel + PDF Attached`,

            attachments: [
                {
                    content: excelBuffer.toString('base64'),
                    filename: `${courier}_Manifest_${todayDate.replace(/\//g, '-')}.xlsx`,
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    disposition: 'attachment'
                },
                {
                    content: pdfBuffer.toString('base64'),
                    filename: `${courier}_Manifest_${todayDate.replace(/\//g, '-')}.pdf`,
                    type: 'application/pdf',
                    disposition: 'attachment'
                }
            ]
        };

        await sgMail.send(msg);

        console.log("✅ Email Sent With Excel + Ultimate PDF");

    } catch (err) {
        console.log("❌ ERROR:", err.response?.body || err);
    }
});

// =====================================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("🚀 Server Running on " + PORT));