const express = require('express');
const sgMail = require('@sendgrid/mail');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

// SendGrid API
sgMail.setApiKey(process.env.SENDGRID_API_KEY);

// ===================================================
// SEND MANIFEST ROUTE
// ===================================================
app.post('/send-manifest', async (req, res) => {

    if (req.body.ping) return res.status(200).json({ success: true });

    const { operator, courier, awbs, count } = req.body;
    const todayDate = new Date().toLocaleDateString('en-IN');

    // Instant response to frontend
    res.status(200).json({ success: true });

    try {

        // ===================================================
        // 1️⃣ EXCEL GENERATION
        // ===================================================
        const excelData = [["SL No.", "Date", "Courier Name", "AWB NUMBER"]];

        awbs.forEach((awb, i) => {
            excelData.push([i + 1, todayDate, courier, awb]);
        });

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(excelData);
        XLSX.utils.book_append_sheet(wb, ws, "Manifest");

        const excelBuffer = XLSX.write(wb, {
            type: 'buffer',
            bookType: 'xlsx'
        });

        // ===================================================
        // 2️⃣ PDF GENERATION (SL + AWB ONLY)
        // ===================================================

        const doc = new PDFDocument({
            size: 'A4',
            margin: 25
        });

        let pdfBuffers = [];
        doc.on('data', pdfBuffers.push.bind(pdfBuffers));

        // ===== HEADER =====
        doc.fontSize(16).text('MEDIKA BAZAAR', { align: 'center' });
        doc.fontSize(12).text('Outbound AWB Manifest', { align: 'center' });

        doc.moveDown(0.5);
        doc.fontSize(10);
        doc.text(`Date : ${todayDate}`);
        doc.text(`Courier : ${courier}`);
        doc.text(`Operator : ${operator}`);
        doc.text(`Total AWB : ${count}`);

        doc.moveDown();

        // ===== AUTO FONT FIT =====
        let fontSize = 10;
        if (awbs.length > 80) fontSize = 9;
        if (awbs.length > 120) fontSize = 8;
        if (awbs.length > 160) fontSize = 7;

        doc.fontSize(fontSize);

        // ===== AUTO COLUMN WIDTH =====
        const pageWidth =
            doc.page.width -
            doc.page.margins.left -
            doc.page.margins.right;

        const getWidth = (t) => doc.widthOfString(String(t));

        let slWidth = getWidth("SL No") + 15;
        let awbWidth = getWidth("AWB Number") + 20;

        awbs.forEach((awb, i) => {
            slWidth = Math.max(slWidth, getWidth(i + 1) + 15);
            awbWidth = Math.max(awbWidth, getWidth(awb) + 20);
        });

        // Scale if exceeding page
        let totalWidth = slWidth + awbWidth;
        if (totalWidth > pageWidth) {
            let ratio = pageWidth / totalWidth;
            slWidth *= ratio;
            awbWidth *= ratio;
        }

        // Column positions
        let xSL = doc.page.margins.left;
        let xAWB = xSL + slWidth;

        // ===== TABLE HEADER =====
        doc.fontSize(fontSize + 1);
        doc.text("SL No", xSL, doc.y);
        doc.text("AWB Number", xAWB, doc.y);

        doc.moveDown(0.5);
        doc.fontSize(fontSize);

        // ===== TABLE DATA =====
        awbs.forEach((awb, i) => {
            doc.text(i + 1, xSL);
            doc.text(awb, xAWB);
        });

        doc.end();

        const pdfBuffer = await new Promise(resolve => {
            doc.on('end', () => resolve(Buffer.concat(pdfBuffers)));
        });

        // ===================================================
        // 3️⃣ EMAIL SEND
        // ===================================================

        const msg = {
            to: [
                'rajeshrak413@outlook.com' ],
            cc: [''],

            from: {
                email: 'Rajeshrak413@gmail.com',
                name: 'Medika Logistics Portal'
            },

            subject: `Outbound Manifest - ${courier} - ${todayDate}`,

            text: `Hello,

Please find the outbound manifest details below:

Date :- ${todayDate}
Courier Name :- ${courier}
Operator Name :- ${operator}
Total count of AWB :- ${count}

Attached: Excel + PDF Manifest.

Regards,
Outbound`,

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

        console.log(`✅ Excel + PDF Sent Successfully for ${courier}`);

    } catch (error) {
        console.error("❌ Send Error:");
        console.error(error.response?.body || error);
    }
});

// ===================================================
// SERVER START
// ===================================================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Server running on port ${PORT}`);
});