const express = require('express');
const sgMail = require('@sendgrid/mail');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

sgMail.setApiKey(process.env.SENDGRID_API_KEY);

app.post('/send-manifest', async (req, res) => {

    if (req.body.ping) return res.json({ success: true });

    const { operator, courier, awbs, count } = req.body;
    const todayDate = new Date().toLocaleDateString('en-IN');

    // Instant response
    res.json({ success: true });

    try {
        // =========================
        // ✅ EXCEL GENERATION
        // =========================
        const excelData = [["SL No.", "Date", "Courier Name", "AWB NUMBER"]];
        awbs.forEach((a, i) => excelData.push([i + 1, todayDate, courier, a]));

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(excelData);
        XLSX.utils.book_append_sheet(wb, ws, "Manifest");

        const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

        // =========================
        // ✅ PDF GENERATION (MULTI-COLUMN)
        // =========================
        const doc = new PDFDocument({ size: 'A4', margin: 40, layout: 'portrait' });
        let buffers = [];
        doc.on('data', buffers.push.bind(buffers));

        // HEADER
        doc.fontSize(20).font('Helvetica-Bold').text("MEDIKA BAZAAR", { align: 'center' });
        doc.moveDown(0.5);
        doc.fontSize(28).text(`${courier.toUpperCase()} MANIFEST`, { align: 'center' });
        doc.moveDown(0.5);
        doc.fontSize(14).text(`Date: ${todayDate}`, { align: 'left' });
        doc.text(`Operator: ${operator}`, { align: 'left' });
        doc.fontSize(16).text(`TOTAL AWB: ${count}`, { align: 'right' });
        doc.moveDown(1);

        // TABLE CONFIG
        const startX = doc.page.margins.left;
        const startY = doc.y;
        const pageWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
        const pageHeight = doc.page.height - doc.page.margins.bottom - startY;
        const rowHeight = 20;

        const maxRowsPerCol = Math.floor(pageHeight / rowHeight) - 1;
        const totalCols = Math.ceil(awbs.length / maxRowsPerCol);
        const colWidth = pageWidth / totalCols;

        let colX = startX;

        for (let c = 0; c < totalCols; c++) {
            const colAWBs = awbs.slice(c * maxRowsPerCol, (c + 1) * maxRowsPerCol);
            let y = startY;

            // HEADER FOR EACH COLUMN
            doc.rect(colX, y, colWidth * 0.2, rowHeight).fillAndStroke('#d9d9d9', '#000');
            doc.rect(colX + colWidth * 0.2, y, colWidth * 0.8, rowHeight).fillAndStroke('#d9d9d9', '#000');

            doc.fillColor('black').font('Helvetica-Bold');
            doc.text("SL", colX, y + 5, { width: colWidth * 0.2 - 5, align: 'center' });
            doc.text("AWB NUMBER", colX + colWidth * 0.2, y + 5, { width: colWidth * 0.8 - 5, align: 'center' });

            y += rowHeight;

            // ROWS
            doc.font('Helvetica').fontSize(10);
            colAWBs.forEach((awb, i) => {
                // Zebra background
                if (i % 2 === 0) doc.rect(colX, y, colWidth, rowHeight).fill('#f5f5f5');

                // Borders
                doc.strokeColor('#000').lineWidth(1);
                doc.rect(colX, y, colWidth * 0.2, rowHeight).stroke();
                doc.rect(colX + colWidth * 0.2, y, colWidth * 0.8, rowHeight).stroke();

                // Text
                doc.fillColor('black');
                doc.text(i + 1 + c * maxRowsPerCol, colX, y + 5, { width: colWidth * 0.2 - 5, align: 'center' });
                doc.text(awb, colX + colWidth * 0.2 + 5, y + 5, { width: colWidth * 0.8 - 5, align: 'left' });

                y += rowHeight;
            });

            colX += colWidth;
        }

        // FOOTER
        doc.fontSize(9).text(
            `Printed: ${new Date().toLocaleString('en-IN')} | Medika Logistics Portal`,
            doc.page.margins.left,
            doc.page.height - 40,
            { align: 'center', width: pageWidth }
        );

        doc.end();

        const pdfBuffer = await new Promise(resolve => {
            doc.on('end', () => resolve(Buffer.concat(buffers)));
        });

        // =========================
        // SEND EMAIL
        // =========================
        const msg = {
            to: [
                'rajeshrak413@outlook.com',
                'gokulkrishnan.velayutham@medikabazaar.com',
                'bhaskar.r@medikabazaar.com',
                'hanumanta.madival@medikabazaar.com'
            ],
            cc: ['elumalai.b@medikabazaar.com'],
            from: { email: 'Rajeshrak413@gmail.com', name: 'Medika Logistics Portal' },
            subject: `Outbound Manifest - ${courier} - ${todayDate}`,
            text: `Hello,

Please find the outbound manifest details below:

Date :- ${todayDate}
Courier Name :- ${courier}
Operator Name :- ${operator}
Total AWB :- ${count}

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
        console.log(`✅ Manifest sent for ${courier}`);

    } catch (error) {
        console.error("❌ ERROR:", error.response?.body || error);
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));