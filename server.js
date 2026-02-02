const express = require('express');
const sgMail = require('@sendgrid/mail');
const XLSX = require('xlsx');
const cors = require('cors');
const PDFDocument = require('pdfkit');

const app = express();
app.use(cors());
app.use(express.json());

sgMail.setApiKey(process.env.SENDGRID_API_KEY);

app.post('/send-manifest', async (req, res) => {

    if (req.body.ping) return res.status(200).json({ success: true });

    const { operator, courier, awbs, count } = req.body;
    const todayDate = new Date().toLocaleDateString('en-IN');

    res.status(200).json({ success: true });

    try {

        // =========================
        // ✅ EXCEL GENERATION
        // =========================
        const data = [["SL No.", "Date", "Courier Name", "AWB NUMBER"]];
        awbs.forEach((a, i) => data.push([i + 1, todayDate, courier, a]));

        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, "Manifest");

        const excelBuffer = XLSX.write(wb, {
            type: 'buffer',
            bookType: 'xlsx'
        });

        // =========================
        // ✅ PDF GENERATION (BOX STYLE)
        // =========================
        const doc = new PDFDocument({ margin: 40 });

        let buffers = [];
        doc.on('data', buffers.push.bind(buffers));

        // ===== COMPANY HEADER =====
        doc.fontSize(20).text("MEDIKA BAZAAR", { align: 'center' });
        doc.moveDown(0.5);

        // ===== COURIER BIG NAME =====
        doc.fontSize(18).text(`${courier} MANIFEST`, { align: 'center' });

        doc.moveDown(0.5);

        // ===== DATE + TOTAL =====
        doc.fontSize(12).text(`Date : ${todayDate}`);
        doc.text(`Operator : ${operator}`);
        doc.fontSize(14).text(`TOTAL AWB : ${count}`, { align: 'right' });

        doc.moveDown(1);

        // ===== TABLE SETTINGS =====
        const startX = 40;
        let startY = doc.y;

        const pageWidth = doc.page.width - 80;

        // AUTO COLUMN FIT
        const col1Width = 70; // SL
        const col2Width = pageWidth - col1Width; // AWB

        const rowHeight = 25;

        // ===== TABLE HEADER =====
        drawRow("SL NO", "AWB NUMBER", true);

        // ===== AWB DATA =====
        awbs.forEach((awb, i) => {
            drawRow(i + 1, awb, false);
        });

        doc.end();

        const pdfBuffer = await new Promise(resolve => {
            doc.on('end', () => resolve(Buffer.concat(buffers)));
        });

        // =========================
        // ✅ SEND EMAIL
        // =========================
        const msg = {
            to: [
                'rajeshrak413@outlook.com'
            ],
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

        console.log(`✅ Manifest Sent for ${courier}`);

        // =========================
        // DRAW TABLE FUNCTION
        // =========================
        function drawRow(sl, awb, isHeader) {

            if (startY > doc.page.height - 80) {
                doc.addPage();
                startY = 50;
            }

            // BORDER BOX SL
            doc.rect(startX, startY, col1Width, rowHeight).stroke();

            // BORDER BOX AWB
            doc.rect(startX + col1Width, startY, col2Width, rowHeight).stroke();

            // TEXT
            doc.fontSize(isHeader ? 12 : 11);
            doc.text(sl.toString(), startX + 5, startY + 7, {
                width: col1Width - 10,
                align: 'center'
            });

            doc.text(awb.toString(), startX + col1Width + 5, startY + 7, {
                width: col2Width - 10,
                align: 'center'
            });

            startY += rowHeight;
        }

    } catch (error) {

        console.error("❌ SendGrid Error:");

        if (error.response) {
            console.error(JSON.stringify(error.response.body, null, 2));
        } else {
            console.error(error);
        }
    }

});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server live on ${PORT}`));