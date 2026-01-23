const express = require('express');
const sgMail = require('@sendgrid/mail');
const XLSX = require('xlsx');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

sgMail.setApiKey(process.env.SENDGRID_API_KEY);

app.post('/send-manifest', async (req, res) => {
    if (req.body.ping) return res.status(200).json({ success: true });

    const { operator, courier, awbs, count } = req.body;
    res.status(200).json({ success: true });

    try {
        const data = [["SL No.", "AWB NUMBER"]];
        awbs.forEach((a, i) => data.push([i + 1, a]));
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, "Manifest");
        const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

        // THE FIX: "From" is now structured as an object
        const msg = {
            to: 'Rajeshrak413@gmail.com',
            from: {
                email: 'Rajeshrak413@gmail.com', // Must match your verified sender
                name: 'Medika Logistics Portal'
            },
            subject: `Manifest: ${courier} (${count} Parcels)`,
            text: `Operator: ${operator}\nTotal: ${count}\n\nPlease see attached Excel.`,
            attachments: [{
                content: excelBuffer.toString('base64'),
                filename: `${courier}_Manifest.xlsx`,
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                disposition: 'attachment'
            }]
        };

        await sgMail.send(msg);
        console.log("✅ Email sent successfully!");
    } catch (error) {
        console.error("❌ SendGrid Error Details:");
        if (error.response) {
            console.error(JSON.stringify(error.response.body, null, 2));
        } else {
            console.error(error);
        }
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server live on ${PORT}`));
