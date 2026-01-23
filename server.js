const express = require('express');
const sgMail = require('@sendgrid/mail');
const XLSX = require('xlsx');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

// Initialize SendGrid with your new API Key
sgMail.setApiKey(process.env.SENDGRID_API_KEY);

app.post('/send-manifest', async (req, res) => {
    // Handle silent wake-up ping
    if (req.body.ping) return res.status(200).json({ success: true });

    const { operator, courier, awbs, count } = req.body;
    
    // Respond immediately to mobile to keep scanning fast
    res.status(200).json({ success: true });

    try {
        // 1. Generate Excel Buffer
        const data = [["SL No.", "AWB NUMBER"]];
        awbs.forEach((a, i) => data.push([i + 1, a]));
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, "Manifest");
        const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

        // 2. Setup Email using SendGrid API
        const msg = {
            to: 'Rajeshrak413@gmail.com',
            from: process.env.FROM_EMAIL, 
            subject: `Manifest: ${courier} (${count} Parcels)`,
            text: `Operator: ${operator}\nTotal parcels: ${count}\n\nPlease find the Excel manifest attached.`,
            attachments: [
                {
                    content: excelBuffer.toString('base64'),
                    filename: `${courier}_Manifest.xlsx`,
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    disposition: 'attachment'
                }
            ]
        };

        // 3. Fire the email
        await sgMail.send(msg);
        console.log(`Success: Manifest for ${courier} emailed.`);
    } catch (error) {
        console.error("SendGrid error details:", error.response ? error.response.body : error);
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
