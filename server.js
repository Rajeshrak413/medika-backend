const express = require('express');
const nodemailer = require('nodemailer');
const XLSX = require('xlsx');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

app.post('/send-manifest', async (req, res) => {
    const { operator, courier, awbs, count } = req.body;
    try {
        // 1. Generate Excel Buffer
        const data = [["SL No.", "AWB NUMBER"]];
        awbs.forEach((a, i) => data.push([i + 1, a]));
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, "Manifest");
        const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

        // 2. Email Setup
        const transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: { 
                user: 'YOUR_EMAIL@gmail.com', 
                pass: 'YOUR_GOOGLE_APP_PASSWORD' 
            }
        });

        // Format AWB list one-by-one for email body
        const awbListText = awbs.map((a, i) => `${i + 1}. ${a}`).join('\n');

        await transporter.sendMail({
            from: '"Medika Logistics" <YOUR_EMAIL@gmail.com>',
            to: "MANAGER_EMAIL@gmail.com", // Change to your recipient
            subject: `Manifest: ${courier} (${count} Parcels)`,
            text: `Operator: ${operator}\nTotal: ${count}\n\nAWB List:\n${awbListText}`,
            attachments: [{ filename: `${courier}_Manifest.xlsx`, content: excelBuffer }]
        });

        res.status(200).json({ success: true });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Backend running on port ${PORT}`));
