const express = require('express');
const nodemailer = require('nodemailer');
const XLSX = require('xlsx');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

// Optimized transporter with connection pooling
const transporter = nodemailer.createTransport({
    service: 'gmail',
    pool: true, 
    maxConnections: 5,
    auth: { 
        user: process.env.EMAIL_USER, 
        pass: process.env.EMAIL_PASS 
    }
});

app.post('/send-manifest', (req, res) => {
    // Handle Silent Wake-Up Ping
    if (req.body.ping) {
        console.log("Server Woken Up by login activity.");
        return res.status(200).json({ success: true, message: "Awake" });
    }

    const { operator, courier, awbs, count } = req.body;

    // Send FAST response to mobile (don't wait for email)
    res.status(200).json({ success: true });

    // Background Processing: Excel Generation
    try {
        const data = [["SL No.", "AWB NUMBER"]];
        awbs.forEach((a, i) => data.push([i + 1, a]));
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, "Manifest");
        const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

        const awbListText = awbs.map((a, i) => `${i + 1}. ${a}`).join('\n');

        // Send Email
        transporter.sendMail({
            from: `"Medika Logistics" <${process.env.EMAIL_USER}>`,
            to: "Rajeshrak413@gmail.com",
            subject: `Manifest: ${courier} (${count} Parcels)`,
            text: `Operator: ${operator}\nTotal Count: ${count}\n\nAWB List:\n${awbListText}`,
            attachments: [{ filename: `${courier}_Manifest.xlsx`, content: excelBuffer }]
        });
    } catch (err) {
        console.error("Background Error:", err);
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Fast-Server active on port ${PORT}`));
