require('dotenv').config();
const express = require('express');
const multer = require('multer');
const nodemailer = require('nodemailer');
const XLSX = require('xlsx');
const cors = require('cors');

const app = express();
const port = process.env.PORT || 3000;

// ────────────────────────────────────────────────
// Middleware
// ────────────────────────────────────────────────
app.use(cors()); // Allow your frontend origin (or * for testing)
app.use(express.json()); // For any JSON payloads if needed later

// Multer: store files in memory (good for Render – no disk)
const upload = multer({ storage: multer.memoryStorage() });

// ────────────────────────────────────────────────
// Nodemailer Transporter
// ────────────────────────────────────────────────
const transporter = nodemailer.createTransport({
  service: 'gmail', // change to 'hotmail', 'yahoo', or custom SMTP
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS,
  },
  // If using custom SMTP (recommended for production):
  // host: process.env.SMTP_HOST,
  // port: process.env.SMTP_PORT,
  // secure: true/false,
  // auth: { user: ..., pass: ... }
});

// Verify transporter on startup (good practice)
transporter.verify((error) => {
  if (error) {
    console.error('Transporter verification failed:', error);
  } else {
    console.log('Email transporter is ready');
  }
});

// ────────────────────────────────────────────────
// Endpoint: /send-manifest
// Expects multipart/form-data with:
// - operator, courier, count, awbs (JSON string), pdf (file)
// ────────────────────────────────────────────────
app.post('/send-manifest', upload.single('pdf'), async (req, res) => {
  try {
    const { operator, courier, count, awbs } = req.body;
    const pdfBuffer = req.file ? req.file.buffer : null;

    if (!operator || !courier || !count || !awbs) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    let awbsArray;
    try {
      awbsArray = JSON.parse(awbs);
    } catch (e) {
      return res.status(400).json({ error: 'Invalid awbs JSON' });
    }

    // ─── Generate Excel in memory ────────────────────────────────
    const worksheet = XLSX.utils.json_to_sheet(
      awbsArray.map((awb, i) => ({ SL: i + 1, 'AWB Number': awb }))
    );
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Manifest');

    const excelBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    // ─── Email setup ─────────────────────────────────────────────
    const mailOptions = {
      from: `"Medika Logistics" <${process.env.EMAIL_USER}>`,
      to: process.env.RECIPIENT_EMAIL || 'rajeshrak413@outlook.com',
      subject: `Manifest - ${courier} - \( {count} parcels ( \){operator})`,
      text: `New manifest received.\n\nOperator: ${operator}\nCourier: ${courier}\nTotal parcels: ${count}\nDate: ${new Date().toLocaleString('en-IN')}`,
      attachments: [
        {
          filename: `Manifest_\( {courier}_ \){new Date().toISOString().split('T')[0]}.xlsx`,
          content: excelBuffer,
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        },
      ],
    };

    // Attach PDF if received
    if (pdfBuffer) {
      mailOptions.attachments.push({
        filename: req.file.originalname || `Manifest_${courier}.pdf`,
        content: pdfBuffer,
        contentType: 'application/pdf',
      });
    }

    // Send email
    const info = await transporter.sendMail(mailOptions);

    console.log('Email sent:', info.messageId);
    res.status(200).json({ success: true, message: 'Manifest emailed successfully' });
  } catch (error) {
    console.error('Error sending email:', error);
    res.status(500).json({ error: 'Failed to send email', details: error.message });
  }
});

// Health check
app.get('/', (req, res) => {
  res.send('Medika Backend is running 🚀');
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});