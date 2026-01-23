// Ensure this part looks exactly like this in your server.js
const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS
    }
});

// Create the vertical list for the email body
const awbListText = awbs.map((a, i) => `${i + 1}. ${a}`).join('\n');

await transporter.sendMail({
    from: `"Medika Logistics" <${process.env.EMAIL_USER}>`,
    to: "Rajeshrak413@gmail.com", 
    subject: `Manifest: ${courier} (${count} Parcels)`,
    text: `Operator: ${operator}\nTotal: ${count}\n\nAWB List:\n${awbListText}`,
    attachments: [{ filename: `${courier}_Manifest.xlsx`, content: excelBuffer }]
});
