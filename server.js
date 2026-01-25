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
    const todayDate = new Date().toLocaleDateString('en-IN'); 

    // Instant response to Mobile Portal
    res.status(200).json({ success: true });

    try {
        // 1. Generate Excel with Date and Courier Columns
        const data = [["SL No.", "Date", "Courier Name", "AWB NUMBER"]];
        awbs.forEach((a, i) => data.push([i + 1, todayDate, courier, a]));
        
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(data);
        XLSX.utils.book_append_sheet(wb, ws, "Manifest");
        const excelBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

        // 2. Prepare the Email Message
        const msg = {
            // Main Recipients
            to: ['rajeshrak413@outlook.com','gokulkrishnan.velayutham@medikabazaar.com','bhaskar.r@medikabazaar.com','hanumanta.madival@medikabazaar.com'], 
            // CC Recipients (Add more here)
            cc: ['elumalai.b@medikabazaar.com'], 
            from: {
                email: 'Rajeshrak413@gmail.com', // Verified Sender
                name: 'Medika Logistics Portal'
            },
            subject: `Outbound Manifest - ${courier} - ${todayDate}`,
            text: `Hello,

Please find the outbound manifest details below:

Date :- ${todayDate}
Courier Name  :- ${courier}
Operator Name :- ${operator}
Total
count of AWB  :- ${count}

Please find the attached Excel file for the complete AWB list.

Regards,
Outbound`,
            attachments: [{
                content: excelBuffer.toString('base64'),
                filename: `${courier}_Manifest_${todayDate.replace(/\//g, '-')}.xlsx`,
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                disposition: 'attachment'
            }]
        };

        // 3. Send Email
        // Note: Use .send() instead of .sendMultiple() when using CC/BCC fields
        await sgMail.send(msg); 
        console.log(`✅ Success: Manifest for ${courier} sent with CC.`);
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
