const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const path = require('path');
const dotenv = require('dotenv');
//
require('dotenv').config();

dotenv.config();  // Load environment variables from .env

const app = express();
const PORT = process.env.PORT || 3000;

// Set up multer for file uploads
const upload = multer({ dest: 'uploads/' });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Home route to serve the upload form
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Upload route to handle Excel file
app.post('/upload', upload.single('attendanceFile'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('No file uploaded.');
    }
    const filePath = path.join(__dirname, req.file.path);
    handleExcel(filePath)
        .then(() => res.send('Emails sent successfully!'))
        .catch((err) => res.status(500).send('Error: ' + err));
});

// Function to read Excel file and find students with attendance < 60%
const handleExcel = async (filePath) => {
    const workbook = xlsx.readFile(filePath);
    const sheet_name_list = workbook.SheetNames;
    const sheet = workbook.Sheets[sheet_name_list[0]];  // Assuming the first sheet
    const data = xlsx.utils.sheet_to_json(sheet);

    // Filter students with attendance less than 60%
    const studentsWithLowAttendance = data.filter(student => student['Total Percentage'] < 60);

    if (studentsWithLowAttendance.length > 0) {
        await sendEmails(studentsWithLowAttendance);
    } else {
        console.log('No students with low attendance.');
    }
};

// Function to send emails to parents
const sendEmails = async (students) => {
    // Set up email transporter
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: process.env.EMAIL,       // From .env file
            pass: process.env.PASSWORD     // From .env file
        }
    });

    for (const student of students) {
        const mailOptions = {
            from: process.env.EMAIL,
            to: student['Parent Email'],   // Ensure this column exists in your Excel file
            subject: 'Low Attendance Warning',
            text: `Dear Parent,

Your child ${student['Student Name']} (Enrollment: ${student['Enrollment Number']}) has a low attendance of ${student['Total Percentage']}%. Please ensure your child attends more classes.

Regards,
School Administration`
        };

        // Send email
        try {
            await transporter.sendMail(mailOptions);
            console.log(`Email sent to ${student['Parent Email']}`);
        } catch (error) {
            console.error(`Failed to send email to ${student['Parent Email']}:`, error);
        }
    }
};

// Start server
app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
