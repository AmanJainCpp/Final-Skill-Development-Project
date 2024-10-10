const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const path = require('path');
const dotenv = require('dotenv');
const session = require('express-session');
const bcrypt = require('bcryptjs');

// Load environment variables
dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

// Set up multer for file uploads
const upload = multer({ dest: 'uploads/' });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname))); // Serve static files

// Set up session management
app.use(
    session({
        secret: 'secret-key', // Change this in production
        resave: false,
        saveUninitialized: true,
        cookie: { maxAge: 60000 } // Session expiration time (e.g., 1 minute here)
    })
);

// Sample user for login (in-memory, replace with DB later)
const users = [
    {
        id: 1,
        username: 'admin',
        passwordHash: bcrypt.hashSync('password123', 10) // Hashed password
    }
];

// Serve the login page
app.get('/login', (req, res) => {
    if (req.session.userId) {
        return res.redirect('/dashboard');
    }
    res.sendFile(path.join(__dirname, 'login.html'));
});

// Handle login POST request
app.post('/login', (req, res) => {
    const { username, password } = req.body;
    const user = users.find(u => u.username === username);

    if (!user) {
        return res.send('<p>Invalid username or password</p>');
    }

    const passwordMatch = bcrypt.compareSync(password, user.passwordHash);
    if (!passwordMatch) {
        return res.send('<p>Invalid username or password</p>');
    }

    // Set session and redirect to dashboard
    req.session.userId = user.id;
    res.redirect('/dashboard');
});

// Protected dashboard route
app.get('/dashboard', (req, res) => {
    if (!req.session.userId) {
        return res.redirect('/login');
    }
    res.sendFile(path.join(__dirname, 'views/dashboard.html')); // Serve dashboard
});

// Handle logout
app.get('/logout', (req, res) => {
    req.session.destroy((err) => {
        if (err) {
            return res.send('Error logging out.');
        }
        res.redirect('/login');
    });
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
            user: process.env.EMAIL,
            pass: process.env.PASSWORD
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
