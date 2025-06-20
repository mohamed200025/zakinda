const express = require('express');
const session = require('express-session'); // ❌ 
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const pdfParse = require('pdf-parse');
const ExcelJS = require('exceljs');
const mammoth = require("mammoth");

const app = express();
app.use(session({
    secret: 'your-secret-password', // change to something strong
    resave: false,
    saveUninitialized: true
}));
app.use(express.urlencoded({ extended: true })); // needed to read form data

const PORT = 3000;

// Serve static files (HTML, CSS, JS) from "public" folder
app.use(express.static('public'));

// Configure Multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/');
    },
    filename: (req, file, cb) => {
        cb(null, Date.now() + path.extname(file.originalname)); // e.g., 123456.pdf
    }
});
const upload = multer({ storage });

//
app.get('/', (req, res) => {
    if (req.session.loggedIn) {
        res.sendFile(path.join(__dirname, 'index.html'));
    } else {
        res.sendFile(path.join(__dirname, 'public/login.html'));
    }
});

app.post('/login', (req, res) => {
    const password = req.body.password;
    if (password === 'zakinda123') { // <== change to your real password
        req.session.loggedIn = true;
        res.redirect('/');
    } else {
        res.send('Incorrect password. <a href="/">Try again</a>');
    }
});




// Upload endpoint

function requireLogin(req, res, next) {
    if (req.session.loggedIn) {
        next();
    } else {
        res.redirect('/');
    }
}

app.post('/upload', requireLogin, upload.single('file'), async (req, res) => {

    if (!req.file) {
    return res.status(400).send('No file uploaded.');
}
    const filePath = req.file.path;
    const ext = path.extname(req.file.originalname).toLowerCase();
    let text = '';

    try {
        // 📄 1. Read file based on extension
        if (ext === '.pdf') {
            const dataBuffer = fs.readFileSync(filePath);
            const data = await pdfParse(dataBuffer);
            text = data.text;
        } else if (ext === '.docx') {
            const result = await mammoth.extractRawText({ path: filePath });
            text = result.value;
        } else if (ext === '.pptx') {
         const { extractText } = require("pptx2json");
         const result = await extractText(filePath);
         text = result.map(slide => slide.text).join('\n');
        } else {
            return res.status(400).send('Unsupported file type.');
        }

        // 📤 2. Extract lines
        const lines = text.split('\n');
        const extractedData = [];

        for (let line of lines) {
            // Phone first
            const phoneMatch = line.match(/\d{8,}/);
            const phone = phoneMatch ? phoneMatch[0] : '';

            let cleanLine = line;
            if (phone) cleanLine = cleanLine.replace(phone, '');

            // Then email
            const emailMatch = cleanLine.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/);
            const email = emailMatch ? emailMatch[0] : '';
            if (email) cleanLine = cleanLine.replace(email, '');

            // Names
            const nameParts = cleanLine.trim().split(/\s+/);
            const first = nameParts[0] || '';
            const last = nameParts[1] || '';

            if (first || last || email || phone) {
                extractedData.push({ first, last, email, phone });
            }
        }

        // 📥 3. Generate Excel
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Contacts');
        worksheet.columns = [
            { header: 'First Name', key: 'first' },
            { header: 'Last Name', key: 'last' },
            { header: 'Email', key: 'email' },
            { header: 'Phone', key: 'phone' },
        ];
        worksheet.addRows(extractedData);

        const excelPath = `uploads/${Date.now()}_contacts.xlsx`;
        await workbook.xlsx.writeFile(excelPath);

        res.download(excelPath, 'contacts.xlsx', (err) => {
            if (err) console.error(err);
            fs.unlinkSync(filePath);
            fs.unlinkSync(excelPath);
        });

    } catch (err) {
        console.error(err);
        res.status(500).send('Error processing file.');
    }
});



app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
});
