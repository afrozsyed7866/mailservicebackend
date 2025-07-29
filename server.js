require('dotenv').config(); // Load .env variables

   const express = require('express');
   const multer = require('multer');
   const XLSX = require('xlsx');
   const nodemailer = require('nodemailer');
   const cors = require('cors');
   const path = require('path');
   const fs = require('fs');

   const app = express();
   const port = process.env.PORT || 3001;

   // Create uploads directory if it doesn't exist
   const uploadDir = './Uploads';
   if (!fs.existsSync(uploadDir)) {
     fs.mkdirSync(uploadDir);
   }

   // Middleware
   app.use(cors({ origin: 'https://mailingservices-fe-qu5a.vercel.app' }));
   app.use(express.json());

   // Configure multer for file uploads
   const upload = multer({
     dest: 'Uploads/',
     fileFilter: (req, file, cb) => {
       const ext = path.extname(file.originalname).toLowerCase();
       if (ext !== '.xlsx' && ext !== '.xls') {
         return cb(new Error('Only Excel files are allowed'));
       }
       cb(null, true);
     },
   });

   // Configure nodemailer transporter
   const transporter = nodemailer.createTransport({
     host: 'smtp.gmail.com',
     port: 587,
     secure: false,
     auth: {
       user: process.env.EMAIL_USER,
       pass: process.env.EMAIL_PASS,
     },
     logger: true,
     debug: true,
   });

   // Professional HTML email template
   const generateEmailHTML = (job, recipientName) => {
     return `
       <!DOCTYPE html>
       <html>
       <head>
         <style>
           body { font-family: Arial, sans-serif; line-height: 1.6; color: #000000; }
           .container { max-width: 600px; margin: 0 auto; padding: 20px; }
           .header { background-color: #f4f4f4; padding: 10px; text-align: center; }
           .header img { max-width: 150px; }
           .content { padding: 20px; }
           .button { 
             display: inline-block; 
             padding: 10px 20px; 
             background-color: #007bff; 
             color: #ffffff; 
             text-decoration: none; 
             border-radius: 5px; 
             font-weight: bold; 
           }
           .footer { font-size: 12px; color: #777; text-align: center; margin-top: 20px; }
         </style>
       </head>
       <body>
         <div class="container">
           <div class="header">
             <h2>New Job Opportunity at ${job.company}</h2>
           </div>
           <div class="content">
             <p>Dear ${recipientName || 'Valued Candidate'},</p>
             <p>We are excited to share a new job opportunity with you!</p>
             <h3>${job.title}</h3>
             <p><strong>Company:</strong> ${job.company}</p>
             <p><strong>Location:</strong> ${job.location}</p>
             <p><strong>Salary:</strong> ${job.salary}</p>
             <p><strong>Application Deadline:</strong> ${job.applicationDeadline ? new Date(job.applicationDeadline).toLocaleDateString() : 'Not specified'}</p>
             <p><strong>Description:</strong></p>
             <p>${job.description.join('<br>')}</p>
             <p><strong>Requirements:</strong></p>
             <ul>
               ${job.requirements.map(req => `<li>${req}</li>`).join('')}
             </ul>
             <p><strong>Roles and Responsibilities:</strong></p>
             <ul>
               ${job.rolesAndResponsibilities.map(role => `<li>${role}</li>`).join('')}
             </ul>
             <p><strong>About ${job.company}:</strong></p>
             <p>${job.aboutCompany}</p>
             <p style="text-align: center;">
               <a href="https://www.careervalore.com/${job._id}" class="button">Apply Now</a>
             </p>
           </div>
           <div class="footer">
             <p>This is an automated email from Careervalore. Please do not reply directly to this email.</p>
             <p>&copy; ${new Date().getFullYear()} Your Job Posting Site. All rights reserved.</p>
           </div>
         </div>
       </body>
       </html>
     `;
   };

   // Endpoint to handle Excel file upload and send job notification emails
   app.post('/api/send-emails', upload.single('file'), async (req, res) => {
     try {
       console.log('Received request:', req.body, req.file);
       if (!req.file) {
         return res.status(400).json({ error: 'No file uploaded' });
       }

       // Parse job post JSON from request body
       let job;
       try {
         const jobData = req.body.job;
         if (!jobData) {
           return res.status(400).json({ error: 'Job post data is required in request body' });
         }
         job = JSON.parse(jobData);
       } catch (error) {
         return res.status(400).json({ error: `Invalid job post JSON in request body: ${error.message}` });
       }

       // Validate required job fields
       if (!job.title) {
         return res.status(400).json({ error: 'Missing required job fields in request body' });
       }

       // Read Excel file
       const workbook = XLSX.readFile(req.file.path);
       const sheetName = workbook.SheetNames[0];
       const worksheet = workbook.Sheets[sheetName];
       const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

       // Find email and name column indices
       const headers = data[0].map(h => h ? h.toString().toLowerCase().trim() : '');
       const emailColumnIndex = headers.findIndex(h =>
         ['email', 'e-mail', 'email address', 'mail'].includes(h)
       );
       const nameColumnIndex = headers.findIndex(h =>
         ['name', 'full name', 'recipient name', 'First Name'].includes(h)
       );

       // Validate email column exists
       if (emailColumnIndex === -1) {
         return res.status(400).json({ error: 'Excel file must contain a column with header "email", "e-mail", "email address", or "mail"' });
       }

       // Extract recipient data, skipping the header row
       const recipients = data.slice(1).map(row => ({
         email: row[emailColumnIndex] ? String(row[emailColumnIndex]).trim() : null,
         name: nameColumnIndex !== -1 && row[nameColumnIndex] ? String(row[nameColumnIndex]).trim() : null,
       }));

       // Validate recipients
       if (recipients.length === 0 || recipients.every(r => !r.email)) {
         return res.status(400).json({ error: 'No valid email addresses found in Excel file' });
       }

       // Send emails to all valid addresses
       const emailPromises = recipients.map(async (recipient) => {
         if (recipient.email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(recipient.email)) {
           try {
             await transporter.sendMail({
               from: process.env.EMAIL_USER,
               to: recipient.email,
               subject: `New Job Opportunity: ${job.title} at ${job.company}`,
               html: generateEmailHTML(job, recipient.name),
             });
             return { email: recipient.email, status: 'success' };
           } catch (error) {
             return { email: recipient.email, status: 'failed', error: error.message };
           }
         }
         return { email: recipient.email || 'unknown', status: 'failed', error: 'Invalid or missing email' };
       });

       const results = await Promise.all(emailPromises);

       // Clean up uploaded file
       fs.unlinkSync(req.file.path);

       res.json({
         message: 'Email sending process completed',
         results: results,
       });

     } catch (error) {
       console.error('Error in /api/send-emails:', error);
       if (req.file && fs.existsSync(req.file.path)) {
         fs.unlinkSync(req.file.path);
       }
       res.status(500).json({ error: 'Server error: ' + error.message });
     }
   });

   // Start server
   app.listen(port, '0.0.0.0', () => {
     console.log(`Server running on port ${port}`);
   });