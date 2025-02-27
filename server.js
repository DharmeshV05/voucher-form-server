require("dotenv").config();
const express = require("express");
const bodyParser = require("body-parser");
const { google } = require("googleapis");
const { GoogleAuth } = require("google-auth-library");
const multer = require("multer");
const PDFDocument = require("pdfkit");
const fs = require("fs");
const path = require("path");
const cors = require("cors");
const axios = require("axios");
const nodemailer = require("nodemailer"); // Added for email notifications
const app = express();
const PORT = process.env.PORT || 3000;

app.use(
  cors({
    origin: "*",
    credentials: true,
    optionsSuccessStatus: 200,
  })
);

app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, "public")));

const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive",
];

const auth = new GoogleAuth({
  keyFile: "credentials.json",
  scopes: SCOPES,
});

const sheets = google.sheets({ version: "v4", auth });
const drive = google.drive({ version: "v3", auth });

const storage = multer.memoryStorage();
const upload = multer({ storage });

const filterToSpreadsheetMap = {
  Contentstack: process.env.SPREADSHEET_ID_CONTENTSTACK,
  Surfboard: process.env.SPREADSHEET_ID_SURFBOARD,
  RawEngineering: process.env.SPREADSHEET_ID_RAWENGINEERING,
};

const driveFolderId = process.env.DRIVE_FOLDER_ID;

const headerValues = [
  "Voucher No.",
  "Date",
  "Filter",
  "Pay to",
  "Account Head",
  "Towards",
  "The Sum",
  "Amount Rs.",
  "Checked By",
  "Approved By",
  "Receiver Signature",
  "PDF Link",
];

// Email transporter setup
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS,
  },
});

// Centralized voucher number storage (in-memory for simplicity, ideally use a database)
const lastVoucherNumbers = {
  Contentstack: 0,
  Surfboard: 0,
  RawEngineering: 0,
};

setInterval(() => {
  axios
    .get(`http://localhost:${PORT}/ping`)
    .then((response) => {
      console.log("Pinged server to keep it warm.");
    })
    .catch((error) => {
      console.error("Error pinging the server:", error.message);
    });
}, 30000);

// Ping endpoint
app.get("/ping", (req, res) => {
  res.status(200).send({ message: "Server is active" });
});

// Centralized voucher number generation
app.get("/get-voucher-no", async (req, res) => {
  const filter = req.query.filter;
  if (!filter || !filterToSpreadsheetMap[filter]) {
    return res.status(400).send({ error: "Invalid filter option" });
  }
  try {
    const spreadsheetId = filterToSpreadsheetMap[filter];
    const range = `${filter}!A:A`; // Assuming voucher numbers are in column A
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });
    const rows = response.data.values || [];
    const lastNumber = rows.length > 1 ? parseInt(rows[rows.length - 1][0].split("-")[2]) || 0 : 0;
    const prefix = `${filter.slice(0, 2).toUpperCase()}-${new Date().getFullYear()}-`;
    const newNumber = String(lastNumber + 1).padStart(3, "0");
    res.send({ voucherNo: `${prefix}${newNumber}` });
  } catch (error) {
    console.error("Error fetching voucher number:", error);
    // Fallback to in-memory if Sheets fails
    lastVoucherNumbers[filter]++;
    const prefix = `${filter.slice(0, 2).toUpperCase()}-${new Date().getFullYear()}-`;
    res.send({ voucherNo: `${prefix}${String(lastVoucherNumbers[filter]).padStart(3, "0")}` });
  }
});

// Auto-fill suggestions endpoint
app.get("/get-suggestions", async (req, res) => {
  const filter = req.query.filter;
  if (!filter || !filterToSpreadsheetMap[filter]) {
    return res.status(400).send({ error: "Invalid filter option" });
  }
  try {
    const spreadsheetId = filterToSpreadsheetMap[filter];
    const range = `${filter}!D:D`; // "Pay to" is in column D
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range,
    });
    const payToValues = response.data.values ? response.data.values.flat().filter(Boolean) : [];
    const uniquePayToSuggestions = [...new Set(payToValues.slice(1))]; // Remove header and duplicates
    res.send({ payToSuggestions: uniquePayToSuggestions });
  } catch (error) {
    console.error("Error fetching suggestions:", error);
    res.status(500).send({ error: "Failed to fetch suggestions" });
  }
});

app.post("/submit", upload.none(), async (req, res) => {
  try {
    const voucherData = req.body;
    console.log("Received voucherData:", voucherData);
    const filterOption = voucherData.filter;
    const spreadsheetId = filterToSpreadsheetMap[filterOption];

    if (!spreadsheetId) {
      return res.status(400).send({ error: "Invalid filter option" });
    }

    voucherData.voucherNo = voucherData.voucherNo; // Already set by frontend from /get-voucher-no

    const sheetTitle = filterOption;
    const sheetURL = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;

    const getSpreadsheetResponse = await sheets.spreadsheets.get({
      spreadsheetId,
    });

    const sheetsList = getSpreadsheetResponse.data.sheets;
    const sheetExists = sheetsList.some(
      (sheet) => sheet.properties.title === sheetTitle
    );

    if (!sheetExists) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: {
          requests: [
            {
              addSheet: {
                properties: {
                  title: sheetTitle,
                  gridProperties: {
                    rowCount: 1000,
                    columnCount: 14,
                  },
                },
              },
            },
          ],
        },
      });

      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetTitle}!A1:O1`,
        valueInputOption: "RAW",
        requestBody: {
          values: [headerValues],
        },
      });
    }

    const pdfFileName = `${filterOption}_${voucherData.voucherNo}.pdf`;
    const pdfFilePath = path.join(__dirname, pdfFileName);
    const doc = new PDFDocument({ margin: 30 });
    const pdfStream = fs.createWriteStream(pdfFilePath);
    doc.pipe(pdfStream);

    const underlineYPosition = 35;

    doc.fontSize(12).text("Date:", 400, 20);
    doc.fontSize(12).text(voucherData.date, 440, 20);
    doc.moveTo(440, underlineYPosition).lineTo(550, underlineYPosition).stroke();

    doc.fontSize(12).text("Voucher No:", 400, 40);
    doc.fontSize(12).text(voucherData.voucherNo, 470, 40);
    doc.moveTo(440, underlineYPosition + 20).lineTo(550, underlineYPosition + 20).stroke();

    const filterLogoMap = {
      Contentstack: path.join(__dirname, "public", "contentstack.png"),
      Surfboard: path.join(__dirname, "public", "surfboard.png"),
      RawEngineering: path.join(__dirname, "public", "raw.png"),
    };
    const filterLogo = filterLogoMap[voucherData.filter];
    if (fs.existsSync(filterLogo)) {
      doc.image(filterLogo, 30, 30, { width: 100 });
    }

    doc.moveDown(3);

    const drawLineAndText = (label, value, yPosition) => {
      doc.fontSize(12).text(label, 30, yPosition);
      doc.moveTo(120, yPosition + 12).lineTo(550, yPosition + 12).stroke();
      doc.fontSize(12).text(value, 130, yPosition);
    };

    drawLineAndText("Pay to:", voucherData.payTo, 160);
    drawLineAndText("Account Head:", voucherData.accountHead, 200);
    drawLineAndText("Towards:", voucherData.account, 240);

    doc.fontSize(12).text("Amount Rs.", 30, 280);
    doc.moveTo(120, 292).lineTo(550, 292).stroke();
    doc.fontSize(12).text(voucherData.amount, 130, 280);

    doc.fontSize(12).text("The Sum.", 30, 320);
    doc.moveTo(120, 332).lineTo(550, 332).stroke();
    doc.fontSize(12).text(voucherData.amountRs, 130, 320);

    const amountSectionY = 320;
    const gap = 65;
    const signatureSectionY = amountSectionY + gap;

    const drawSignatureLine = (label, xPosition, yPosition) => {
      doc.moveTo(xPosition, yPosition).lineTo(xPosition + 100, yPosition).stroke();
      doc.fontSize(12).text(label, xPosition, yPosition + 5);
    };

    drawSignatureLine("Checked By", 50, signatureSectionY);
    drawSignatureLine("Approved By", 250, signatureSectionY);
    drawSignatureLine("Receiver Signature", 450, signatureSectionY);

    // Add digital signature if provided
    if (voucherData.receiverSignature) {
      const signatureBuffer = Buffer.from(
        voucherData.receiverSignature.split(",")[1],
        "base64"
      );
      doc.image(signatureBuffer, 450, signatureSectionY - 20, { width: 100 });
    }

    doc.end();

    pdfStream.on("finish", async () => {
      try {
        const pdfFileMetadata = {
          name: pdfFileName,
          parents: [driveFolderId],
        };
        const pdfMedia = {
          mimeType: "application/pdf",
          body: fs.createReadStream(pdfFilePath),
        };
        const pdfUploadResponse = await drive.files.create({
          resource: pdfFileMetadata,
          media: pdfMedia,
          fields: "id, webViewLink",
        });

        const pdfFileId = pdfUploadResponse.data.id;
        const pdfLink = pdfUploadResponse.data.webViewLink;

        const values = [
          [
            voucherData.voucherNo,
            voucherData.date,
            voucherData.filter,
            voucherData.payTo,
            voucherData.accountHead,
            voucherData.account,
            voucherData.amount,
            voucherData.amountRs,
            voucherData.checkedBy,
            voucherData.approvedBy,
            voucherData.receiverSignature ? "Signed" : "",
            pdfLink,
          ],
        ];

        await sheets.spreadsheets.values.append({
          spreadsheetId,
          range: `${sheetTitle}!A:O`,
          valueInputOption: "RAW",
          requestBody: {
            values,
          },
        });

        // Send approval notification
        await transporter.sendMail({
          from: process.env.EMAIL_USER,
          to: process.env.APPROVER_EMAIL || "approver@example.com",
          subject: `Voucher ${voucherData.voucherNo} Submitted for Approval`,
          text: `A new voucher has been submitted.\nSheet: ${sheetURL}\nPDF: ${pdfLink}`,
        });
        console.log("Approval email sent");

        fs.unlinkSync(pdfFilePath);

        res.status(200).send({
          message: "Data submitted successfully and PDF uploaded!",
          sheetURL: sheetURL,
          pdfFileId: pdfFileId,
        });
      } catch (error) {
        console.error("Error uploading PDF or sending email:", error);
        res.status(500).send({ error: "Failed to upload PDF or notify approver" });
      }
    });

    pdfStream.on("error", (error) => {
      console.error("Error creating PDF:", error);
      res.status(500).send({ error: "Failed to create PDF" });
    });
  } catch (error) {
    console.error("Error submitting data:", error);
    res.status(500).send({ error: "Failed to submit data" });
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});