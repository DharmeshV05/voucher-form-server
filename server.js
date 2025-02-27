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
];

const lastVoucherNumbers = {
  Contentstack: 0,
  Surfboard: 0,
  RawEngineering: 0,
};

setInterval(() => {
  axios.get(`http://localhost:${PORT}/ping`)
    .then(response => {
      console.log("Pinged server to keep it warm.");
    })
    .catch(error => {
      console.error("Error pinging the server:", error.message);
    });
}, 30000); 

// Ping endpoint
app.get('/ping', (req, res) => {
  res.status(200).send({ message: 'Server is active' });
});

app.get("/get-voucher-no", (req, res) => {
  const filter = req.query.filter;
  if (!filter || !filterToSpreadsheetMap[filter]) {
    return res.status(400).send({ error: "Invalid filter option" });
  }
  res.send({ voucherNo: lastVoucherNumbers[filter] + 1 });
});

app.post("/submit", upload.none(), async (req, res) => {
  try {
    const voucherData = req.body;
    console.log("Received voucherData:", voucherData); // Debug log to verify incoming data
    const filterOption = voucherData.filter;
    const spreadsheetId = filterToSpreadsheetMap[filterOption];

    if (!spreadsheetId) {
      return res.status(400).send({ error: "Invalid filter option" });
    }

    lastVoucherNumbers[filterOption]++;
    const voucherNo = lastVoucherNumbers[filterOption];
    voucherData.voucherNo = voucherNo;

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
          values: [headerValues.concat("PDF Link")],
        },
      });
    }

    const pdfFileName = `${filterOption}_${voucherNo}.pdf`;
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

    // Define filterLogoMap with resolved file paths
    const filterLogoMap = {
      Contentstack: path.join(__dirname, "public", "contentstack.png"), // Use absolute path
      Surfboard: path.join(__dirname, "public", "surfboard.png"),
      RawEngineering: path.join(__dirname, "public", "raw.png"),
    };
    const filterLogo = filterLogoMap[voucherData.filter];
    console.log("Resolved logo path:", filterLogo); // Debug log to verify path
    if (filterLogo) {
      try {
        // Verify file exists before adding to PDF
        if (fs.existsSync(filterLogo)) {
          doc.image(filterLogo, 30, 30, { width: 100 });
          console.log(`Successfully added logo: ${filterLogo}`);
        } else {
          console.error(`Logo file not found: ${filterLogo}`);
        }
      } catch (error) {
        console.error(`Error loading logo ${filterLogo}:`, error.message);
      }
    }

    doc.moveDown(3);

    const drawLineAndText = (label, value, yPosition) => {
      doc.fontSize(12).text(label, 30, yPosition);
      doc
        .moveTo(120, yPosition + 12)
        .lineTo(550, yPosition + 12)
        .stroke();
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
      doc
        .moveTo(xPosition, yPosition)
        .lineTo(xPosition + 100, yPosition)
        .stroke();
      doc.fontSize(12).text(label, xPosition, yPosition + 5);
    };

    drawSignatureLine("Checked By", 50, signatureSectionY);
    drawSignatureLine("Approved By", 250, signatureSectionY);
    drawSignatureLine("Receiver Signature", 450, signatureSectionY);

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
            voucherData.receiverSignature,
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

        fs.unlinkSync(pdfFilePath);

        res.status(200).send({
          message: "Data submitted successfully and PDF uploaded!",
          sheetURL: sheetURL,
          pdfFileId: pdfFileId,
        });
      } catch (error) {
        console.error("Error uploading PDF:", error);
        res.status(500).send({ error: "Failed to upload PDF" });
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