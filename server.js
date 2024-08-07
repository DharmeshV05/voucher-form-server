require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const { google } = require('googleapis');
const { GoogleAuth } = require('google-auth-library');
const multer = require('multer');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;


app.use(cors({
  origin: '*', // Replace with your frontend URL
  credentials: true,
  optionsSuccessStatus: 200
}));
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive',
];

const auth = new GoogleAuth({
  keyFile: "credentials.json",
  scopes: SCOPES,
});
const sheets = google.sheets({ version: 'v4', auth });
const drive = google.drive({ version: 'v3', auth });

const storage = multer.memoryStorage();
const upload = multer({ storage });

const filterToSpreadsheetMap = {
  Contentstack: process.env.SPREADSHEET_ID_CONTENTSTACK,
  Surfboard: process.env.SPREADSHEET_ID_SURFBOARD,
  RawEngineering: process.env.SPREADSHEET_ID_RAWENGINEERING,
};

const driveFolderId = process.env.DRIVE_FOLDER_ID;

const headerValues = [
  'Voucher No.',
  'Date',
  'Filter',
  'Pay to',
  'Account Head',
  'Paid by',
  'Towards',
  'The Sum',
  'Amount Rs.',
  'Prepared By',
  'Checked By',
  'Approved By',
  'Receiver Signature',
];

const lastVoucherNumbers = {
  Contentstack: 0,
  Surfboard: 0,
  RawEngineering: 0,
};

// New endpoint to get the current voucher number for a given filter
app.get('/get-voucher-no', (req, res) => {
  const filter = req.query.filter;
  if (!filter || !filterToSpreadsheetMap[filter]) {
    return res.status(400).send({ error: 'Invalid filter option' });
  }
  lastVoucherNumbers[filter]++;
  res.send({ voucherNo: lastVoucherNumbers[filter] });
});

app.post('/submit', upload.none(), async (req, res) => {
  try {
    const voucherData = req.body;
    const filterOption = voucherData.filter;
    const spreadsheetId = filterToSpreadsheetMap[filterOption];

    if (!spreadsheetId) {
      return res.status(400).send({ error: 'Invalid filter option' });
    }

    const voucherNo = parseInt(voucherData.voucherNo, 10);
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
        range: `${sheetTitle}!A1:N1`,
        valueInputOption: 'RAW',
        requestBody: {
          values: [headerValues],
        },
      });
    }

    const values = [
      [
        voucherData.voucherNo,
        voucherData.date,
        voucherData.filter,
        voucherData.payTo,
        voucherData.accountHead,
        voucherData.paidBy,
        voucherData.account,
        voucherData.amount,
        voucherData.amountRs,
        voucherData.preparedBy,
        voucherData.checkedBy,
        voucherData.approvedBy,
        voucherData.receiverSignature,
      ],
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetTitle}!A:N`,
      valueInputOption: 'RAW',
      requestBody: {
        values,
      },
    });

    const pdfFileName = `${filterOption}_${voucherNo}.pdf`;
    const pdfFilePath = path.join(__dirname, pdfFileName);
    const doc = new PDFDocument({ margin: 30 });
    const pdfStream = fs.createWriteStream(pdfFilePath);
    doc.pipe(pdfStream);

    doc.fontSize(12).text('Date', 450, 20);
    doc.fontSize(12).text(voucherData.date, 450, 40);
    doc.fontSize(12).text('Voucher No.', 450, 60);
    doc.fontSize(12).text(voucherData.voucherNo, 450, 80);

    const filterLogoMap = {
      Contentstack: 'public/contentstack.png',
      Surfboard: 'public/surfboard.png',
      RawEngineering: 'public/raw.png',
    };
    const filterLogo = filterLogoMap[voucherData.filter];
    if (filterLogo) {
      doc.image(filterLogo, 30, 30, { width: 100 });
    }

    doc.moveDown(3);

    const drawLineAndText = (label, value, yPosition) => {
      doc.fontSize(12).text(label, 30, yPosition);
      doc.moveTo(120, yPosition + 12).lineTo(550, yPosition + 12).stroke();
      doc.fontSize(12).text(value, 130, yPosition);
    };

    drawLineAndText('Pay to', voucherData.payTo, 160);
    drawLineAndText('Pay by', voucherData.paidBy, 200);
    drawLineAndText('Account Head', voucherData.accountHead, 240);
    drawLineAndText('Towards', voucherData.account, 280);

    doc.fontSize(12).text('Amount Rs.', 30, 320);
    doc.moveTo(120, 332).lineTo(250, 332).stroke();
    doc.fontSize(12).text(voucherData.amount, 130, 320);

    doc.fontSize(12).text('The Sum.', 30, 360);
    doc.fontSize(12).text(voucherData.amountRs, 150, 360);

    const amountSectionY = 360;
    const gap = 50;
    const signatureSectionY = amountSectionY + gap;

    const drawSignatureLine = (label, xPosition, yPosition) => {
      doc.moveTo(xPosition, yPosition).lineTo(xPosition + 100, yPosition).stroke();
      doc.fontSize(12).text(label, xPosition, yPosition + 5);
    };

    drawSignatureLine('Prepared By', 30, signatureSectionY);
    drawSignatureLine('Checked By', 180, signatureSectionY);
    drawSignatureLine('Approved By', 330, signatureSectionY);
    drawSignatureLine('Receiver Signature', 480, signatureSectionY);

    doc.end();

    pdfStream.on('finish', async () => {
      try {
        const pdfFileMetadata = {
          name: pdfFileName,
          parents: [driveFolderId],
        };
        const pdfMedia = {
          mimeType: 'application/pdf',
          body: fs.createReadStream(pdfFilePath),
        };
        const pdfUploadResponse = await drive.files.create({
          resource: pdfFileMetadata,
          media: pdfMedia,
          fields: 'id',
        });

        fs.unlinkSync(pdfFilePath);

        res.status(200).send({
          message: 'Data submitted successfully and PDF uploaded!',
          sheetURL: sheetURL,
          pdfFileId: pdfUploadResponse.data.id,
        });
      } catch (error) {
        console.error('Error uploading PDF:', error);
        res.status(500).send({ error: 'Failed to upload PDF' });
      }
    });

    pdfStream.on('error', (error) => {
      console.error('Error creating PDF:', error);
      res.status(500).send({ error: 'Failed to create PDF' });
    });
  } catch (error) {
    console.error('Error submitting data:', error);
    res.status(500).send({ error: 'Failed to submit data' });
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
