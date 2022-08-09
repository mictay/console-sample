import * as fs from 'fs';
import * as path from 'path';
import Excel from 'exceljs';
import DocXTemplater from 'docxtemplater';
import PizZip from 'pizzip';
//const execFileAsync = require('util').promisify(require('child_process').execFile);
import childProcess from 'child_process';

// Path to libreoffice's soffice 
//const pathToSoffice = 'soffice';
const pathToSoffice = '\"D:/Program Files/LibreOffice/program/soffice\"';

// Path to the PDF Files
const outputPathPDF = '../data/output/';

// Path to the Excel File
const filePathDataExcel = path.resolve(__dirname, '../data/excel/iso_2digit_alpha_country_codes.xlsx');
console.log('Excel file path:', filePathDataExcel);

// Path to the Word Template
const filePathDataWord = path.resolve(__dirname, '../data/word/tag-example.docx');
console.log('Word file path:', filePathDataWord);

type ISOCountryCode = {
  isoCode: string;
  countryName: string;
};

const getCellValue = (row:  Excel.Row, cellIndex: number) => {
  const cell = row.getCell(cellIndex);
  
  return cell.value ? cell.value.toString() : '';
};

const main = async () => {

  //DocX is actually a zip content file
  const templateBinaries = fs.readFileSync(filePathDataWord, 'binary');
  const zip = new PizZip(templateBinaries);

  //Open the word template
  const doc = new DocXTemplater(zip, {paragraphLoop: true, linebreaks: true});

  const workbook = new Excel.Workbook();
  const content = await workbook.xlsx.readFile(filePathDataExcel);  

  const worksheet = content.getWorksheet('ISO 2-Digit Alpha Country Code');

  const rowStartIndex = 2;
  const numberOfRows = worksheet.rowCount - 3;

  const rows = worksheet.getRows(rowStartIndex, numberOfRows) ?? [];

  const countries = rows.map((row): ISOCountryCode => {
    return {
      isoCode: getCellValue(row,1),
      countryName: getCellValue(row, 2),
    }
  });

  doc.render(
    {
      last_name: "Doe",
      first_name: "John",
      description: "Hello World",
      phone: "555-555-1234",
      countries: countries
    }
  );

  const buf = doc.getZip().generate( {
    type: "nodebuffer",
    compression: "DEFLATE"
  });


  //Write the merged output.docx file
  fs.writeFileSync(path.resolve(__dirname, "../data/output/output.docx"), buf);

  // Convert it to pdf format with undefined filter (see Libreoffice docs about filter)
  // const inputPath = path.resolve(__dirname, "../data/output/output.docx");
  // const docxBuf = fs.readFileSync(inputPath);
  // let pdfBuf = await libre.convertWithOptionsAsync(docxBuf, '.pdf', undefined, {tmpdir: __dirname + '/data/tmp', unsafeCleanup: true});
  // fs.writeFileSync(outputPathPDF, pdfBuf);

  //let command = `${pathToSoffice} --headless --convert-to pdf ../data/output/output.docx --outdir ../data/output/`;

  // try {
  //   await execFileAsync('soffice', [
  //     '--headless',
  //     '--convert-to',
  //     'pdf',
  //     '../data/output/output.docx',
  //     '--outdir',
  //     '../data/output/'
  //   ]);
  // } catch(e) {
  //   console.error(e);
  // }

  try {
    const soffice = childProcess.execSync(`${pathToSoffice} --headless --convert-to pdf ./data/output/output.docx --outdir ./data/output/`);
  } catch(e) {
    console.error(e);
  }

  console.log(countries);
};

main().then();