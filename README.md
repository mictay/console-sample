

# Prerequisites
Need to have a recent version of Node.js and npm installed on your computer.


```
npm install typescript --save-dev
npm install @tsconfig/node16 --save-dev
npm install @types/node --save-dev
```

# Translate TS to JS and run
```
npx tsc
node src/main.js
```


# Or run the typescript src directly
```
npx ts-node src/main.ts
```

# References

https://github.com/exceljs/exceljs#access-worksheets  

** Note: Make sure they are .xlsx files (not .xls files)  

https://github.com/open-xml-templating/docxtemplater  

https://docxtemplater.com/docs/get-started-node/  

For the PDF generation, please install LibreOffice, this will
install the libraries we need to use to convert DocX to PDF  

https://www.npmjs.com/package/libreoffice-convert  




