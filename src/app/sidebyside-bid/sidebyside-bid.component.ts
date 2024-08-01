import { Component, OnInit } from '@angular/core';

import * as XLSX from 'xlsx';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

import * as pdfMake from 'pdfmake/build/pdfmake';
import * as pdfFonts from 'pdfmake/build/vfs_fonts';
pdfMake.vfs = pdfFonts.pdfMake.vfs;

@Component({
  selector: 'app-sidebyside-bid',
  templateUrl: './sidebyside-bid.component.html',
  styleUrls: ['./sidebyside-bid.component.css']
})
export class SidebySideBidComponent implements OnInit {
  tender: any;
  vendors: string[];
  generalQuestions: any[];
  products: any[];

  constructor() {
    this.tender = {
      title: "Tender for A4 bundles and Pens",
      id: "DEU-1-TEN-52",
      startDate: "11/07/2024 03:20:00 PM",
      publishDate: "11/07/2024 03:20:00 PM",
      endDate: "11/07/2024 04:00:00 PM",
      openDate: "11/07/2024 04:01:00 PM",
    };

    this.vendors = [
      "CeylonCraft Stationers",
      "Devon Industries",
      "Al Infinita Pvt Ltd",
    ];

    this.generalQuestions = [
      {
        question: "How long have you been serving this product to the market?",
        answers: ["10 years", "15 years", "20 years"]
      },
      {
        question: "How many years have you been in operations?",
        answers: ["10", "15", "20"]
      }
    ];

    this.products = [
      {
        name: "A4 Paper",
        specs: "90 GSM Color : white Quantity - 1000 Units",
        questions: [
          {
            question: "Total price for this product/service with tax",
            answers: ["1,500,000", "1,400,000", "Not Eligible"]
          },
          {
            question: "Can you provide references or testimonials from previous clients?",
            answers: ["yes", "yes", "No Response"]
          },
          {
            question: "Are the A4s 80 gsm",
            answers: ["yes", "yes", "No Response"]
          }
        ]
      },
      {
        name: "Ball Point Pens",
        specs: "Tip size : 0.7mm Quantity - 100 Units Specifications - ( Color - Blue)",
        questions: [
          {
            question: "Total price for this product/service with tax",
            answers: ["10,000", "15,000", "15,000"]
          },
          {
            question: "Can you provide references or testimonials from previous clients who have purchased A4 with similar specifications?",
            answers: ["yes", "no", "yes"]
          },
          {
            question: "Are the pens refillable or disposable?",
            answers: ["yes", "yes", "yes"]
          }
        ]
      },
      {
        name: "Water Bottles",
        specs: "Capacity : 500ml Quantity - 100 Units Specifications",
        questions: [
          {
            question: "Total price for this product/service with tax",
            answers: ["10,000", "15,000", "15,000"]
          },
          {
            question: "Can you provide references or testimonials from previous clients who have purchased it",
            answers: ["yes", "no", "yes"]
          },
          {
            question: "Are the bottles refillable or disposable?",
            answers: ["yes", "yes", "yes"]
          }
        ]
      }

    ];
  }

  ngOnInit() {
  }

  // exportToExcel(): void {
  //   const wb: XLSX.WorkBook = XLSX.utils.book_new();
  //   const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet([]);

  //   let rowIndex = 0;

  //   // Helper function to add styled header row
  //   const addStyledHeader = (sheet: XLSX.WorkSheet, rowIndex: number, headerText: string, colSpan: number) => {
  //     XLSX.utils.sheet_add_aoa(sheet, [[headerText]], { origin: { r: rowIndex, c: 0 } });
  //     sheet['!merges'] = sheet['!merges'] || [];
  //     sheet['!merges'].push({ s: { r: rowIndex, c: 0 }, e: { r: rowIndex, c: colSpan - 1 } });

  //     // Apply green background to header with white text
  //     for (let i = 0; i < colSpan; i++) {
  //       const cellRef = XLSX.utils.encode_cell({ r: rowIndex, c: i });
  //       if (!sheet[cellRef]) sheet[cellRef] = {};
  //       sheet[cellRef].s = {
  //         fill: { fgColor: { rgb: "008C72" }, patternType: "solid" },
  //         font: { color: { rgb: "FFFFFF" } },
  //         border: {
  //           top: { style: "thin", color: { rgb: "008C72" } },
  //           bottom: { style: "thin", color: { rgb: "008C72" } },
  //           left: { style: "thin", color: { rgb: "008C72" } },
  //           right: { style: "thin", color: { rgb: "008C72" } }
  //         }
  //       };
  //     }

  //     return rowIndex + 1;
  //   };

  //   // Helper function to add borders to a cell range
  //   const addBordersToRange = (sheet: XLSX.WorkSheet, startRow: number, endRow: number, startCol: number, endCol: number) => {
  //     for (let r = startRow; r <= endRow; r++) {
  //       for (let c = startCol; c <= endCol; c++) {
  //         const cellRef = XLSX.utils.encode_cell({ r, c });
  //         if (!sheet[cellRef]) sheet[cellRef] = {};
  //         if (!sheet[cellRef].s) sheet[cellRef].s = {};
  //         sheet[cellRef].s.border = {
  //           top: { style: "thin", color: { rgb: "008C72" } },
  //           bottom: { style: "thin", color: { rgb: "008C72" } },
  //           left: { style: "thin", color: { rgb: "008C72" } },
  //           right: { style: "thin", color: { rgb: "008C72" } }
  //         };
  //       }
  //     }
  //   };

  //   // Add tender details
  //   rowIndex = addStyledHeader(ws, rowIndex, 'Tender Details', 2);
  //   const tenderDetails = [
  //     ['Tender Title', this.tender.title],
  //     ['Tender ID', this.tender.id],
  //     ['Start Date', this.tender.startDate],
  //     ['Publish Date', this.tender.publishDate],
  //     ['End Date', this.tender.endDate],
  //     ['Open Date', this.tender.openDate]
  //   ];
  //   XLSX.utils.sheet_add_aoa(ws, tenderDetails, { origin: { r: rowIndex, c: 0 } });
  //   addBordersToRange(ws, rowIndex, rowIndex + tenderDetails.length - 1, 0, 1);
  //   rowIndex += tenderDetails.length + 1;  // +1 for spacing

  //   // Add general questions
  //   rowIndex = addStyledHeader(ws, rowIndex, 'General Questions', this.vendors.length + 1);
  //   const generalQuestionsData = [
  //     ['Questions', ...this.vendors],
  //     ...this.generalQuestions.map(q => [q.question, ...q.answers])
  //   ];
  //   XLSX.utils.sheet_add_aoa(ws, generalQuestionsData, { origin: { r: rowIndex, c: 0 } });
  //   addBordersToRange(ws, rowIndex, rowIndex + generalQuestionsData.length - 1, 0, this.vendors.length);
  //   rowIndex += generalQuestionsData.length + 1;  // +1 for spacing

  //   // Add product tables
  //   this.products.forEach((product, index) => {
  //     rowIndex = addStyledHeader(ws, rowIndex, `Product ${index + 1}: ${product.name}`, this.vendors.length + 1);
  //     XLSX.utils.sheet_add_aoa(ws, [[product.specs]], { origin: { r: rowIndex, c: 0 } });
  //     addBordersToRange(ws, rowIndex, rowIndex, 0, this.vendors.length);
  //     rowIndex++;

  //     const productData = [
  //       ['Questions', ...this.vendors],
  //       ...product.questions.map(q => [q.question, ...q.answers])
  //     ];
  //     XLSX.utils.sheet_add_aoa(ws, productData, { origin: { r: rowIndex, c: 0 } });
  //     addBordersToRange(ws, rowIndex, rowIndex + productData.length - 1, 0, this.vendors.length);
  //     rowIndex += productData.length + 1;  // +1 for spacing
  //   });

  //   // Auto-size columns
  //   const range = XLSX.utils.decode_range(ws['!ref']);
  //   for (let C = range.s.c; C <= range.e.c; ++C) {
  //     let max_width = 0;
  //     for (let R = range.s.r; R <= range.e.r; ++R) {
  //       const cell = ws[XLSX.utils.encode_cell({ c: C, r: R })];
  //       if (cell && cell.v) {
  //         const width = (cell.v.toString().length + 2) * 1.2;
  //         if (width > max_width) max_width = width;
  //       }
  //     }
  //     ws['!cols'] = ws['!cols'] || [];
  //     ws['!cols'][C] = { width: max_width };
  //   }

  //   XLSX.utils.book_append_sheet(wb, ws, 'Bid Comparison');

  //   // Save the workbook
  //   XLSX.writeFile(wb, 'SideBySideBid_Complete.xlsx');
  // }

  exportToExcel(): void {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Bid Comparison');

    // Define styles
    const headerStyle: Partial<ExcelJS.Style> = {
      font: { color: { argb: 'FFFFFFFF' }, bold: true },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4CAF50' } },
      alignment: { horizontal: 'left', vertical: 'middle' }
    };

    const subheaderStyle: Partial<ExcelJS.Style> = {
      font: { color: { argb: 'FFFFFFFF' }, bold: true },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008C72' } },
      alignment: { horizontal: 'left', vertical: 'middle' }
    };

    const borderStyle: Partial<ExcelJS.Borders> = {
      top: { style: 'thin', color: { argb: 'FF008C72' } },
      left: { style: 'thin', color: { argb: 'FF008C72' } },
      bottom: { style: 'thin', color: { argb: 'FF008C72' } },
      right: { style: 'thin', color: { argb: 'FF008C72' } }
    };

    // Helper function to add styled header
    const addStyledHeader = (rowNumber: number, text: string, style: Partial<ExcelJS.Style>) => {
      const row = worksheet.getRow(rowNumber);
      const cell = row.getCell(1);
      cell.value = text;
      cell.style = { ...style, border: borderStyle };
      worksheet.mergeCells(rowNumber, 1, rowNumber, this.vendors.length + 1);
      return rowNumber + 1;
    };

    // Add tender title
    let rowNumber = addStyledHeader(1, this.tender.title, headerStyle);

    // Add empty row
    rowNumber++;

    // Add tender details
    rowNumber = addStyledHeader(rowNumber, 'Tender Details', subheaderStyle);

    const tenderDetails = [
      ['Tender ID', this.tender.id],
      ['Start Date', this.tender.startDate],
      ['Publish Date', this.tender.publishDate],
      ['End Date', this.tender.endDate],
      ['Open Date', this.tender.openDate]
    ];

    tenderDetails.forEach(detail => {
      const row = worksheet.getRow(rowNumber++);
      row.values = detail;
      row.eachCell(cell => {
        cell.border = borderStyle;
      });
    });

    // Add empty row
    rowNumber++;

    // Add general questions
    rowNumber = addStyledHeader(rowNumber, 'General Questions', headerStyle);

    const questionHeader = worksheet.getRow(rowNumber++);
    questionHeader.values = ['Questions', ...this.vendors];
    questionHeader.eachCell(cell => {
      cell.style = { ...headerStyle, border: borderStyle };
    });

    this.generalQuestions.forEach(q => {
      const row = worksheet.getRow(rowNumber++);
      row.values = [q.question, ...q.answers];
      row.eachCell(cell => {
        cell.border = borderStyle;
      });
    });

    // Add empty row
    rowNumber++;

    // Add product tables
    this.products.forEach((product, index) => {
      rowNumber = addStyledHeader(rowNumber, `Product ${index + 1}: ${product.name}`, headerStyle);

      const specRow = worksheet.getRow(rowNumber++);
      specRow.getCell(1).value = product.specs;
      specRow.getCell(1).border = borderStyle;

      const productHeader = worksheet.getRow(rowNumber++);
      productHeader.values = ['Questions', ...this.vendors];
      productHeader.eachCell(cell => {
        cell.style = { ...headerStyle, border: borderStyle };
      });

      product.questions.forEach(q => {
        const row = worksheet.getRow(rowNumber++);
        row.values = [q.question, ...q.answers];
        row.eachCell(cell => {
          cell.border = borderStyle;
        });
      });

      // Add empty row
      rowNumber++;
    });

    // Auto-fit columns
    (worksheet.columns as ExcelJS.Column[]).forEach((column) => {
      let maxLength = 0;
      (column as any).eachCell({ includeEmpty: true }, (cell: ExcelJS.Cell) => {
        const columnLength = cell.value ? cell.value.toString().length : 10;
        if (columnLength > maxLength) {
          maxLength = columnLength;
        }
      });
      column.width = maxLength < 10 ? 10 : maxLength;
    });

    // Generate Excel file
    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, 'SideBySideBid_Complete.xlsx');
    });
  }


  exportToPDF(): void {
    const docDefinition = {
      pageSize: 'A4',
      pageOrientation: 'landscape',
      content: [
        { text: this.tender.title, style: 'header' },
        {
          text: [
            { text: 'Tender ID: ', bold: true }, this.tender.id, '\n',
            { text: 'Start Date: ', bold: true }, this.tender.startDate, '\n',
            { text: 'Publish Date: ', bold: true }, this.tender.publishDate, '\n',
            { text: 'End Date: ', bold: true }, this.tender.endDate, '\n',
            { text: 'Open Date: ', bold: true }, this.tender.openDate, '\n'
          ],
          margin: [0, 0, 0, 20]
        },
        { text: 'General Questions', style: 'subheader' },
        {
          table: {
            headerRows: 1,
            widths: ['*', ...this.vendors.map(() => '*')],
            body: [
              [{ text: 'Questions', style: 'tableHeader' }, ...this.vendors.map(v => ({ text: v, style: 'tableHeader' }))],
              ...this.generalQuestions.map(q => [q.question, ...q.answers])
            ]
          },
          layout: {
            fillColor: function (rowIndex: number, node: any, columnIndex: number) {
              return rowIndex === 0 ? '#4caf50' : null;
            },
            hLineColor: function (i: number, node: any) {
              return '#008c72';
            },
            vLineColor: function (i: number, node: any) {
              return '#008c72';
            },
            hLineWidth: function (i: number, node: any) {
              return 1;
            },
            vLineWidth: function (i: number, node: any) {
              return 1;
            },
            paddingLeft: function (i: number, node: any) { return 5; },
            paddingRight: function (i: number, node: any) { return 5; },
            paddingTop: function (i: number, node: any) { return 5; },
            paddingBottom: function (i: number, node: any) { return 5; }
          },
          margin: [0, 0, 0, 20]
        },
        ...this.products.map((product, index) => [
          { text: `Product ${index + 1}: ${product.name}`, style: 'subheader' },
          { text: product.specs, margin: [0, 0, 0, 10] },
          {
            table: {
              headerRows: 1,
              widths: ['*', ...this.vendors.map(() => '*')],
              body: [
                [{ text: 'Questions', style: 'tableHeader' }, ...this.vendors.map(v => ({ text: v, style: 'tableHeader' }))],
                ...product.questions.map(q => [q.question, ...q.answers])
              ]
            },
            layout: {
              fillColor: function (rowIndex: number, node: any, columnIndex: number) {
                return rowIndex === 0 ? '#4caf50' : null;
              },
              hLineColor: function (i: number, node: any) {
                return '#008c72';
              },
              vLineColor: function (i: number, node: any) {
                return '#008c72';
              },
              hLineWidth: function (i: number, node: any) {
                return 1;
              },
              vLineWidth: function (i: number, node: any) {
                return 1;
              },
              paddingLeft: function (i: number, node: any) { return 5; },
              paddingRight: function (i: number, node: any) { return 5; },
              paddingTop: function (i: number, node: any) { return 5; },
              paddingBottom: function (i: number, node: any) { return 5; }
            },
            margin: [0, 0, 0, 20]
          }
        ]).reduce((acc, val) => acc.concat(val), [])
      ],
      styles: {
        header: {
          fontSize: 18,
          bold: true,
          margin: [0, 20, 0, 10],
          color: '#4caf50'
        },
        subheader: {
          fontSize: 14,
          bold: true,
          margin: [0, 20, 0, 10],
          color: '#008c72'
        },
        tableHeader: {
          bold: true,
          fontSize: 12,
          color: 'white',
          fillColor: '#4caf50'
        }
      }
    };

    pdfMake.createPdf(docDefinition).download('BidComparison.pdf');
  }

}
