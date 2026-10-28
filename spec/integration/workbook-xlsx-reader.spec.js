const fs = require('fs');
const JSZip = require('jszip');
const unzip = require('unzipper');

const testutils = require('../utils/index');

const ExcelJS = verquire('exceljs');

const TEST_FILE_NAME = './spec/out/wb.test.xlsx';
const PADDED_FILE_NAME = './spec/out/wb.padded.xlsx';

// need some architectural changes to make stream read work properly
// because of: shared strings, sheet names, etc are not read in guaranteed order
describe('WorkbookReader', () => {
  describe('Serialise', () => {
    it('xlsx file', function() {
      this.timeout(10000);
      const wb = testutils.createTestBook(new ExcelJS.Workbook(), 'xlsx');

      return wb.xlsx
        .writeFile(TEST_FILE_NAME)
        .then(() => testutils.checkTestBookReader(TEST_FILE_NAME));
    });
  });

  describe('#readFile', () => {
    describe('Row limit', () => {
      it('should bail out if the file contains more rows than the limit', () => {
        const workbook = new ExcelJS.Workbook();
        // The Fibonacci sheet has 19 rows
        return workbook.xlsx
          .readFile('./spec/integration/data/fibonacci.xlsx', {maxRows: 10})
          .then(
            () => {
              throw new Error('Promise unexpectedly fulfilled');
            },
            err => {
              expect(err.message).to.equal('Max row count (10) exceeded');
            }
          );
      });

      it('should fail fast on a huge file', function() {
        this.timeout(5000);
        const workbook = new ExcelJS.Workbook();
        return workbook.xlsx
          .readFile('./spec/integration/data/huge.xlsx', {maxRows: 100})
          .then(
            () => {
              throw new Error('Promise unexpectedly fulfilled');
            },
            err => {
              expect(err.message).to.equal('Max row count (100) exceeded');
            }
          );
      });

      it('should parse fine if the limit is not exceeded', () => {
        const workbook = new ExcelJS.Workbook();
        return workbook.xlsx.readFile(
          './spec/integration/data/fibonacci.xlsx',
          {maxRows: 20}
        );
      });
    });

    describe('Column limit', () => {
      it('should bail out if the file contains more cells than the limit', () => {
        const workbook = new ExcelJS.Workbook();
        // The many-columns sheet has 20 columns in row 2
        return workbook.xlsx
          .readFile('./spec/integration/data/many-columns.xlsx', {
            maxCols: 15,
          })
          .then(
            () => {
              throw new Error('Promise unexpectedly fulfilled');
            },
            err => {
              expect(err.message).to.equal('Max column count (15) exceeded');
            }
          );
      });

      it('should fail fast on a huge file', function() {
        this.timeout(5000);
        const workbook = new ExcelJS.Workbook();
        return workbook.xlsx
          .readFile('./spec/integration/data/huge.xlsx', {maxCols: 10})
          .then(
            () => {
              throw new Error('Promise unexpectedly fulfilled');
            },
            err => {
              expect(err.message).to.equal('Max column count (10) exceeded');
            }
          );
      });

      it('should parse fine if the limit is not exceeded', () => {
        const workbook = new ExcelJS.Workbook();
        return workbook.xlsx.readFile(
          './spec/integration/data/many-columns.xlsx',
          {maxCols: 40}
        );
      });
    });
  });

  describe('#read', () => {
    describe('Row limit', () => {
      it('should bail out if the file contains more rows than the limit', () => {
        const workbook = new ExcelJS.Workbook();
        // The Fibonacci sheet has 19 rows
        return workbook.xlsx
          .read(fs.createReadStream('./spec/integration/data/fibonacci.xlsx'), {
            maxRows: 10,
          })
          .then(
            () => {
              throw new Error('Promise unexpectedly fulfilled');
            },
            err => {
              expect(err.message).to.equal('Max row count (10) exceeded');
            }
          );
      });

      it('should parse fine if the limit is not exceeded', () => {
        const workbook = new ExcelJS.Workbook();
        return workbook.xlsx.read(
          fs.createReadStream('./spec/integration/data/fibonacci.xlsx'),
          {maxRows: 20}
        );
      });
    });
  });

  describe('edit styles in existing file', () => {
    beforeEach(function() {
      this.wb = new ExcelJS.Workbook();
      return this.wb.xlsx.readFile(
        './spec/integration/data/test-row-styles.xlsx'
      );
    });

    it('edit styles of single row instead of all', function() {
      const ws = this.wb.getWorksheet(1);

      ws.eachRow((row, rowNo) => {
        if (rowNo % 5 === 0) {
          row.font = {color: {argb: '00ff00'}};
        }
      });

      expect(ws.getRow(3).font.color.argb).to.be.equal(
        ws.getRow(6).font.color.argb
      );
      expect(ws.getRow(6).font.color.argb).to.be.equal(
        ws.getRow(9).font.color.argb
      );
      expect(ws.getRow(9).font.color.argb).to.be.equal(
        ws.getRow(12).font.color.argb
      );
      expect(ws.getRow(12).font.color.argb).not.to.be.equal(
        ws.getRow(15).font.color.argb
      );
      expect(ws.getRow(15).font.color.argb).not.to.be.equal(
        ws.getRow(18).font.color.argb
      );
      expect(ws.getRow(15).font.color.argb).to.be.equal(
        ws.getRow(10).font.color.argb
      );
      expect(ws.getRow(10).font.color.argb).to.be.equal(
        ws.getRow(5).font.color.argb
      );
    });
  });

  describe('with a spreadsheet that contains formulas', () => {
    before(function() {
      const testContext = this;
      const workbook = new ExcelJS.Workbook();
      return workbook.xlsx
        .read(fs.createReadStream('./spec/integration/data/formulas.xlsx'))
        .then(() => {
          testContext.worksheet = workbook.getWorksheet();
        });
    });

    describe('with a cell that contains a regular formula', () => {
      beforeEach(function() {
        this.cell = this.worksheet.getCell('A2');
      });

      it('should be classified as a formula cell', function() {
        expect(this.cell.type).to.equal(ExcelJS.ValueType.Formula);
      });

      it('should have text corresponding to the evaluated formula result', function() {
        expect(this.cell.text).to.equal('someone@example.com');
      });

      it('should have the formula source', function() {
        expect(this.cell.model.formula).to.equal(
          '_xlfn.CONCAT("someone","@example.com")'
        );
      });
    });

    describe('with a cell that contains a hyperlinked formula', () => {
      beforeEach(function() {
        this.cell = this.worksheet.getCell('A1');
      });

      it('should be classified as a formula cell', function() {
        expect(this.cell.type).to.equal(ExcelJS.ValueType.Hyperlink);
      });

      it('should have text corresponding to the evaluated formula result', function() {
        expect(this.cell.value.text).to.equal('someone@example.com');
      });

      it('should have the formula source', function() {
        expect(this.cell.model.formula).to.equal(
          '_xlfn.CONCAT("someone","@example.com")'
        );
      });

      it('should contain the linked url', function() {
        expect(this.cell.value.hyperlink).to.equal(
          'mailto:someone@example.com'
        );
        expect(this.cell.hyperlink).to.equal('mailto:someone@example.com');
      });
    });
  });

  describe('with a spreadsheet that contains a shared string with an escaped underscore', () => {
    before(function() {
      const testContext = this;
      const workbook = new ExcelJS.Workbook();
      return workbook.xlsx
        .read(
          fs.createReadStream(
            './spec/integration/data/shared_string_with_escape.xlsx'
          )
        )
        .then(() => {
          testContext.worksheet = workbook.getWorksheet();
        });
    });

    it('should decode the underscore', function() {
      const cell = this.worksheet.getCell('A1');
      expect(cell.value).to.equal('_x000D_');
    });
  });

  describe('with a spreadsheet that has an XML parse error in a worksheet', () => {
    let unhandledRejection;
    function unhandledRejectionHandler(err) {
      unhandledRejection = err;
    }
    beforeEach(() => {
      process.on('unhandledRejection', unhandledRejectionHandler);
    });
    afterEach(() => {
      process.removeListener('unhandledRejection', unhandledRejectionHandler);
    });

    it('should reject the promise with the sax error', () => {
      const workbook = new ExcelJS.Workbook();
      return workbook.xlsx
        .readFile('./spec/integration/data/invalid-xml.xlsx')
        .then(
          () => {
            throw new Error('Promise unexpectedly fulfilled');
          },
          err => {
            expect(err.message).to.equal(
              '3:1: text data outside of root node.'
            );
            // Wait a tick before checking for an unhandled rejection
            return new Promise(setImmediate);
          }
        )
        .then(() => {
          expect(unhandledRejection).to.be.undefined();
        });
    });
  });

  describe('with a spreadsheet that is missing some files in the zip container', () => {
    it('should not break', () => {
      const workbook = new ExcelJS.Workbook();
      return workbook.xlsx.readFile(
        './spec/integration/data/missing-bits.xlsx'
      );
    });
  });

  describe('with a spreadsheet that contains images', () => {
    before(function() {
      const testContext = this;
      const workbook = new ExcelJS.Workbook();
      return workbook.xlsx
        .read(fs.createReadStream('./spec/integration/data/images.xlsx'))
        .then(() => {
          testContext.worksheet = workbook.getWorksheet();
        });
    });

    describe('with image`s tl anchor', () => {
      it('Should integer part of col equals nativeCol', function() {
        this.worksheet.getImages().forEach(image => {
          expect(Math.floor(image.range.tl.col)).to.equal(
            image.range.tl.nativeCol
          );
        });
      });
      it('Should integer part of row equals nativeRow', function() {
        this.worksheet.getImages().forEach(image => {
          expect(Math.floor(image.range.tl.row)).to.equal(
            image.range.tl.nativeRow
          );
        });
      });
      it('Should anchor width equals to column width when custom', function() {
        const ws = this.worksheet;

        ws.getImages().forEach(image => {
          const col = ws.getColumn(image.range.tl.nativeCol + 1);

          if (col.isCustomWidth) {
            expect(image.range.tl.colWidth).to.equal(
              Math.floor(col.width * 10000)
            );
          } else {
            expect(image.range.tl.colWidth).to.equal(640000);
          }
        });
      });
      it('Should anchor height equals to row height', function() {
        const ws = this.worksheet;

        ws.getImages().forEach(image => {
          const row = ws.getRow(image.range.tl.nativeRow + 1);

          if (row.height) {
            expect(image.range.tl.rowHeight).to.equal(
              Math.floor(row.height * 10000)
            );
          } else {
            expect(image.range.tl.rowHeight).to.equal(180000);
          }
        });
      });
    });

    describe('with image`s br anchor', () => {
      it('Should integer part of col equals nativeCol', function() {
        this.worksheet.getImages().forEach(image => {
          expect(Math.floor(image.range.br.col)).to.equal(
            image.range.br.nativeCol
          );
        });
      });
      it('Should integer part of row equals nativeRow', function() {
        this.worksheet.getImages().forEach(image => {
          expect(Math.floor(image.range.br.row)).to.equal(
            image.range.br.nativeRow
          );
        });
      });
      it('Should anchor width equals to column width when custom', function() {
        const ws = this.worksheet;

        ws.getImages().forEach(image => {
          const col = ws.getColumn(image.range.br.nativeCol + 1);

          if (col.isCustomWidth) {
            expect(image.range.br.colWidth).to.equal(
              Math.floor(col.width * 10000)
            );
          } else {
            expect(image.range.br.colWidth).to.equal(640000);
          }
        });
      });
      it('Should anchor height equals to row height', function() {
        const ws = this.worksheet;

        ws.getImages().forEach(image => {
          const row = ws.getRow(image.range.br.nativeRow + 1);

          if (row.height) {
            expect(image.range.br.rowHeight).to.equal(
              Math.floor(row.height * 10000)
            );
          } else {
            expect(image.range.br.rowHeight).to.equal(180000);
          }
        });
      });
    });
  });
  describe('with a spreadsheet containing a defined name that kinda looks like it contains a range', () => {
    it('should not crash', () => {
      const workbook = new ExcelJS.Workbook();
      return workbook.xlsx.read(
        fs.createReadStream('./spec/integration/data/bogus-defined-name.xlsx')
      );
    });
  });

  describe('with a zip whose data descriptors carry 64-bit sizes', () => {
    it('should stream the rows from a file path', () => {
      const rows = [];
      const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(
        './spec/integration/data/zip64-data-descriptors.xlsx',
        {sharedStrings: 'cache', worksheets: 'emit'}
      );
      workbookReader.on('worksheet', worksheet =>
        worksheet.on('row', row => rows.push(row.values.slice(1)))
      );
      return new Promise((resolve, reject) => {
        workbookReader.on('error', reject);
        workbookReader.on('finished', resolve);
        workbookReader.read();
      }).then(() => {
        expect(rows).to.deep.equal([
          ['Addend 1', 'Addend 2', 'Sum'],
          [1, 1, {formula: 'SUM(A2,B2)', result: 2}],
        ]);
      });
    });
  });

  describe('with a zip holding entries the reader skips', () => {
    const {file: openFile} = unzip.Open;
    const opened = [];

    before(async () => {
      const zip = await JSZip.loadAsync(
        fs.readFileSync('./spec/integration/data/zip64-data-descriptors.xlsx')
      );
      for (let i = 0; i < 200; i++) {
        zip.file(`xl/media/pad${i}.bin`, Buffer.alloc(1024, i));
      }
      const buffer = await zip.generateAsync({
        type: 'nodebuffer',
        compression: 'DEFLATE',
      });
      fs.writeFileSync(PADDED_FILE_NAME, buffer);
      unzip.Open.file = async (...args) => {
        const directory = await openFile(...args);
        for (const file of directory.files) {
          const {stream} = file;
          file.stream = (...rest) => {
            opened.push(file.path);
            return stream.apply(file, rest);
          };
        }
        return directory;
      };
    });

    after(() => {
      unzip.Open.file = openFile;
    });

    it('should open only the entries it parses', async () => {
      const rows = [];
      const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(
        PADDED_FILE_NAME,
        {sharedStrings: 'cache', worksheets: 'emit'}
      );
      for await (const worksheet of workbookReader) {
        for await (const row of worksheet) {
          rows.push(row.values.slice(1));
        }
      }
      expect(rows).to.have.length(2);
      expect(opened).to.include('xl/worksheets/sheet1.xml');
      expect(opened.filter(path => path.startsWith('xl/media/'))).to.deep.equal(
        []
      );
    });
  });
});
