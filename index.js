async function run() {

const axios = require('axios');
const cheerio = require('cheerio');
const Papa = require('papaparse');
const fs = require('fs');
const ExcelJS = require('exceljs');

const isAlpha = (c) => (97 <= c.charCodeAt(0) && c.charCodeAt(0) <= 122) || (65 <= c.charCodeAt(0) && c.charCodeAt(0) <= 90);

// open file
const file = fs.readFileSync('search_this.csv', 'utf8');
const data = Papa.parse(file, { header: true }).data;

// create spreadsheet
const wb = new ExcelJS.Workbook();
const ws = wb.addWorksheet('SOSCA');
ws.columns = [
  { header: 'Entity #', key: 'entity' },
  { header: 'Registration Date', key: 'regDate' },
  { header: 'Status', key: 'status' },
  { header: 'Entity Name', key: 'entityName' },
  { header: 'Jurisdiction', key: 'jurisdiction' },
  { header: 'Agent', key: 'agent' },
  { header: '# Of Entries', key: 'entries' },
  { header: 'Link', key: 'link' },
];

// iterate over records
for (let row of data) {
  // fetch page
  const url = `https://businesssearch.sos.ca.gov/CBS/SearchResults?filing=&SearchType=LPLLC&SearchCriteria=${encodeURIComponent(row[0])}&SearchSubType=Keyword`;
  console.log(`Loading ${url}`);

  try {
    const res = await axios.get(url);
    if (res.status !== 200) throw new Error('Failed to load the page');
    
    const $ = cheerio.load(res.data);
    const entityTable = $('#enitityTable');
    const tds = entityTable ? entityTable.find('td') : [];

    if (tds.length != 0) {
      tds.each((i, td) => {
        let cellValue = $(td).text().trim();
        if (i % 6 === 3) {
          let off = 0;
          let nl = cellValue.slice(1).indexOf('\n');
          while (!isAlpha(cellValue.charAt(off + nl))) {
            off++;
          }
          cellValue = cellValue.slice(nl + off + 1);
        }
        ws.addRow([cellValue]);
      });
    } else {
      ws.addRow([row[0]]);
    }
    ws.addRow([url]);
  } catch (e) {
    console.error(e);
  }
}

// save the workbook
await wb.xlsx.writeFile('./SOSCA.xlsx');

}

run();