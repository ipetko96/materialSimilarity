const xlsx = require("node-xlsx");
const fuzz = require("fuzzball");
const fs = require("fs");
const { format } = require("@fast-csv/format");

// === VSTUPNE PARAMETRE ===

// cesta k suboru
const workSheetsFromFile = xlsx.parse("Materials_MPN.xlsx");
// slova ktore budu ignorovane (oddelene ciarkou)
let nehladat = [];
// cislo sheetu
let sheetNumber = 1;
// cislo stlpca
let columnNumber = 3;
// koeficient podobnosti
let fuzzRatio = 80;
// =========================

const podobnostWS = fs.createWriteStream("podobnost.csv");
const stream = format({ headers: true });
stream.pipe(podobnostWS);

let uniqueMaterials = new Set();

for (let kazdy = 1; kazdy < workSheetsFromFile[sheetNumber - 1].data.length; kazdy++) {
  if (!workSheetsFromFile[sheetNumber - 1].data[kazdy][columnNumber - 1]) {
    continue;
  }
  const fromExcelKazdy = workSheetsFromFile[sheetNumber - 1].data[kazdy][columnNumber - 1].toString();
  if (nehladat.includes(fromExcelKazdy) || !/\S/.test(fromExcelKazdy) || uniqueMaterials.has(fromExcelKazdy)) {
    continue;
  }
  uniqueMaterials.add(fromExcelKazdy);
  for (let kazdym = 1; kazdym < workSheetsFromFile[sheetNumber - 1].data.length; kazdym++) {
    const fromExcelKazdym = workSheetsFromFile[sheetNumber - 1].data[kazdym][columnNumber - 1].toString();
    if (kazdy == kazdym) {
      continue;
    }
    let podobnost = fuzz.ratio(fromExcelKazdy, fromExcelKazdym, { full_process: false });
    if (podobnost >= fuzzRatio) {
      let podobneMPN = { MPN: fromExcelKazdy, "z riadku": kazdy + 1, "zhoda na": podobnost, s: fromExcelKazdym, "na riadku": kazdym + 1 };
      uniqueMaterials.add(fromExcelKazdym);
      stream.write(podobneMPN);
    }
  }
}
stream.end();
fs.writeFileSync("unique.json", JSON.stringify([...uniqueMaterials]));
