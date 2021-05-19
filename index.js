const xlsx = require("node-xlsx");
const fuzz = require("fuzzball");
const fs = require("fs");
const { format } = require("@fast-csv/format");
const { performance } = require("perf_hooks");
const cliProgress = require("cli-progress");

// === VSTUPNE PARAMETRE ===

// cesta k suboru
const workSheetsFromFile = xlsx.parse("ipet-test-data.xlsx");
// slova ktore budu ignorovane (oddelene ciarkou)
let nehladat = [];
// cislo sheetu
let sheetNumber = 1;
// cislo stlpca s MPN
let columnNumberMPN = 1;
// cislo stlpca s Material
let columnNumberMaterial = 3;
// koeficient podobnosti
let fuzzRatio = 80;
// =========================

if (!fs.existsSync("output")) {
  fs.mkdir("output", (err) => {
    if (err) throw err;
  });
}

const podobnostWS = fs.createWriteStream(`output/podobnost_${fuzzRatio}.csv`);
const stream = format({ headers: true });
stream.pipe(podobnostWS);
let uniqueMaterials = new Set();

// progress bar and info
const bar1 = new cliProgress.SingleBar({}, cliProgress.Presets.shades_classic);
console.log("=== PROGRAM STARTED ===\n!! DO NOT CLOSE THIS TERMINAL !!\nIT WILL TELL YOU WHEN IT IS DONE.\nComputing...");
bar1.start(workSheetsFromFile[sheetNumber - 1].data.length, 0);

const start = performance.now();
for (let kazdy = 1; kazdy < workSheetsFromFile[sheetNumber - 1].data.length; kazdy++) {
  if (!workSheetsFromFile[sheetNumber - 1].data[kazdy][columnNumberMPN - 1]) {
    continue;
  }
  const fromExcelKazdy = workSheetsFromFile[sheetNumber - 1].data[kazdy][columnNumberMPN - 1].toString();
  if (nehladat.includes(fromExcelKazdy) || !/\S/.test(fromExcelKazdy) || uniqueMaterials.has(fromExcelKazdy)) {
    continue;
  }
  uniqueMaterials.add(fromExcelKazdy);
  for (let kazdym = 1; kazdym < workSheetsFromFile[sheetNumber - 1].data.length; kazdym++) {
    const fromExcelKazdym = workSheetsFromFile[sheetNumber - 1].data[kazdym][columnNumberMPN - 1].toString();
    if (kazdy == kazdym) {
      continue;
    }
    let podobnost = fuzz.ratio(fromExcelKazdy, fromExcelKazdym, { full_process: false });
    if (podobnost >= fuzzRatio) {
      let podobneMPN = {
        MPN: fromExcelKazdy,
        material: workSheetsFromFile[sheetNumber - 1].data[kazdy][columnNumberMaterial - 1],
        "z riadku": kazdy + 1,
        "zhoda na": podobnost,
        s: fromExcelKazdym,
        materialom: workSheetsFromFile[sheetNumber - 1].data[kazdym][columnNumberMaterial - 1],
        "na riadku": kazdym + 1,
      };
      uniqueMaterials.add(fromExcelKazdym);
      stream.write(podobneMPN);
    }
  }
  bar1.update(kazdy);
}
bar1.stop();
const end = performance.now();
const inMs = end - start;
console.log(`!! DONE !! in ${msToMinAndSec(inMs)}s`);

stream.end();
fs.writeFileSync("output/unique.json", JSON.stringify([...uniqueMaterials]));

function msToMinAndSec(ms) {
  var minutes = Math.floor(ms / 60000);
  var seconds = ((ms % 60000) / 1000).toFixed(0);
  return seconds == 60 ? minutes + 1 + ":00" : minutes + ":" + (seconds < 10 ? "0" : "") + seconds;
}
