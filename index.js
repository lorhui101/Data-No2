import xlsx from "xlsx";

const hocSinhRaw = [];

// DSSV1
const DSSV1 = xlsx.readFile("./DSSV1.xlsx");
let DSSV1_Data = xlsx.utils.sheet_to_json(DSSV1.Sheets[DSSV1.SheetNames[0]]);
DSSV1_Data.forEach((raw) => {
  hocSinhRaw.push({
    stt: raw.STT,
    mssv: raw.MSSV,
    ho: raw.Ho5,
    ten: raw.Ten,
    ngay: raw.Ngay,
    thang: raw.Thang,
    nam: raw.Nam,
  });
});

// DSSV2
const DSSV2 = xlsx.readFile("./DSSV2.xlsx");
let DSSV2_Data = xlsx.utils.sheet_to_json(DSSV2.Sheets[DSSV2.SheetNames[0]]);
DSSV2_Data.forEach((raw) => {
  var ngay,
    thang,
    nam = null;

  const rawDate = String(raw["yyyy-mm-dd"]);
  if (rawDate != null) {
    const splitDate = rawDate.split("-");

    if (splitDate.length > 2) {
      ngay = splitDate[2];
      thang = splitDate[1];
      nam = splitDate[0];
    } else if (splitDate.length === 0) {
      ngay = null;
      thang = null;
      nam = rawDate;
    }
  }

  hocSinhRaw.push({
    stt: raw.STT,
    mssv: raw.MSSV,
    ho: raw["Họ Lót"],
    ten: raw["Tên"],
    ngay: ngay,
    thang: thang,
    nam: nam,
  });
});
//DSSV3
const DSSV3 = xlsx.readFile("./DSSV3.xlsx");
let DSSV3_Data = xlsx.utils.sheet_to_json(DSSV3.Sheets[DSSV3.SheetNames[0]]);
DSSV3_Data.forEach((raw) => {
  var ngay,
    thang,
    nam = null;

  const rawDate = String(raw["dd/mm/yyyy"]);
  if (rawDate != null) {
    const splitDate = rawDate.split("/");

    ngay = splitDate[0];
    thang = splitDate[1];
    nam = splitDate[2];
  }

  hocSinhRaw.push({
    stt: raw["__EMPTY"],
    mssv: raw["__EMPTY_1"],
    ho: raw["__EMPTY_2"],
    ten: raw["__EMPTY_3"],
    ngay: ngay,
    thang: thang,
    nam: nam,
  });
});

// DSSV4
const DSSV4 = xlsx.readFile("./DSSV4.xlsx");
let DSSV4_Data = xlsx.utils.sheet_to_json(DSSV4.Sheets[DSSV4.SheetNames[0]]);
DSSV4_Data.forEach((raw) => {
  var ngay,
    thang,
    nam = null;

  const rawDate = String(raw["dd-mm-yyyy"]);
  if (rawDate != null) {
    const splitDate = rawDate.split("-");

    if (splitDate.length > 2) {
      ngay = splitDate[0];
      thang = splitDate[1];
      nam = splitDate[2];
    } else if (splitDate.length === 0) {
      ngay = null;
      thang = null;
      nam = rawDate;
    }
  }

  hocSinhRaw.push({
    stt: raw.STT,
    mssv: raw.MSSV,
    ho: raw["Họ Lót"],
    ten: raw["Tên"],
    ngay: ngay,
    thang: thang,
    nam: nam,
  });
});

const hocSinhFiltered = [];
const hocSinhUnqualifed = [];

hocSinhRaw.forEach((hs, index) => {
  const allFieldFilled = Object.keys(hs).every((k) => {
    if (hs[k] == null || hs[k] == undefined) {
      return false;
    } else {
      return true;
    }
  });

  var dateCorrect = true;
  if (
    Number(hs.ngay) > 31 ||
    Number(hs.ngay) < 1 ||
    Number(hs.thang) > 12 ||
    Number(hs.thang) < 1
  ) {
    dateCorrect = false;
  }

  if (allFieldFilled && dateCorrect) {
    hs.stt = hocSinhFiltered.length + 1;
    hocSinhFiltered.push(hs);
  } else hocSinhUnqualifed.push(hs);
});

// Assuming the year column is in column 1
for (let i = 2; i <= worksheet.rowCount; i++) {
  let cell = worksheet.getCell(`B${i}`);
  let value = cell.value.toString();
  let newValue = value.replace(/,/g, '');
  cell.value = newValue;
}


var FilteredWB = xlsx.utils.book_new();
var FilteredWS = xlsx.utils.json_to_sheet(hocSinhFiltered);
xlsx.utils.book_append_sheet(FilteredWB, FilteredWS, "name");
xlsx.writeFile(FilteredWB, "Filtered.xlsx");

var UnqualifedWB = xlsx.utils.book_new();
var UnqualifedWS = xlsx.utils.json_to_sheet(hocSinhUnqualifed);
xlsx.utils.book_append_sheet(UnqualifedWB, UnqualifedWS, "name");
xlsx.writeFile(UnqualifedWB, "Unqualifed.xlsx");
