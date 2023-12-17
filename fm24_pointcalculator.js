const fs = require("fs");
const path = require("path");
const cheerio = require("cheerio");
const ExcelJS = require("exceljs");

// Klasörün yolu
const folderPath = "files";

// Klasördeki tüm dosyaları al
fs.readdir(folderPath, (err, files) => {
  if (err) {
    console.error("Error reading folder:", err);
    return;
  }

  // .html uzantılı dosyaları filtrele
  const htmlFiles = files.filter((file) => file.endsWith(".html"));

  // En son değiştirilen dosyayı bul
  const latestFile = htmlFiles.reduce((latest, file) => {
    const filePath = path.join(folderPath, file);
    const stat = fs.statSync(filePath);

    if (!latest || stat.mtime > latest.mtime) {
      return { file, mtime: stat.mtime };
    } else {
      return latest;
    }
  }, null);

  if (!latestFile) {
    console.log("No HTML files found in the folder.");
    return;
  }

  // En son dosyanın içeriğini oku
  const latestFilePath = path.join(folderPath, latestFile.file);
  const fileContent = fs.readFileSync(latestFilePath, "utf-8");

  // Cheerio kullanarak HTML içeriğini ayrıştır
  const $ = cheerio.load(fileContent);

  // Baştaki başlıkları al
  const headers = [];
  $("tr:first-child th").each((index, element) => {
    const header = $(element).text().trim();
    headers.push(header);
  });

  // Tüm başlıkların altındaki verileri çek
  const data = {};
  headers.forEach((header) => {
    data[header] = [];
    $(`tr:not(:first-child) td:nth-child(${headers.indexOf(header) + 1})`).each(
      (index, element) => {
        const value = $(element).text().trim();
        data[header].push(value);
      }
    );
  });

  //EXCEL CODE BELOW
  // Yeni bir çalışma kitabı oluştur
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Oyuncu Verileri");

  // Verileri çalışma sayfasına ekleyin
  Object.keys(data).forEach((key, index) => {
    worksheet.getColumn(index + 1).header = key;
    data[key].forEach((value, rowIndex) => {
      worksheet.getCell(rowIndex + 2, index + 1).value = value;
    });
  });

  // GoalKeeper Başlık Ekle
  const newColumnIndex = worksheet.columnCount + 1;
  worksheet.getColumn(newColumnIndex).header = "GK";
  worksheet.getColumn(newColumnIndex + 1).header = "CD";
  worksheet.getColumn(newColumnIndex + 2).header = "FB";
  worksheet.getColumn(newColumnIndex + 3).header = "DM";
  worksheet.getColumn(newColumnIndex + 4).header = "W";
  worksheet.getColumn(newColumnIndex + 5).header = "AM";
  worksheet.getColumn(newColumnIndex + 6).header = "ST";

  // Values and Formulas
  for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex++) {
    const accValue = worksheet.getCell(`C${rowIndex}`).value || 0;
    const worValue = worksheet.getCell(`D${rowIndex}`).value || 0;
    const visValue = worksheet.getCell(`E${rowIndex}`).value || 0;
    const thrValue = worksheet.getCell(`F${rowIndex}`).value || 0;
    const tecValue = worksheet.getCell(`G${rowIndex}`).value || 0;
    const teaValue = worksheet.getCell(`H${rowIndex}`).value || 0;
    const tckValue = worksheet.getCell(`I${rowIndex}`).value || 0;
    const strValue = worksheet.getCell(`J${rowIndex}`).value || 0;
    const staValue = worksheet.getCell(`K${rowIndex}`).value || 0;
    const troValue = worksheet.getCell(`L${rowIndex}`).value || 0;
    const refValue = worksheet.getCell(`M${rowIndex}`).value || 0;
    const punValue = worksheet.getCell(`N${rowIndex}`).value || 0;
    const posValue = worksheet.getCell(`O${rowIndex}`).value || 0;
    const penValue = worksheet.getCell(`P${rowIndex}`).value || 0;
    const pasValue = worksheet.getCell(`Q${rowIndex}`).value || 0;
    const pacValue = worksheet.getCell(`R${rowIndex}`).value || 0;
    const ovovalue = worksheet.getCell(`S${rowIndex}`).value || 0;
    const otbValue = worksheet.getCell(`T${rowIndex}`).value || 0;
    const natValue = worksheet.getCell(`U${rowIndex}`).value || 0;
    const marValue = worksheet.getCell(`V${rowIndex}`).value || 0;
    const lthValue = worksheet.getCell(`W${rowIndex}`).value || 0;
    const lonValue = worksheet.getCell(`X${rowIndex}`).value || 0;
    const ldrValue = worksheet.getCell(`Y${rowIndex}`).value || 0;
    const kicValue = worksheet.getCell(`Z${rowIndex}`).value || 0;
    const jumValue = worksheet.getCell(`AA${rowIndex}`).value || 0;
    const heaValue = worksheet.getCell(`AB${rowIndex}`).value || 0;
    const hanValue = worksheet.getCell(`AC${rowIndex}`).value || 0;
    const freValue = worksheet.getCell(`AD${rowIndex}`).value || 0;
    const flaValue = worksheet.getCell(`AE${rowIndex}`).value || 0;
    const firValue = worksheet.getCell(`AF${rowIndex}`).value || 0;
    const finValue = worksheet.getCell(`AG${rowIndex}`).value || 0;
    const eccValue = worksheet.getCell(`AH${rowIndex}`).value || 0;
    const driValue = worksheet.getCell(`AI${rowIndex}`).value || 0;
    const detValue = worksheet.getCell(`AJ${rowIndex}`).value || 0;
    const decValue = worksheet.getCell(`AK${rowIndex}`).value || 0;
    const croValue = worksheet.getCell(`AL${rowIndex}`).value || 0;
    const corValue = worksheet.getCell(`AM${rowIndex}`).value || 0;
    const cntValue = worksheet.getCell(`AN${rowIndex}`).value || 0;
    const cmpValue = worksheet.getCell(`AO${rowIndex}`).value || 0;
    const comValue = worksheet.getCell(`AP${rowIndex}`).value || 0;
    const cmdValue = worksheet.getCell(`AQ${rowIndex}`).value || 0;
    const braValue = worksheet.getCell(`AR${rowIndex}`).value || 0;
    const balValue = worksheet.getCell(`AS${rowIndex}`).value || 0;
    const antValue = worksheet.getCell(`AT${rowIndex}`).value || 0;
    const agiValue = worksheet.getCell(`AU${rowIndex}`).value || 0;
    const aggValue = worksheet.getCell(`AV${rowIndex}`).value || 0;
    const aerValue = worksheet.getCell(`AW${rowIndex}`).value || 0;

    const gkSum =
      (parseInt(aerValue) * 60 +
        parseInt(cmdValue) * 40 +
        parseInt(comValue) * 30 +
        parseInt(eccValue) * 20 +
        parseInt(firValue) * 30 +
        parseInt(hanValue) * 50 +
        parseInt(kicValue) * 35 +
        parseInt(ovovalue) * 45 +
        parseInt(pasValue) * 45 +
        parseInt(refValue) * 80 +
        parseInt(troValue) * 40 +
        parseInt(thrValue) * 30 +
        parseInt(aggValue) * 40 +
        parseInt(antValue) * 40 +
        parseInt(braValue) * 30 +
        parseInt(cmpValue) * 40 +
        parseInt(cntValue) * 65 +
        parseInt(decValue) * 50 +
        parseInt(detValue) * 20 +
        parseInt(flaValue) * 20 +
        parseInt(ldrValue) * 10 +
        parseInt(posValue) * 40 +
        parseInt(teaValue) * 10 +
        parseInt(visValue) * 40 +
        parseInt(worValue) * 10 +
        parseInt(accValue) * 70 +
        parseInt(agiValue) * 100 +
        parseInt(balValue) * 20 +
        parseInt(jumValue) * 45 +
        parseInt(natValue) * 10 +
        parseInt(pacValue) * 50 +
        parseInt(staValue) * 10 +
        parseInt(strValue) * 70) /
      33;

    const cdSum =
      (parseInt(corValue) * 5 +
        parseInt(croValue) * 1 +
        parseInt(driValue) * 40 +
        parseInt(finValue) * 10 +
        parseInt(firValue) * 35 +
        parseInt(freValue) * 10 +
        parseInt(heaValue) * 55 +
        parseInt(lonValue) * 10 +
        parseInt(lthValue) * 5 +
        parseInt(marValue) * 55 +
        parseInt(pasValue) * 55 +
        parseInt(penValue) * 10 +
        parseInt(tckValue) * 40 +
        parseInt(tecValue) * 35 +
        parseInt(aggValue) * 40 +
        parseInt(antValue) * 50 +
        parseInt(braValue) * 30 +
        parseInt(cmpValue) * 80 +
        parseInt(cntValue) * 50 +
        parseInt(decValue) * 50 +
        parseInt(detValue) * 20 +
        parseInt(flaValue) * 10 +
        parseInt(ldrValue) * 10 +
        parseInt(otbValue) * 10 +
        parseInt(posValue) * 55 +
        parseInt(teaValue) * 10 +
        parseInt(visValue) * 50 +
        parseInt(worValue) * 55 +
        parseInt(accValue) * 90 +
        parseInt(agiValue) * 60 +
        parseInt(balValue) * 35 +
        parseInt(jumValue) * 65 +
        parseInt(natValue) * 10 +
        parseInt(pacValue) * 90 +
        parseInt(staValue) * 30 +
        parseInt(strValue) * 50) /
      36;

    const fbSum =
      (parseInt(corValue) * 30 +
        parseInt(croValue) * 25 +
        parseInt(driValue) * 50 +
        parseInt(finValue) * 10 +
        parseInt(firValue) * 30 +
        parseInt(freValue) * 10 +
        parseInt(heaValue) * 20 +
        parseInt(lonValue) * 10 +
        parseInt(lthValue) * 30 +
        parseInt(marValue) * 45 +
        parseInt(pasValue) * 45 +
        parseInt(penValue) * 10 +
        parseInt(tckValue) * 50 +
        parseInt(tecValue) * 45 +
        parseInt(aggValue) * 45 +
        parseInt(antValue) * 45 +
        parseInt(braValue) * 20 +
        parseInt(cmpValue) * 30 +
        parseInt(cntValue) * 45 +
        parseInt(decValue) * 45 +
        parseInt(detValue) * 20 +
        parseInt(flaValue) * 20 +
        parseInt(ldrValue) * 10 +
        parseInt(otbValue) * 70 +
        parseInt(posValue) * 30 +
        parseInt(teaValue) * 10 +
        parseInt(visValue) * 25 +
        parseInt(worValue) * 90 +
        parseInt(accValue) * 100 +
        parseInt(agiValue) * 60 +
        parseInt(balValue) * 25 +
        parseInt(jumValue) * 40 +
        parseInt(natValue) * 10 +
        parseInt(pacValue) * 90 +
        parseInt(staValue) * 100 +
        parseInt(strValue) * 25) /
      36;

    const dmSum =
      (parseInt(corValue) * 10 +
        parseInt(croValue) * 10 +
        parseInt(driValue) * 45 +
        parseInt(finValue) * 20 +
        parseInt(firValue) * 50 +
        parseInt(freValue) * 30 +
        parseInt(heaValue) * 10 +
        parseInt(lonValue) * 40 +
        parseInt(lthValue) * 5 +
        parseInt(marValue) * 20 +
        parseInt(pasValue) * 65 +
        parseInt(penValue) * 10 +
        parseInt(tckValue) * 35 +
        parseInt(tecValue) * 50 +
        parseInt(aggValue) * 50 +
        parseInt(antValue) * 55 +
        parseInt(braValue) * 30 +
        parseInt(cmpValue) * 60 +
        parseInt(cntValue) * 50 +
        parseInt(decValue) * 65 +
        parseInt(detValue) * 20 +
        parseInt(flaValue) * 50 +
        parseInt(ldrValue) * 10 +
        parseInt(otbValue) * 40 +
        parseInt(posValue) * 65 +
        parseInt(teaValue) * 10 +
        parseInt(visValue) * 55 +
        parseInt(worValue) * 90 +
        parseInt(accValue) * 65 +
        parseInt(agiValue) * 45 +
        parseInt(balValue) * 35 +
        parseInt(jumValue) * 15 +
        parseInt(natValue) * 10 +
        parseInt(pacValue) * 70 +
        parseInt(staValue) * 70 +
        parseInt(strValue) * 35) /
      36;

    const wSum =
      (parseInt(corValue) * 30 +
        parseInt(croValue) * 45 +
        parseInt(driValue) * 55 +
        parseInt(finValue) * 45 +
        parseInt(firValue) * 30 +
        parseInt(freValue) * 10 +
        parseInt(heaValue) * 10 +
        parseInt(lonValue) * 10 +
        parseInt(lthValue) * 30 +
        parseInt(marValue) * 35 +
        parseInt(pasValue) * 50 +
        parseInt(penValue) * 15 +
        parseInt(tckValue) * 35 +
        parseInt(tecValue) * 50 +
        parseInt(aggValue) * 35 +
        parseInt(antValue) * 45 +
        parseInt(braValue) * 15 +
        parseInt(cmpValue) * 30 +
        parseInt(cntValue) * 35 +
        parseInt(decValue) * 35 +
        parseInt(detValue) * 20 +
        parseInt(flaValue) * 20 +
        parseInt(ldrValue) * 10 +
        parseInt(otbValue) * 40 +
        parseInt(posValue) * 35 +
        parseInt(teaValue) * 10 +
        parseInt(visValue) * 35 +
        parseInt(worValue) * 75 +
        parseInt(accValue) * 100 +
        parseInt(agiValue) * 50 +
        parseInt(balValue) * 15 +
        parseInt(jumValue) * 10 +
        parseInt(natValue) * 10 +
        parseInt(pacValue) * 100 +
        parseInt(staValue) * 75 +
        parseInt(strValue) * 30) /
      36;

    const amSum =
      (parseInt(corValue) * 5 +
        parseInt(croValue) * 5 +
        parseInt(driValue) * 65 +
        parseInt(finValue) * 65 +
        parseInt(firValue) * 40 +
        parseInt(freValue) * 30 +
        parseInt(heaValue) * 10 +
        parseInt(lonValue) * 20 +
        parseInt(lthValue) * 1 +
        parseInt(marValue) * 5 +
        parseInt(pasValue) * 50 +
        parseInt(penValue) * 15 +
        parseInt(tckValue) * 15 +
        parseInt(tecValue) * 65 +
        parseInt(aggValue) * 50 +
        parseInt(antValue) * 70 +
        parseInt(braValue) * 20 +
        parseInt(cmpValue) * 35 +
        parseInt(cntValue) * 25 +
        parseInt(decValue) * 40 +
        parseInt(detValue) * 20 +
        parseInt(flaValue) * 20 +
        parseInt(ldrValue) * 10 +
        parseInt(otbValue) * 35 +
        parseInt(posValue) * 10 +
        parseInt(teaValue) * 10 +
        parseInt(visValue) * 30 +
        parseInt(worValue) * 80 +
        parseInt(accValue) * 100 +
        parseInt(agiValue) * 30 +
        parseInt(balValue) * 50 +
        parseInt(jumValue) * 10 +
        parseInt(natValue) * 10 +
        parseInt(pacValue) * 80 +
        parseInt(staValue) * 80 +
        parseInt(strValue) * 30) /
      36;

    const stSum =
      (parseInt(corValue) * 5 +
        parseInt(croValue) * 5 +
        parseInt(driValue) * 75 +
        parseInt(finValue) * 80 +
        parseInt(firValue) * 50 +
        parseInt(freValue) * 5 +
        parseInt(heaValue) * 25 +
        parseInt(lonValue) * 25 +
        parseInt(lthValue) * 1 +
        parseInt(marValue) * 1 +
        parseInt(pasValue) * 40 +
        parseInt(penValue) * 20 +
        parseInt(tckValue) * 5 +
        parseInt(tecValue) * 65 +
        parseInt(aggValue) * 50 +
        parseInt(antValue) * 50 +
        parseInt(braValue) * 20 +
        parseInt(cmpValue) * 35 +
        parseInt(cntValue) * 5 +
        parseInt(decValue) * 45 +
        parseInt(detValue) * 20 +
        parseInt(flaValue) * 25 +
        parseInt(ldrValue) * 10 +
        parseInt(otbValue) * 45 +
        parseInt(posValue) * 5 +
        parseInt(teaValue) * 10 +
        parseInt(visValue) * 20 +
        parseInt(worValue) * 60 +
        parseInt(accValue) * 100 +
        parseInt(agiValue) * 30 +
        parseInt(balValue) * 50 +
        parseInt(jumValue) * 20 +
        parseInt(natValue) * 10 +
        parseInt(pacValue) * 70 +
        parseInt(staValue) * 65 +
        parseInt(strValue) * 25) /
      36;

    worksheet.getCell(`AX${rowIndex}`).value = gkSum;
    worksheet.getCell(`AY${rowIndex}`).value = cdSum;
    worksheet.getCell(`AZ${rowIndex}`).value = fbSum;
    worksheet.getCell(`BA${rowIndex}`).value = dmSum;
    worksheet.getCell(`BB${rowIndex}`).value = wSum;
    worksheet.getCell(`BC${rowIndex}`).value = amSum;
    worksheet.getCell(`BD${rowIndex}`).value = stSum;
  }

  // Dosyayı kaydet
  workbook.xlsx
    .writeFile("oyuncu_verileri.xlsx")
    .then(() => {
      console.log("Excel dosyası başarıyla oluşturuldu: oyuncu_verileri.xlsx");
    })
    .catch((error) => {
      console.error("Excel dosyası oluşturulurken bir hata oluştu:", error);
    });
});
