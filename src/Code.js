function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(1000)
    .setHeight(700);
}

function getMasterData() {
  const ss = SpreadsheetApp.openById("1RI0-awz98-Qq8JmMgn_q3--XqDf1JUZjFhM7XtmaYHM");

  const regionSheet = ss.getSheetByName("region");
  const subRegionSheet = ss.getSheetByName("sub_region");
  const areaSheet = ss.getSheetByName("sales_area");
  const provinceSheet = ss.getSheetByName("province");
  const typeSheet = ss.getSheetByName("customer_type");
  const salespersonSheet = ss.getSheetByName("salesperson");

  const regionData = regionSheet.getDataRange().getValues();
  const subRegionData = subRegionSheet.getDataRange().getValues();
  const areaData = areaSheet.getDataRange().getValues();
  const provinceData = provinceSheet.getDataRange().getValues();
  const typeData = typeSheet.getDataRange().getValues();
  const salespersonData = salespersonSheet.getDataRange().getValues();

  const regions = regionData.slice(1).map(row => ({
    code: String(row[0] || "").trim(),
    name: row[1]
  }));

  const subRegions = subRegionData.slice(1).map(row => ({
    code: String(row[0] || "").trim(),
    name: row[1],
    region_code: String(row[2] || "").trim()
  }));

  const areas = areaData.slice(1).map(row => ({
    code: String(row[0] || "").trim(),
    name: row[1],
    sub_region_code: String(row[2] || "").trim(),
    region_code: String(row[3] || "").trim()
  }));

  const provinces = provinceData.slice(1).map(row => ({
    code: String(row[0] || "").trim(),
    name: row[1],
    area_code: String(row[2] || "").trim()
  }));

  const customerTypes = typeData.slice(1).map(row => ({
    code: String(row[0] || "").trim(),
    name: row[1]
  }));

  const salespersons = salespersonData.slice(1).map(row => ({
    area_code: String(row[0] || "").trim(),
    name: row[1]
  }));

  return { regions, subRegions, areas, provinces, customerTypes, salespersons };
}

function addNewRow(formData) {
  const ss = SpreadsheetApp.openById("1RI0-awz98-Qq8JmMgn_q3--XqDf1JUZjFhM7XtmaYHM");
  const sheet = ss.getSheetByName("customer");

  sheet.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
    formData.cv_code,
    formData.customer_name,
    formData.customer_type,
    formData.region_code,
    formData.area_name,
    formData.province_name,
    formData.ref_cv_code,
    formData.ref_customer_name,
    formData.created_by,
    formData.remark,
    formData.doc_channels,
    formData.doc_email_address
  ]);

  return "บันทึกข้อมูลสำเร็จ!";
}

function generatePDF(data) {
  //const templateId = "1z5rf-TGd71Echn9PMdOLQ57LbpADMkPIpcXKGyJzom8";
  const templateId = "1EG_erNwkF0qyYwcqifQ908tzrEDem2qmZ9f3Na-9Sw0";

  // copy template
  const copyFile = DriveApp.getFileById(templateId).makeCopy("temp_form_" + new Date().getTime());
  const copyId = copyFile.getId();
  const doc = DocumentApp.openById(copyId);
  const body = doc.getBody();

  // replace placeholder
  body.replaceText("\\{\\{cv_code\\}\\}", data.cv_code || "");
  body.replaceText("\\{\\{cv_name\\}\\}", data.customer_name || "");
  body.replaceText("\\{\\{customer_type\\}\\}", data.customer_type || "");
  body.replaceText("\\{\\{ref_cv_code\\}\\}", data.ref_cv_code || "");
  body.replaceText("\\{\\{ref_customer\\}\\}", data.ref_customer_name || "");
  body.replaceText("\\{\\{region\\}\\}", data.region_name || "");
  body.replaceText("\\{\\{sub_region\\}\\}", data.sub_region_name || "");
  body.replaceText("\\{\\{sales_area\\}\\}", data.area_name || "");
  body.replaceText("\\{\\{province\\}\\}", data.province_name || "");
  body.replaceText("\\{\\{mail\\}\\}", data.doc_channels?.includes("ไปรษณีย์") ? "☑" : "☐");
  body.replaceText("\\{\\{line\\}\\}", data.doc_channels?.includes("Line Chat") ? "☑" : "☐");
  body.replaceText("\\{\\{email\\}\\}", data.doc_channels?.includes("E-mail") ? "☑" : "☐");
  body.replaceText("\\{\\{doc_email\\}\\}", data.doc_email_address || "");
  body.replaceText("\\{\\{created_by\\}\\}", data.created_by || "");
  body.replaceText("\\{\\{remark\\}\\}", data.remark || "");

  doc.saveAndClose();

  // export PDF
  const pdfBlob = DriveApp.getFileById(copyId).getAs('application/pdf');
  const base64 = Utilities.base64Encode(pdfBlob.getBytes());

  // ลบ copy ทิ้ง
  DriveApp.getFileById(copyId).setTrashed(true);

  return base64;
}

function getReportData(month, year) {
  const ss = SpreadsheetApp.openById("1RI0-awz98-Qq8JmMgn_q3--XqDf1JUZjFhM7XtmaYHM");
  const sheet = ss.getSheetByName("customer");
  const data = sheet.getDataRange().getValues();

  const filtered = data.slice(1).filter(row => {
    if (!row[0]) return false;
    const parts = String(row[0]).split("/");
    const d = new Date(+parts[2], +parts[1] - 1, +parts[0]);
    return d.getMonth() + 1 === month && d.getFullYear() === year;
  });

  const byRegion = {};
  filtered.forEach(row => {
    const region = row[4] || "ไม่ระบุ";
    byRegion[region] = (byRegion[region] || 0) + 1;
  });

  const byType = {};
  filtered.forEach(row => {
    const type = row[3] || "ไม่ระบุ";
    byType[type] = (byType[type] || 0) + 1;
  });

  const bySales = {};
  filtered.forEach(row => {
    const sales = row[9] || "ไม่ระบุ";
    const area = row[5] || "";
    if (!bySales[sales]) bySales[sales] = { count: 0, area: area };
    bySales[sales].count++;
  });

  const regionArr = Object.entries(byRegion).map(([name, count]) => ({ name, count })).sort((a, b) => b.count - a.count);
  const typeArr = Object.entries(byType).map(([name, count]) => ({ name, count })).sort((a, b) => b.count - a.count);
  const salesArr = Object.entries(bySales).map(([name, val]) => ({ name, count: val.count, area: val.area })).sort((a, b) => b.count - a.count).slice(0, 10);

  return {
    total: filtered.length,
    topRegion: regionArr[0] || { name: "-", count: 0 },
    topSales: salesArr[0] || { name: "-", count: 0, area: "" },
    byRegion: regionArr,
    byType: typeArr,
    bySales: salesArr
  };
}