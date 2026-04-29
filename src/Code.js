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

function getReportData(month, year) {
  const ss = SpreadsheetApp.openById("1RI0-awz98-Qq8JmMgn_q3--XqDf1JUZjFhM7XtmaYHM");
  const sheet = ss.getSheetByName("customer");
  const data = sheet.getDataRange().getValues();

  const filtered = data.slice(1).filter(row => {
    if (!row[0]) return false;
    const d = new Date(row[0]);
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

  const regionArr = Object.entries(byRegion)
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count);

  const typeArr = Object.entries(byType)
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count);

  const salesArr = Object.entries(bySales)
    .map(([name, val]) => ({ name, count: val.count, area: val.area }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 10);

  return {
    total: filtered.length,
    topRegion: regionArr[0] || { name: "-", count: 0 },
    topSales: salesArr[0] || { name: "-", count: 0, area: "" },
    byRegion: regionArr,
    byType: typeArr,
    bySales: salesArr
  };
}