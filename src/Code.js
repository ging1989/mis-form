function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(800)
    .setHeight(600);
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
    code: String(row[0]).trim(),
    name: row[1]
  }));

  const subRegions = subRegionData.slice(1).map(row => ({
    code: String(row[0]).trim(),
    name: row[1],
    region_code: String(row[2]).trim()
  }));

  const areas = areaData.slice(1).map(row => ({
    code: String(row[0]).trim(),
    name: row[1],
    sub_region_code: String(row[2]).trim(),
    region_code: String(row[3]).trim()
  }));

  const provinces = provinceData.slice(1).map(row => ({
    code: String(row[0]).trim(),
    name: row[1],
    area_code: String(row[2]).trim()
  }));

  const customerTypes = typeData.slice(1).map(row => ({
    code: String(row[0]).trim(),
    name: row[1]
  }));

  const salespersons = salespersonData.slice(1).map(row => ({
    area_code: String(row[0]).trim(),
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
    formData.remark
  ]);

  return "บันทึกข้อมูลสำเร็จ!";
}
