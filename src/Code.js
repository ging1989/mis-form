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
  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
    <style>
      @font-face {
        font-family: 'NotoSansThai';
        src: url('https://fonts.gstatic.com/s/notosansthai/v25/iJWnBXeUZi_OHPqn4wq6hQ2_hbJ1xyN9wd43SofpAtgv.woff2') format('woff2');
      }
      * { box-sizing: border-box; margin: 0; padding: 0; }
      body { font-family: 'NotoSansThai', 'Tahoma', sans-serif; font-size: 13px; color: #000; padding: 32px 40px; }
      h2 { text-align: center; font-size: 15px; font-weight: 600; margin-bottom: 4px; }
      .sub { text-align: center; font-size: 12px; color: #555; margin-bottom: 32px; }
      .section { margin-bottom: 24px; }
      .row { display: flex; gap: 20px; margin-top: 10px; }
      .field { flex: 1; }
      .field-label { font-size: 11px; color: #888; margin-bottom: 4px; }
      .field-value { border-bottom: 1px solid #ccc; padding-bottom: 5px; min-height: 24px; font-size: 13px; }
      .stitle { font-size: 11px; font-weight: 600; background: #f0f0f0; padding: 5px 10px; border-left: 3px solid #1a3a5c; letter-spacing: 0.5px; text-transform: uppercase; color: #444; }
      .cb-list { display: flex; flex-direction: column; gap: 8px; margin-top: 10px; }
      .cb-row { display: flex; align-items: center; gap: 8px; font-size: 12px; }
      .cb { width: 13px; height: 13px; border: 1px solid #555; display: inline-flex; align-items: center; justify-content: center; font-size: 10px; flex-shrink: 0; }
      .sign-row { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 20px; margin-top: 48px; }
      .sign-box { text-align: center; }
      .sign-line { border-bottom: 1px solid #555; margin: 60px 10px 8px; }
      .sign-label { font-size: 11px; color: #555; }
      .sign-date { font-size: 11px; color: #888; margin-top: 6px; }
    </style></head><body>

    <h2>แบบฟอร์มสำหรับกรอกข้อมูลลูกค้าเข้าระบบ SMARTSOFT</h2>
    <div class="sub">ประจำปี 2026</div>

    <div class="section">
      <div class="stitle">ข้อมูลลูกค้า</div>
      <div class="row">
        <div class="field"><div class="field-label">CV Code</div><div class="field-value">${data.cv_code || ""}</div></div>
        <div class="field" style="flex:2"><div class="field-label">ชื่อลูกค้า</div><div class="field-value">${data.customer_name || ""}</div></div>
        <div class="field"><div class="field-label">ประเภทลูกค้า</div><div class="field-value">${data.customer_type || ""}</div></div>
      </div>
    </div>

    <div class="section">
      <div class="stitle">ลูกค้าบัญชีหลัก (ถ้ามี)</div>
      <div class="row">
        <div class="field"><div class="field-label">CV Code</div><div class="field-value">${data.ref_cv_code || ""}</div></div>
        <div class="field" style="flex:2"><div class="field-label">ชื่อลูกค้าบัญชีหลัก</div><div class="field-value">${data.ref_customer_name || ""}</div></div>
      </div>
    </div>

    <div class="section">
      <div class="stitle">พื้นที่การขาย</div>
      <div class="row">
        <div class="field"><div class="field-label">ภาค</div><div class="field-value">${data.region_name || ""}</div></div>
        <div class="field"><div class="field-label">เขตการขาย</div><div class="field-value">${data.area_name || ""}</div></div>
        <div class="field"><div class="field-label">จังหวัด</div><div class="field-value">${data.province_name || ""}</div></div>
      </div>
    </div>

    <div class="section">
      <div class="stitle">ช่องทางการรับเอกสาร</div>
      <div class="cb-list">
        <div class="cb-row"><div class="cb">${data.doc_channels?.includes("ไปรษณีย์") ? "✓" : ""}</div> จัดส่งทางไปรษณีย์</div>
        <div class="cb-row"><div class="cb">${data.doc_channels?.includes("Line Chat") ? "✓" : ""}</div> จัดส่งผ่าน Line Chat</div>
        <div class="cb-row"><div class="cb">${data.doc_channels?.includes("E-mail") ? "✓" : ""}</div> จัดส่งผ่าน E-mail ${data.doc_email_address ? `(${data.doc_email_address})` : ""}</div>
      </div>
    </div>

    <div class="section">
      <div class="stitle">พนักงานขาย</div>
      <div class="row">
        <div class="field"><div class="field-label">ชื่อพนักงานขาย</div><div class="field-value">${data.created_by || ""}</div></div>
      </div>
    </div>

    ${data.remark ? `
    <div class="section">
      <div class="stitle">หมายเหตุ</div>
      <div style="font-size:13px; padding: 8px 0;">${data.remark}</div>
    </div>` : ""}

    <div class="sign-row">
      <div class="sign-box"><div class="sign-line"></div><div class="sign-label">(ผู้กรอกข้อมูล)</div><div class="sign-date">วันที่ ......./......./.......</div></div>
      <div class="sign-box"><div class="sign-line"></div><div class="sign-label">(ผู้อนุมัติ)</div><div class="sign-date">วันที่ ......./......./.......</div></div>
      <div class="sign-box"><div class="sign-line"></div><div class="sign-label">(ผู้บันทึกข้อมูล)</div><div class="sign-date">วันที่ ......./......./.......</div></div>
    </div>

  </body></html>`;

  const blob = Utilities.newBlob(html, 'text/html', 'form.html');
  const pdfBlob = blob.getAs('application/pdf');
  return Utilities.base64Encode(pdfBlob.getBytes());
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