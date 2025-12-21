// ==== ระบบสลับแท็บ ====
  const tabButtons = document.querySelectorAll('.tab-btn');
  const tabPanels  = {
    barcode: document.getElementById('tab-barcode'),
    serial:  document.getElementById('tab-serial'),
    grn:     document.getElementById('tab-grn')
  };

  tabButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.getAttribute('data-tab');

      tabButtons.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      Object.keys(tabPanels).forEach(key => {
        tabPanels[key].classList.toggle('active', key === target);
      });
    });
  });

  const PLANT_URLS = {
    '0301': 'https://opensheet.elk.sh/1x-B1xekpMm4p7fkKucvLjaewtp66uGIp8ZIxJJZAxMk/Sheet1',
    '0304': 'https://opensheet.elk.sh/1miQgObvPdIocjf2Mwn-GZxRvDZy0R5gIl1zhxMWvM-E/Sheet1',
    '0307': 'https://opensheet.elk.sh/1C9vfwtdIO-XjHrDkmp5cUahjmcJ3vk8pTzhFCgHBg1Q/Sheet1',
    '0309': 'https://opensheet.elk.sh/1ntRtlRIndxgEZ3udnh7Nj8LSQiaUdAeAg5j8qd_z8sA/Sheet1',
    '0311': 'https://opensheet.elk.sh/1OKDdqLY_TOjjfLJ58xFiq-WHIDbMrxMcDho6FO5Rq8o/Sheet1',
    '0312': 'https://opensheet.elk.sh/1ilYbVsCqA1cSoUnZwwULCmPkdakobICkyq5JpBJyyAU/Sheet1',
    '0313': 'https://opensheet.elk.sh/12FSbkYSB1zi6xo2WGAsM0jFyp3IkIywLFVF0Zn2axpE/Sheet1',
    '0319': 'https://opensheet.elk.sh/1pYJsp3ZVSwQ1h_Ayqp3BrMTsm-ZtLREU03Qw5WUtgAw/Sheet1',
    '0320': 'https://opensheet.elk.sh/1V5uhH1wekza4FLN2VqtVBfgIA1sbTY_mqq4ohnTmW_M/Sheet1',
    '0326': 'https://opensheet.elk.sh/1gtZLR5Tm574o5xRbdrRm9yFzRukGx_UAzUnoah36cxQ/Sheet1',
    '0366': 'https://opensheet.elk.sh/1NtRsaxPkeNb4sAcTm-Tn4oN-o5F6z_56S1nhTg7kvBc/Sheet1',
    '0369': 'https://opensheet.elk.sh/13NYKhEWkWLLnA2DZZEP-zZEz2r33-roUDPMSWoFR3Qk/Sheet1'
  };
  const EMP_URL = 'https://opensheet.elk.sh/1eqVoLsZxGguEbRCC5rdI4iMVtQ7CK4T3uXRdx8zE3uw/Employee';
  const USER_URL = 'https://opensheet.elk.sh/1eqVoLsZxGguEbRCC5rdI4iMVtQ7CK4T3uXRdx8zE3uw/UserPlantCode';

  // ? NEW: ฐาน Parts  Material mapping
  const PARTS_URL     = 'https://opensheet.elk.sh/1Bs0_VQuEp5hdKX33oIXBebz3avrD4wRoQtK_BUZQfzc/Mainsap_Parts';
  const GRN_URL       = 'https://opensheet.elk.sh/1orr4yVzy2ViyljrH6ENWnBSysTPy4aOLRD4cBjyPdIM/Serial';

  let materialDB = [];
  let employeeDB = [];
  let userDB = [];
  let partsDB    = [];     // ? NEW
  let partsMap   = new Map(); // ? NEW: key=Parts, value=Material
  let grnDB      = [];
  let selectedPlant = '';
  let dataUrl = '';
  let currentUser = null;

  let nextIndex = 1;

  const dataStatusEl = document.getElementById('dataStatus');
  const rowStatusEl  = document.getElementById('rowStatus');
  const tableBody    = document.querySelector('#entryTable tbody');

  const empCodeInput = document.getElementById('empCodeInput');
  const empStatus    = document.getElementById('empStatus');
  const spinner      = document.getElementById('spinner');
  const grnInvoiceFilter = document.getElementById('grnInvoiceFilter');
  const grnDateFilter = document.getElementById('grnDateFilter');
  const grnMaterialFilter = document.getElementById('grnMaterialFilter');
  const grnTableBody = document.querySelector('#grnTable tbody');
  const grnStatusEl = document.getElementById('grnStatus');
  const grnPrintBtn = document.getElementById('grnPrintBtn');
  const grnSelectAll = document.getElementById('grnSelectAll');
  const grnFilters = {
    invoice: '',
    date: '',
    material: ''
  };

  let employeeCode = '';
  let employeeName = '';
  let employeeTeam = '';

  async function loadMaterialData() {
    try {
      if (!dataUrl) {
        throw new Error('missing plant data url');
      }
      const res = await fetch(dataUrl);
      const data = await res.json();
      materialDB = data;
      if (dataStatusEl) {
        dataStatusEl.textContent = `โหลดฐานข้อมูล ${selectedPlant} สำเร็จ (จำนวน ${materialDB.length} รายการ)`;
      }
    } catch (err) {
      console.error(err);
      if (dataStatusEl) {
        dataStatusEl.textContent = 'โหลดฐานข้อมูลไม่สำเร็จ กรุณารีหน้าใหม่';
      }
    }
  }

  async function loadEmployeeData() {
    try {
      const res = await fetch(EMP_URL);
      const data = await res.json();
      employeeDB = data;
      empStatus.textContent = `โหลดข้อมูลพนักงานสำเร็จ (จำนวน ${employeeDB.length} รายการ)`;
    } catch (err) {
      console.error(err);
      empStatus.textContent = 'โหลดข้อมูลพนักงานไม่สำเร็จ กรุณารีหน้าใหม่';
    }
  }

  async function loadUserData() {
    try {
      const res = await fetch(USER_URL);
      const data = await res.json();
      userDB = data || [];
    } catch (err) {
      console.error(err);
      userDB = [];
    }
  }

  // ? NEW: โหลดฐาน Parts และทำ map เพื่อค้นหาไว
  async function loadPartsData() {
    try {
      const res = await fetch(PARTS_URL);
      const data = await res.json();
      partsDB = data || [];

      // สร้าง Map: Parts -> Material
      partsMap = new Map();
      partsDB.forEach(row => {
        const parts =
          row.Parts || row.PARTS || row.parts || row.Part || row.PART || row.part || row['Parts Code'] || '';
        const material =
          row.Material || row.MATERIAL || row.material || row['Material Code'] || '';

        const p = String(parts || '').trim();
        const m = String(material || '').trim();
        if (p && m) partsMap.set(p, m);
      });

      // (ไม่บังคับ) โชว์สถานะรวม
      if (dataStatusEl) {
        dataStatusEl.textContent =
          `โหลดฐานข้อมูล ${selectedPlant} สำเร็จ (จำนวน ${materialDB.length} รายการ) | โหลด Parts mapping สำเร็จ (จำนวน ${partsMap.size} รายการ)`;
      }
    } catch (err) {
      console.error(err);
      // ถ้าโหลดไม่ได้ก็ยังให้ระบบทำงานต่อได้ (แค่ไม่ map parts)
      if (dataStatusEl) {
        dataStatusEl.textContent =
          `โหลดฐานข้อมูล ${selectedPlant} สำเร็จ (จำนวน ${materialDB.length} รายการ) | โหลด Parts mapping ไม่สำเร็จ`;
      }
    }
  }

  function normalizeMaterialCode(raw) {
    if (!raw) return '';
    const trimmed = String(raw).trim();
    const len = trimmed.length;

    if (len === 3)  return '30000' + trimmed;
    if (len === 4)  return '3000'  + trimmed;
    if (len === 5)  return '300'   + trimmed;
    return trimmed;
  }

  // ? NEW: normalize key สำหรับ Parts (กัน whitespace/ชนิดข้อมูล)
  function normalizePartsKey(raw) {
    if (!raw) return '';
    return String(raw).trim();
  }

  // ? NEW: ถ้ากรอกเป็น Parts ให้ map เป็น Material
  function mapPartsToMaterialIfFound(inputValue) {
    const key = normalizePartsKey(inputValue);
    if (!key) return null;

    // ตรงเป๊ะก่อน
    if (partsMap.has(key)) return partsMap.get(key);

    // ลองตัดช่องว่างซ้ำ ๆ หรือรูปแบบแปลก (เผื่อกรณีสแกน)
    const compact = key.replace(/\s+/g, '');
    if (compact !== key && partsMap.has(compact)) return partsMap.get(compact);

    return null;
  }

  function findMaterialInfo(materialCode) {
    if (!materialCode) return null;

    const m = materialDB.find(row => {
      const mat = row.Material || row.MATERIAL || row.material || row['Material Code'];
      return mat && String(mat).trim() === String(materialCode).trim();
    });

    if (!m) return null;

    const description =
      m['Material description'] ||
      m['material description'] ||
      m['MATERIAL DESCRIPTION'] ||
      '';

    const storageBin =
      m['Storage bin'] ||
      m['STORAGE BIN'] ||
      m.StorageBin ||
      m['Storage Bin'] ||
      '';

    const unrestricted =
      m['Unrestricted'] ||
      m['unrestricted'] ||
      m['UNRESTRICTED'] ||
      m['Unrestrict'] ||
      '';

    return { description, storageBin, unrestricted };
  }

  function findEmployeeById(id) {
    if (!id) return null;
    const trimmed = String(id).trim();

    return employeeDB.find(row => {
      const rec = row.IDRec || row.idrec || row.Id || row.ID;
      return rec && String(rec).trim() === trimmed;
    });
  }

  // ? แก้ตรงนี้: กรอก Material แล้วไปเช็ค Parts ก่อน
  function fillRowByMaterial(tr) {
    const inputs = tr.querySelectorAll('input');
    const [inputMat, inputDesc, inputBin, inputRemain, inputQty, inputPack] = inputs;

    let raw = inputMat.value.trim();
    if (!raw) {
      rowStatusEl.textContent = 'กรุณากรอก Material';
      return false;
    }

    // 1) ลอง map จาก Parts -> Material ก่อน
    const mappedMaterial = mapPartsToMaterialIfFound(raw);

    // ถ้าเจอ parts ให้เอา material มาแทน
    if (mappedMaterial) {
      raw = mappedMaterial;
    }

    // 2) normalize เป็น material 0301 format (300xxxx)
    const normalized = normalizeMaterialCode(raw);
    inputMat.value = normalized;

    // 3) ไปหา info จากฐาน 0301
    const info = findMaterialInfo(normalized);

    if (!info || !info.description) {
      inputDesc.value   = '';
      inputBin.value    = '';
      inputRemain.value = '';
      rowStatusEl.innerHTML =
        `<span class="not-found">ไม่พบ Material หรือไม่มี Description: ${normalized} กรุณากรอกใหม่</span>`;
      inputMat.value = '';
      inputMat.focus();
      return false;
    }

    inputDesc.value   = info.description || '';
    inputBin.value    = info.storageBin || '';
    inputRemain.value = info.unrestricted || '';

    rowStatusEl.textContent = '';
    return true;
  }

  function addNewRow() {
    const tr = document.createElement('tr');

    // ปุ่มลบ
    const tdDelete = document.createElement('td');
    const delBtn = document.createElement('button');
    delBtn.type = 'button';
    delBtn.className = 'delete-row-btn';
    delBtn.textContent = 'ลบ';
    delBtn.addEventListener('click', () => {
      tr.remove();
      updateRowIndices();

      const rows = tableBody.querySelectorAll('tr');
      if (rows.length === 0) {
        nextIndex = 1;
        addNewRow();
      }
    });
    tdDelete.appendChild(delBtn);
    tr.appendChild(tdDelete);

    // ลำดับ
    const tdIndex = document.createElement('td');
    tdIndex.className = 'index-cell';
    tdIndex.textContent = nextIndex;
    tr.appendChild(tdIndex);

    // Material
    const tdMat = document.createElement('td');
    const inputMat = document.createElement('input');
    inputMat.type = 'text';
    inputMat.placeholder = 'กรอก Material หรือ Parts หรือสแกนบาร์โค้ด';
    tdMat.appendChild(inputMat);
    tr.appendChild(tdMat);

    // Description
    const tdDesc = document.createElement('td');
    const inputDesc = document.createElement('input');
    inputDesc.type = 'text';
    inputDesc.readOnly = true;
    tdDesc.appendChild(inputDesc);
    tr.appendChild(tdDesc);

    // Storage bin
    const tdBin = document.createElement('td');
    const inputBin = document.createElement('input');
    inputBin.type = 'text';
    inputBin.readOnly = true;
    tdBin.appendChild(inputBin);
    tr.appendChild(tdBin);

    // คงเหลือ
    const tdRemain = document.createElement('td');
    const inputRemain = document.createElement('input');
    inputRemain.type = 'text';
    inputRemain.readOnly = true;
    tdRemain.appendChild(inputRemain);
    tr.appendChild(tdRemain);

    // จำนวน
    const tdQty = document.createElement('td');
    const inputQty = document.createElement('input');
    inputQty.type = 'number';
    inputQty.min = '1';
    inputQty.placeholder = '1';
    inputQty.value = 1;
    tdQty.appendChild(inputQty);
    tr.appendChild(tdQty);

    // Pack
    const tdPack = document.createElement('td');
    const inputPack = document.createElement('input');
    inputPack.type = 'number';
    inputPack.min = '1';
    inputPack.placeholder = '1';
    inputPack.value = 1;
    tdPack.appendChild(inputPack);
    tr.appendChild(tdPack);

    inputMat.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const ok = fillRowByMaterial(tr);
        if (ok) {
          inputQty.focus();
          inputQty.select();
        }
      }
    });

    inputQty.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        if (!inputMat.value.trim() || !inputDesc.value.trim()) {
          rowStatusEl.textContent = 'กรุณากรอก Material ให้ถูกต้องก่อน';
          inputMat.focus();
          return;
        }
        nextIndex++;
        addNewRow();
        updateRowIndices();
      }
    });

    tableBody.appendChild(tr);
    inputMat.focus();
  }

  if (empCodeInput) {
    empCodeInput.addEventListener('input', () => {
      const code = empCodeInput.value.trim();
      employeeCode = code;

      if (!code) {
        empStatus.textContent = 'กรุณากรอกรหัสพนักงาน';
        employeeName = '';
        employeeTeam = '';
        return;
      }

      const emp = findEmployeeById(code);
      if (!emp) {
        empStatus.innerHTML =
          `<span class="not-found">ไม่พบรหัสพนักงาน: ${code}</span>`;
        employeeName = '';
        employeeTeam = '';
        return;
      }

      const name = emp.Name || emp.NAME || emp['Employee Name'] || '';
      const team = emp.Team || emp.TEAM || emp['Team Name'] || '';
      employeeName = name;
      employeeTeam = team;
      empStatus.textContent =
        `รหัสพนักงาน ${code} : ${name} (${team})`;

      const firstMat = document.querySelector('#entryTable tbody tr:first-child input');
      if (firstMat) {
        firstMat.focus();
        firstMat.select();
      }
    });
  }

  function updateRowIndices() {
    const rows = tableBody.querySelectorAll('tr');
    let idx = 1;
    rows.forEach(tr => {
      const indexCell = tr.querySelector('td.index-cell');
      if (indexCell) {
        indexCell.textContent = idx++;
      }
    });
    nextIndex = idx;
  }

  function collectPrintData() {
    const rows = Array.from(tableBody.querySelectorAll('tr'));
    const data = [];

    rows.forEach(tr => {
      const inputs = tr.querySelectorAll('input');
      if (inputs.length < 6) return;

      const [inputMat, inputDesc, inputBin, inputRemain, inputQty, inputPack] = inputs;
      const material = inputMat.value.trim();
      if (!material) return;

      const qty = parseInt(inputQty.value, 10) || 1;

      data.push({
        Material:   material,
        จำนวน:      qty,
        Pack:       inputPack.value || '',
        IDช่าง:     employeeCode || '',
        Serial:     '-',
        Description:inputDesc.value || '',
        Storagebin: inputBin.value || '',
        ชื่อช่าง:   employeeName || '',
        Unit:       'pc',
        แผนก:       employeeTeam || ''
      });
    });

    return data;
  }

  // Small label
  function createPrintLabelSmall(row) {
    const label = document.createElement('div');
    label.className = 'label-small';

    const rowDiv = document.createElement('div');
    rowDiv.className = 'row-small';

    const qrCanvas = document.createElement('canvas');
    qrCanvas.className = 'qr-small';
    rowDiv.appendChild(qrCanvas);

    const textContainer = document.createElement('div');
    textContainer.style.display = 'flex';
    textContainer.style.flexDirection = 'column';

    const materialDiv = document.createElement('div');
    materialDiv.className = 'material-small';
    materialDiv.textContent = row.Material;
    textContainer.appendChild(materialDiv);

    const techNameRaw = row.ชื่อช่าง || '';
    const technicianName = techNameRaw.replace(/^นาย\s+/, '').trim();
    if (technicianName) {
      const technicianDiv = document.createElement('div');
      technicianDiv.className = 'technician-small';
      technicianDiv.textContent = technicianName;
      textContainer.appendChild(technicianDiv);
    }

    const dept = (row.แผนก || '').trim();
    if (dept) {
      const departmentDiv = document.createElement('div');
      departmentDiv.className = 'department-small';
      departmentDiv.textContent = dept;
      textContainer.appendChild(departmentDiv);
    }

    rowDiv.appendChild(textContainer);
    label.appendChild(rowDiv);

    const descText = (row.Description || '').trim();
    if (descText) {
      const desc = document.createElement('div');
      desc.className = 'description-small';
      desc.textContent = descText;
      label.appendChild(desc);
    }

    const locLine = document.createElement('div');
    locLine.style.display = 'flex';
    locLine.style.alignItems = 'center';
    locLine.style.fontSize = '12px';
    locLine.style.fontWeight = 'bold';
    locLine.style.lineHeight = '1.1';
    locLine.style.marginLeft = '2mm';
    locLine.style.marginTop = '1px';
    locLine.style.marginBottom = '0';

    const locTextStr = (row.Storagebin || '').trim();
    if (locTextStr) {
      const locText = document.createElement('span');
      locText.textContent = `Loc: ${locTextStr}`;
      locLine.appendChild(locText);
    }

    const barcodeCanvas = document.createElement('canvas');
    barcodeCanvas.className = 'barcode-small';
    locLine.appendChild(barcodeCanvas);
    label.appendChild(locLine);

    const now = new Date();
    const formattedDate =
      `${now.getFullYear().toString().slice(2)}` +
      `${String(now.getMonth() + 1).padStart(2, '0')}` +
      `${String(now.getDate()).padStart(2, '0')}` +
      `${String(now.getHours()).padStart(2, '0')}` +
      `${String(now.getMinutes()).padStart(2, '0')}`;

    const lot = document.createElement('div');
    lot.className = 'lot-line-small';

    const parts = [];
    parts.push(`Lot: ${formattedDate}`);
    const qtyText  = row.จำนวน ? `(${row.จำนวน})` : '';
    if (qtyText) parts.push(qtyText);
    const packText = (row.Pack || '').trim();
    if (packText) parts.push(packText);
    const unitText = (row.Unit || '').trim();
    if (unitText) parts.push(unitText);

    lot.innerHTML = `<span>${parts.join(' ')}</span>`;
    label.appendChild(lot);

    if (typeof QRCode !== 'undefined') {
      QRCode.toCanvas(qrCanvas, row.Material || 'N/A', {
        width: 40,
        height: 40,
        margin: 0,
        errorCorrectionLevel: 'L'
      }, err => { if (err) console.error('QR Error:', err); });
    }

    if (typeof JsBarcode !== 'undefined' && row.Material) {
      JsBarcode(barcodeCanvas, row.Material, {
        format: 'CODE128',
        width: 1,
        height: 15,
        displayValue: false,
        margin: 0
      });
    }

    return label;
  }

  function buildPrintLabelsSmall(data) {
    const printArea = document.getElementById('printArea');
    printArea.innerHTML = '';

    const container = document.createElement('div');
    container.className = 'container-small';
    printArea.appendChild(container);

    const expanded = [];
    data.forEach(row => {
      const qty = Math.max(1, row.จำนวน);
      for (let i = 0; i < qty; i++) {
        expanded.push({ ...row });
      }
    });

    for (let i = 0; i < expanded.length; i += 2) {
      const sticker = document.createElement('div');
      sticker.className = 'sticker-small';

      const label1 = createPrintLabelSmall(expanded[i]);
      sticker.appendChild(label1);

      if (i + 1 < expanded.length) {
        const label2 = createPrintLabelSmall(expanded[i + 1]);
        sticker.appendChild(label2);
      } else {
        const emptyLabel = document.createElement('div');
        emptyLabel.className = 'label-small';
        const emptyDiv = document.createElement('div');
        emptyDiv.className = 'description-small';
        emptyDiv.textContent = 'ไม่มีข้อมูล';
        emptyLabel.appendChild(emptyDiv);
        sticker.appendChild(emptyLabel);
      }

      container.appendChild(sticker);
    }
  }

  // Big label (barcode tab)
  function createPrintLabelBig(row) {
    const label = document.createElement('div');
    label.className = 'label-big';

    const rowDiv = document.createElement('div');
    rowDiv.className = 'row-big';

    const leftCol = document.createElement('div');
    leftCol.style.display = 'flex';
    leftCol.style.flexDirection = 'column';
    leftCol.style.alignItems = 'flex-start';

    const qrCanvas = document.createElement('canvas');
    qrCanvas.className = 'qr-big';
    leftCol.appendChild(qrCanvas);

    const barcodeCanvas = document.createElement('canvas');
    barcodeCanvas.className = 'barcode-big';
    leftCol.appendChild(barcodeCanvas);

    rowDiv.appendChild(leftCol);

    const textContainer = document.createElement('div');
    textContainer.className = 'text-container-big';

    const materialDiv = document.createElement('div');
    materialDiv.className = 'material-big';
    materialDiv.textContent = row.Material;
    textContainer.appendChild(materialDiv);

    const techNameRaw = row.ชื่อช่าง || '';
    const technicianName = techNameRaw.replace(/^นาย\s+/, '').trim();
    if (technicianName) {
      const technicianDiv = document.createElement('div');
      technicianDiv.className = 'technician-big';
      technicianDiv.textContent = technicianName;
      textContainer.appendChild(technicianDiv);
    }

    const dept = (row.แผนก || '').trim();
    if (dept) {
      const departmentDiv = document.createElement('div');
      departmentDiv.className = 'department-big';
      departmentDiv.textContent = dept;
      textContainer.appendChild(departmentDiv);
    }

    rowDiv.appendChild(textContainer);

    const descText = (row.Description || '').trim();
    if (descText) {
      const desc = document.createElement('div');
      desc.className = 'description-big';
      desc.textContent = descText;
      rowDiv.appendChild(desc);
    }

    label.appendChild(rowDiv);

    const bottomContainer = document.createElement('div');
    bottomContainer.className = 'bottom-container-big';

    const now = new Date();
    const formattedDate =
      `${now.getFullYear().toString().slice(2)}` +
      `${String(now.getMonth() + 1).padStart(2, '0')}` +
      `${String(now.getDate()).padStart(2, '0')}` +
      `${String(now.getHours()).padStart(2, '0')}` +
      `${String(now.getMinutes()).padStart(2, '0')}`;

    const locTextStr = (row.Storagebin || '').trim();
    if (locTextStr) {
      const loc = document.createElement('div');
      loc.className = 'location-big';
      loc.textContent = `Location: ${locTextStr}`;
      bottomContainer.appendChild(loc);
    }

    const lot = document.createElement('div');
    lot.className = 'lot-line-big';
    lot.textContent = `Lot: ${formattedDate} Qty(${row.จำนวน || 1})`;
    bottomContainer.appendChild(lot);

    label.appendChild(bottomContainer);

    const centerPackUnit = document.createElement('div');
    centerPackUnit.className = 'center-pack-unit-big';
    const packText = (row.Pack || '').trim();
    const unitText = (row.Unit || '').trim();
    centerPackUnit.textContent = `${packText} ${unitText}`.trim();
    label.appendChild(centerPackUnit);

    if (typeof QRCode !== 'undefined') {
      QRCode.toCanvas(qrCanvas, row.Material || 'N/A', {
        width: 55,
        height: 55,
        margin: 0,
        errorCorrectionLevel: 'L'
      }, err => { if (err) console.error('QR Error:', err); });
    }

    if (typeof JsBarcode !== 'undefined' && row.Material) {
      JsBarcode(barcodeCanvas, row.Material, {
        format: 'CODE128',
        width: 1,
        height: 25,
        displayValue: false,
        margin: 0
      });
    }

    return label;
  }

  function buildPrintLabelsBig(data) {
    const printArea = document.getElementById('printArea');
    printArea.innerHTML = '';

    const container = document.createElement('div');
    container.className = 'container-big';
    printArea.appendChild(container);

    const expanded = [];
    data.forEach(row => {
      const qty = Math.max(1, row.จำนวน);
      for (let i = 0; i < qty; i++) {
        expanded.push({ ...row });
      }
    });

    expanded.forEach(row => {
      const sticker = document.createElement('div');
      sticker.className = 'sticker-big';
      const label = createPrintLabelBig(row);
      sticker.appendChild(label);
      container.appendChild(sticker);
    });
  }

  function handlePrintSmall() {
    const data = collectPrintData();
    if (!data.length) {
      alert('กรุณากรอก Material อย่างน้อย 1 แถวก่อนพิมพ์');
      return;
    }
    spinner.style.display = 'flex';
    setTimeout(() => {
      buildPrintLabelsSmall(data);
      spinner.style.display = 'none';
      window.print();
    }, 200);
  }

  function handlePrintBig() {
    const data = collectPrintData();
    if (!data.length) {
      alert('กรุณากรอก Material อย่างน้อย 1 แถวก่อนพิมพ์');
      return;
    }
    spinner.style.display = 'flex';
    setTimeout(() => {
      buildPrintLabelsBig(data);
      spinner.style.display = 'none';
      window.print();
    }, 200);
  }

  document.getElementById('printSmallBtn').addEventListener('click', handlePrintSmall);
  document.getElementById('printBigBtn').addEventListener('click', handlePrintBig);

  // ====== ส่วน Serial Tab ======
  const serialInvoiceInput = document.getElementById('serialInvoiceInput');
  const serialMatInput     = document.getElementById('serialMatInput');
  const serialDescInput    = document.getElementById('serialDescInput');
  const serialBinInput     = document.getElementById('serialBinInput');
  const serialPrefixInput  = document.getElementById('serialPrefixInput');
  const serialSuffixInput  = document.getElementById('serialSuffixInput');
  const serialTopStatus    = document.getElementById('serialTopStatus');
  const serialTableBody    = document.querySelector('#serialTable tbody');
  const serialStatusEl     = document.getElementById('serialStatus');
  const serialSummaryInvoice = document.getElementById('serialSummaryInvoice');
  const serialSummaryMaterial = document.getElementById('serialSummaryMaterial');
  const serialSummaryCount = document.getElementById('serialSummaryCount');
  const saveSerialBtn      = document.getElementById('saveSerialBtn');
  const printSerialBtn     = document.getElementById('printSerialBtn');
  let serialNextIndex      = 1;
  const SERIAL_LOCK_MESSAGE = 'กรอก Material ก่อนจึงจะกรอก Serial ในตารางได้';
  const SERIAL_SAVE_URL = 'https://script.google.com/macros/s/AKfycbxp4CCxoAEpPoyYIkPrVZGgZ6ocB4Z-CM6_rqU67Gl6h9UGATCVUYM2d8uWmDYJVrYfVg/exec';
  const confirmModal = document.getElementById('confirmModal');
  const confirmModalMessage = document.getElementById('confirmModalMessage');
  const confirmSaveBtn = document.getElementById('confirmSaveBtn');
  const cancelSaveBtn = document.getElementById('cancelSaveBtn');
  const confirmModalTitle = document.getElementById('confirmModalTitle');
  const cancelDefaultText = cancelSaveBtn ? cancelSaveBtn.textContent : '';

  function getNowString() {
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    const hh = String(now.getHours()).padStart(2, '0');
    const mi = String(now.getMinutes()).padStart(2, '0');
    const ss = String(now.getSeconds()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss}`;
  }

  function updateSerialSummary() {
    if (serialSummaryInvoice) {
      const v = serialInvoiceInput.value.trim() || '-';
      serialSummaryInvoice.querySelector('.serial-chip-value').textContent = v;
    }
    if (serialSummaryMaterial) {
      const v = serialMatInput.value.trim() || '-';
      serialSummaryMaterial.querySelector('.serial-chip-value').textContent = v;
    }
    if (serialSummaryCount) {
      const count = serialTableBody.querySelectorAll('tr').length;
      serialSummaryCount.querySelector('.serial-chip-value').textContent = `${count} แถว`;
    }
  }

  function isSerialMaterialReady() {
    return serialMatInput && serialMatInput.value.trim();
  }

  function updateSerialTableLock() {
    if (!serialTableBody) return;
    const ready = isSerialMaterialReady();
    const serialInputs = serialTableBody.querySelectorAll('.serial-entry-input');
    serialInputs.forEach(input => {
      input.disabled = !ready;
    });
    if (!ready && serialStatusEl) {
      serialStatusEl.textContent = SERIAL_LOCK_MESSAGE;
    } else if (ready && serialStatusEl && serialStatusEl.textContent === SERIAL_LOCK_MESSAGE) {
      serialStatusEl.textContent = 'แถวแรกถูกสร้างให้แล้ว กรอก Serial แล้วกด Enter เพื่อสร้างแถวถัดไป';
    }
  }

  if (serialInvoiceInput) {
    serialInvoiceInput.addEventListener('input', updateSerialSummary);
  }

  function fillSerialTopByMaterial() {
    let mat = serialMatInput.value.trim();
    if (!mat) {
      serialDescInput.value = '';
      serialBinInput.value  = '';
    serialTopStatus.textContent = 'กรอก Material เพื่อดึง Description / Storage bin จากฐานข้อมูล';
    updateSerialSummary();
    updateSerialTableLock();
    return;
  }

    const normalized = normalizeMaterialCode(mat);
    serialMatInput.value = normalized;

    const info = findMaterialInfo(normalized);
    if (!info || !info.description) {
      serialDescInput.value = '';
      serialBinInput.value  = '';
      serialTopStatus.innerHTML =
        `<span class="not-found">ไม่พบ Material หรือไม่มี Description: ${normalized}</span>`;
      updateSerialSummary();
      updateSerialTableLock();
      return;
    }

    serialDescInput.value = info.description || '';
    serialBinInput.value  = info.storageBin || '';
    serialTopStatus.textContent = '';
    updateSerialSummary();
    updateSerialTableLock();
  }

  let serialMatTimer = null;

  if (serialMatInput) {
    serialMatInput.addEventListener('input', () => {
      if (serialMatTimer) clearTimeout(serialMatTimer);
      serialMatTimer = setTimeout(() => {
        fillSerialTopByMaterial();
      }, 400);
    });

    serialMatInput.addEventListener('blur', () => {
      fillSerialTopByMaterial();
    });
  }

  function addSerialRow() {
    const tr = document.createElement('tr');

    const tdDelete = document.createElement('td');
    const delBtn = document.createElement('button');
    delBtn.type = 'button';
    delBtn.textContent = 'ลบ';
    delBtn.className = 'delete-btn';
    delBtn.addEventListener('click', () => {
      tr.remove();
      reindexSerialRows();
    });
    tdDelete.appendChild(delBtn);
    tr.appendChild(tdDelete);

    const tdIndex = document.createElement('td');
    tdIndex.textContent = serialNextIndex;
    tr.appendChild(tdIndex);

    const tdSerial = document.createElement('td');
    const inputSerial = document.createElement('input');
    inputSerial.type = 'text';
    inputSerial.className = 'serial-entry-input';
    inputSerial.placeholder = 'กรอก Serial แล้วกด Enter';
    inputSerial.disabled = !isSerialMaterialReady();
    tdSerial.appendChild(inputSerial);
    tr.appendChild(tdSerial);

    const tdPrefix = document.createElement('td');
    const inputPrefix = document.createElement('input');
    inputPrefix.type = 'text';
    inputPrefix.readOnly = true;
    inputPrefix.value = serialPrefixInput.value || '';
    tdPrefix.appendChild(inputPrefix);
    tr.appendChild(tdPrefix);

    const tdSuffix = document.createElement('td');
    const inputSuffix = document.createElement('input');
    inputSuffix.type = 'text';
    inputSuffix.readOnly = true;
    inputSuffix.value = serialSuffixInput.value || '';
    tdSuffix.appendChild(inputSuffix);
    tr.appendChild(tdSuffix);

    const tdSerialCP = document.createElement('td');
    const inputSerialCP = document.createElement('input');
    inputSerialCP.type = 'text';
    inputSerialCP.readOnly = true;
    tdSerialCP.appendChild(inputSerialCP);
    tr.appendChild(tdSerialCP);

    const tdInvoice = document.createElement('td');
    const inputInvoice = document.createElement('input');
    inputInvoice.type = 'text';
    inputInvoice.readOnly = true;
    inputInvoice.value = serialInvoiceInput.value || '';
    tdInvoice.appendChild(inputInvoice);
    tr.appendChild(tdInvoice);

    const tdMat = document.createElement('td');
    const inputMat = document.createElement('input');
    inputMat.type = 'text';
    inputMat.readOnly = true;
    inputMat.value = serialMatInput.value || '';
    tdMat.appendChild(inputMat);
    tr.appendChild(tdMat);

    const tdDesc = document.createElement('td');
    const inputDesc = document.createElement('input');
    inputDesc.type = 'text';
    inputDesc.readOnly = true;
    inputDesc.value = serialDescInput.value || '';
    tdDesc.appendChild(inputDesc);
    tr.appendChild(tdDesc);

    const tdBin = document.createElement('td');
    const inputBin = document.createElement('input');
    inputBin.type = 'text';
    inputBin.readOnly = true;
    inputBin.value = serialBinInput.value || '';
    tdBin.appendChild(inputBin);
    tr.appendChild(tdBin);

    const tdDate = document.createElement('td');
    const inputDate = document.createElement('input');
    inputDate.type = 'text';
    inputDate.readOnly = true;
    inputDate.value = getNowString();
    tdDate.appendChild(inputDate);
    tr.appendChild(tdDate);

    let committed = false;

    function commitSerialRow() {
      if (committed) return;
      const serialVal = inputSerial.value.trim();

      if (!serialVal) {
        serialStatusEl.innerHTML =
          '<span class="not-found">ไม่อนุญาติให้ขึ้นบรรทัดใหม่ถ้า Serial ยังเป็นค่าว่าง</span>';
        return;
      }

      committed = true;

      inputPrefix.value  = serialPrefixInput.value || '';
      inputSuffix.value  = serialSuffixInput.value || '';
      inputInvoice.value = serialInvoiceInput.value || '';
      inputMat.value     = serialMatInput.value || '';
      inputDesc.value    = serialDescInput.value || '';
      inputBin.value     = serialBinInput.value || '';
      inputDate.value    = getNowString();

      inputSerialCP.value =
        (inputPrefix.value || '') + serialVal + (inputSuffix.value || '');

      serialStatusEl.textContent = 'บันทึก Serial สำเร็จ และสร้างแถวใหม่แล้ว';
      serialNextIndex++;
      addSerialRow();
      updateSerialSummary();
    }

    inputSerial.addEventListener('keydown', e => {
      if (e.key === 'Enter') {
        e.preventDefault();
        commitSerialRow();
      }
    });

    inputSerial.addEventListener('blur', () => {
      if (!inputSerial.value.trim()) return;
      commitSerialRow();
    });

    serialTableBody.appendChild(tr);
    inputSerial.focus();
    updateSerialSummary();
    updateSerialTableLock();
  }

  function reindexSerialRows() {
    const rows = serialTableBody.querySelectorAll('tr');
    let idx = 1;
    rows.forEach(tr => {
      const indexTd = tr.children[1];
      indexTd.textContent = idx++;
    });
    serialNextIndex = idx;

    if (!rows.length) {
      serialNextIndex = 1;
      serialStatusEl.textContent = 'ไม่มีข้อมูลในตาราง สร้างแถวใหม่ให้พร้อมกรอกแล้ว';
      addSerialRow();
    }
    updateSerialSummary();
  }

  function collectSerialPrintData() {
    const rows = Array.from(serialTableBody.querySelectorAll('tr'));
    const data = [];

    rows.forEach(tr => {
      const inputs = tr.querySelectorAll('input');
      if (inputs.length < 9) return;

      const [
        inputSerial,
        inputPrefix,
        inputSuffix,
        inputSerialCP,
        inputInvoice,
        inputMat,
        inputDesc,
        inputBin,
        inputDate
      ] = inputs;

      const material = inputMat.value.trim();
      const serialCP = inputSerialCP.value.trim();
      const desc = inputDesc.value || '';
      const bin = inputBin.value || '';

      if (!material || !serialCP) return;

      data.push({
        Material: material,
        Serial: serialCP,
        Description: desc,
        Storagebin: bin
      });
    });

    data.sort((a, b) => {
      const m = a.Material.localeCompare(b.Material);
      if (m !== 0) return m;
      return a.Serial.localeCompare(b.Serial);
    });

    return data;
  }

  function collectSerialSaveData() {
    const rows = Array.from(serialTableBody.querySelectorAll('tr'));
    const data = [];

    rows.forEach((tr, idx) => {
      const inputs = tr.querySelectorAll('input');
      if (inputs.length < 9) return;

      const [
        inputSerial,
        inputPrefix,
        inputSuffix,
        inputSerialCP,
        inputInvoice,
        inputMat,
        inputDesc,
        inputBin,
        inputDate
      ] = inputs;

      const serialCP = inputSerialCP.value.trim();
      const material = inputMat.value.trim();

      if (!serialCP || !material) return;

      data.push({
        IDrow: idx + 1,
        SerialCP: serialCP,
        Invoice: inputInvoice.value.trim(),
        Material: material,
        Description: inputDesc.value || '',
        'storage bin': inputBin.value || '',
        Date: inputDate.value || ''
      });
    });

    return data;
  }

  function resetSerialForm() {
    if (serialInvoiceInput) serialInvoiceInput.value = '';
    if (serialMatInput) serialMatInput.value = '';
    if (serialDescInput) serialDescInput.value = '';
    if (serialBinInput) serialBinInput.value = '';
    if (serialPrefixInput) serialPrefixInput.value = '';
    if (serialSuffixInput) serialSuffixInput.value = '';
    if (serialTableBody) serialTableBody.innerHTML = '';
    serialNextIndex = 1;
    addSerialRow();
    updateSerialSummary();
    updateSerialTableLock();
    if (serialTopStatus) {
      serialTopStatus.textContent = 'กรอก Invoice และ Material เพื่อดึง Description / Storage bin จากฐานข้อมูล';
    }
    if (serialStatusEl) {
      serialStatusEl.textContent = 'แถวแรกถูกสร้างให้แล้ว กรอก Serial แล้วกด Enter เพื่อสร้างแถวถัดไป';
    }
  }

  function toggleConfirmModal(open) {
    if (!confirmModal) return;
    confirmModal.classList.toggle('open', open);
    confirmModal.setAttribute('aria-hidden', open ? 'false' : 'true');
  }

  function showConfirmModal(message, onConfirm) {
    if (!confirmModal) {
      if (typeof onConfirm === 'function') onConfirm();
      return;
    }

    if (confirmModalMessage) {
      confirmModalMessage.textContent = message;
    }

    toggleConfirmModal(true);

    const handleConfirm = () => {
      toggleConfirmModal(false);
      if (typeof onConfirm === 'function') onConfirm();
      cleanup();
    };

    const handleCancel = () => {
      toggleConfirmModal(false);
      cleanup();
    };

    const handleBackdrop = (event) => {
      if (event.target === confirmModal) {
        toggleConfirmModal(false);
        cleanup();
      }
    };

    const cleanup = () => {
      if (confirmSaveBtn) confirmSaveBtn.removeEventListener('click', handleConfirm);
      if (cancelSaveBtn) cancelSaveBtn.removeEventListener('click', handleCancel);
      confirmModal.removeEventListener('click', handleBackdrop);
    };

    if (confirmSaveBtn) confirmSaveBtn.addEventListener('click', handleConfirm);
    if (cancelSaveBtn) cancelSaveBtn.addEventListener('click', handleCancel);
    confirmModal.addEventListener('click', handleBackdrop);
  }

  function showInfoModal(title, message) {
    if (!confirmModal) return;
    if (confirmModalTitle) confirmModalTitle.textContent = title;
    if (confirmModalMessage) confirmModalMessage.textContent = message;
    if (confirmSaveBtn) confirmSaveBtn.classList.add('is-hidden');
    if (cancelSaveBtn) cancelSaveBtn.textContent = 'ตกลง';

    toggleConfirmModal(true);

    const handleClose = () => {
      toggleConfirmModal(false);
      if (confirmSaveBtn) confirmSaveBtn.classList.remove('is-hidden');
      if (cancelSaveBtn) cancelSaveBtn.textContent = cancelDefaultText || 'ยกเลิก';
      cleanup();
    };

    const handleBackdrop = (event) => {
      if (event.target === confirmModal) {
        handleClose();
      }
    };

    const cleanup = () => {
      if (cancelSaveBtn) cancelSaveBtn.removeEventListener('click', handleClose);
      confirmModal.removeEventListener('click', handleBackdrop);
    };

    if (cancelSaveBtn) cancelSaveBtn.addEventListener('click', handleClose);
    confirmModal.addEventListener('click', handleBackdrop);
  }

  function createLabel(row) {
    const label = document.createElement('div');
    label.className = 'label-big label-serial-big';

    const rowDiv = document.createElement('div');
    rowDiv.className = 'row-big';

    const leftCol = document.createElement('div');
    leftCol.style.display = 'flex';
    leftCol.style.flexDirection = 'column';
    leftCol.style.alignItems = 'flex-start';

    const qrCanvasMat = document.createElement('canvas');
    qrCanvasMat.className = 'qr-big';
    leftCol.appendChild(qrCanvasMat);

    const barcodeMatCanvas = document.createElement('canvas');
    barcodeMatCanvas.className = 'barcode-big';
    leftCol.appendChild(barcodeMatCanvas);

    rowDiv.appendChild(leftCol);

    const textRight = document.createElement('div');
    textRight.className = 'text-container-big';

    const materialDiv = document.createElement('div');
    materialDiv.className = 'material-big';
    materialDiv.textContent = row.Material;
    textRight.appendChild(materialDiv);

    rowDiv.appendChild(textRight);
    label.appendChild(rowDiv);

    const infoBlock = document.createElement('div');
    infoBlock.className = 'info-block-big';

    const descText = (row.Description || '').trim();
    if (descText) {
      const descDiv = document.createElement('div');
      descDiv.className = 'description-big';
      descDiv.textContent = descText;
      descDiv.style.marginTop = '0';
      infoBlock.appendChild(descDiv);
    }

    if (row.Serial) {
      const serialDiv = document.createElement('div');
      serialDiv.className = 'serial-big';
      serialDiv.textContent = `S/N: ${row.Serial}`;
      infoBlock.appendChild(serialDiv);

      const qrSerialCanvas = document.createElement('canvas');
      qrSerialCanvas.className = 'qr-serial-big';
      infoBlock.appendChild(qrSerialCanvas);

      if (typeof QRCode !== 'undefined') {
        QRCode.toCanvas(qrSerialCanvas, row.Serial || 'N/A', {
          width: 40,
          height: 40,
          margin: 0,
          errorCorrectionLevel: 'L'
        }, err => {
          if (err) console.error('QR Error for Serial:', err);
        });
      }
    }

    label.appendChild(infoBlock);

    const bottomContainer = document.createElement('div');
    bottomContainer.className = 'bottom-container-big';

    const now = new Date();
    const formattedDate =
      `${now.getFullYear().toString().slice(2)}` +
      `${String(now.getMonth() + 1).padStart(2, '0')}` +
      `${String(now.getDate()).padStart(2, '0')}` +
      `${String(now.getHours()).padStart(2, '0')}` +
      `${String(now.getMinutes()).padStart(2, '0')}`;

    const locTextStr = (row.Storagebin || '').trim();
    if (locTextStr) {
      const loc = document.createElement('div');
      loc.className = 'location-big';
      loc.textContent = `Location: ${locTextStr}`;
      bottomContainer.appendChild(loc);
    }

    const lot = document.createElement('div');
    lot.className = 'lot-line-big';
    lot.textContent = `Lot: ${formattedDate}`;
    bottomContainer.appendChild(lot);

    label.appendChild(bottomContainer);

    if (typeof QRCode !== 'undefined') {
      QRCode.toCanvas(qrCanvasMat, row.Material || 'N/A', {
        width: 55,
        height: 55,
        margin: 0,
        errorCorrectionLevel: 'L'
      }, err => {
        if (err) console.error('QR Error for Material:', err);
      });
    }

    if (typeof JsBarcode !== 'undefined') {
      JsBarcode(barcodeMatCanvas, row.Material || '00000000', {
        format: 'CODE128',
        width: 1,
        height: 25,
        displayValue: false,
        margin: 0
      });
    }

    return label;
  }

  function buildSerialLabelsForPrint(data) {
    const printArea = document.getElementById('printArea');
    printArea.innerHTML = '';

    const container = document.createElement('div');
    container.className = 'container-big';
    printArea.appendChild(container);

    if (!data.length) {
      container.innerHTML = '<p style="text-align:center;">ไม่พบข้อมูลให้ Print</p>';
      return;
    }

    data.forEach(row => {
      const sticker = document.createElement('div');
      sticker.className = 'sticker-big';

      const label = createLabel(row);
      sticker.appendChild(label);

      container.appendChild(sticker);
    });
  }

  if (saveSerialBtn) {
    const performSerialSave = async () => {
      const data = collectSerialSaveData();
      if (!data.length) {
        alert('ยังไม่มี Serial ในตารางให้บันทึก');
        return;
      }

      try {
        saveSerialBtn.disabled = true;
        serialStatusEl.textContent = 'กำลังบันทึกข้อมูลลง Google Sheet...';

        const body = new URLSearchParams({
          rows: JSON.stringify(data)
        });

        const res = await fetch(SERIAL_SAVE_URL, {
          method: 'POST',
          body
        });

        const result = await res.json();
        if (!result || result.ok !== true) {
          throw new Error((result && result.error) || 'บันทึกไม่สำเร็จ');
        }

        serialStatusEl.textContent = `บันทึกสำเร็จ ${result.inserted || data.length} แถว`;
        showInfoModal('บันทึกสำเร็จ', `บันทึกข้อมูล ${result.inserted || data.length} แถวเรียบร้อยแล้ว`);
        resetSerialForm();
      } catch (err) {
        serialStatusEl.textContent = `บันทึกไม่สำเร็จ: ${err.message || err}`;
      } finally {
        saveSerialBtn.disabled = false;
      }
    };

    saveSerialBtn.addEventListener('click', () => {
      showConfirmModal('ต้องการบันทึกข้อมูลลง Google Sheet ใช่หรือไม่', performSerialSave);
    });
  }

  if (printSerialBtn) {
    printSerialBtn.addEventListener('click', () => {
      const data = collectSerialPrintData();
      if (!data.length) {
        alert('ยังไม่มี Serial ในตารางสำหรับพิมพ์');
        return;
      }

      spinner.style.display = 'flex';

      setTimeout(() => {
        buildSerialLabelsForPrint(data);

        setTimeout(() => {
          spinner.style.display = 'none';
          window.print();
        }, 300);
      }, 200);
    });
  }

  const loginGate = document.getElementById('loginGate');

  function normalizeGrnRow(row) {
    return {
      SerialCP: row.SerialCP || row.Serial || row.serialCP || row.serialcp || '',
      Invoice: row.Invoice || row.invoice || '',
      Material: row.Material || row.material || '',
      Description: row.Description || row.description || '',
      Storagebin: row["storage bin"] || row["Storage bin"] || row.StorageBin || row.storageBin || row.storagebin || '',
      Date: row.Date || row.date || ''
    };
  }

  function applyGrnFilters(data, filters) {
    return data.filter(row => {
      if (filters.invoice && row.Invoice !== filters.invoice) return false;
      if (filters.date && row.Date !== filters.date) return false;
      if (filters.material && row.Material !== filters.material) return false;
      return true;
    });
  }

  function getUniqueValues(data, key) {
    const set = new Set();
    data.forEach(row => {
      const val = row[key];
      if (val) set.add(val);
    });
    return Array.from(set).sort((a, b) => String(a).localeCompare(String(b)));
  }

  function setGrnSelectOptions(selectEl, values, selectedValue) {
    if (!selectEl) return selectedValue || '';
    selectEl.innerHTML = '';
    const allOption = document.createElement('option');
    allOption.value = '';
    allOption.textContent = 'ทั้งหมด';
    selectEl.appendChild(allOption);

    values.forEach(val => {
      const option = document.createElement('option');
      option.value = val;
      option.textContent = val;
      selectEl.appendChild(option);
    });

    if (selectedValue && values.includes(selectedValue)) {
      selectEl.value = selectedValue;
    } else {
      selectEl.value = '';
    }

    return selectEl.value;
  }

  function renderGrnTable(rows) {
    if (!grnTableBody) return;
    grnTableBody.innerHTML = '';

    if (grnSelectAll) grnSelectAll.checked = false;
    rows.forEach(row => {
      const tr = document.createElement('tr');

      const tdSelect = document.createElement('td');
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.className = 'grn-checkbox';
      checkbox.dataset.serial = row.SerialCP || '';
      checkbox.dataset.material = row.Material || '';
      checkbox.dataset.description = row.Description || '';
      checkbox.dataset.storagebin = row.Storagebin || '';
      tdSelect.appendChild(checkbox);
      tr.appendChild(tdSelect);

      const tdSerial = document.createElement('td');
      tdSerial.textContent = row.SerialCP || '-';
      tr.appendChild(tdSerial);

      const tdInvoice = document.createElement('td');
      tdInvoice.textContent = row.Invoice || '-';
      tr.appendChild(tdInvoice);

      const tdMaterial = document.createElement('td');
      tdMaterial.textContent = row.Material || '-';
      tr.appendChild(tdMaterial);

      const tdDesc = document.createElement('td');
      tdDesc.textContent = row.Description || '-';
      tr.appendChild(tdDesc);

      const tdBin = document.createElement('td');
      tdBin.textContent = row.Storagebin || '-';
      tr.appendChild(tdBin);

      const tdDate = document.createElement('td');
      tdDate.textContent = row.Date || '-';
      tr.appendChild(tdDate);

      grnTableBody.appendChild(tr);
    });
  }

  function updateGrnFiltersAndTable() {
    if (!grnDB.length) {
      renderGrnTable([]);
      if (grnStatusEl) grnStatusEl.textContent = 'ไม่พบข้อมูล GRN';
      return;
    }

    const invoiceValues = getUniqueValues(
      applyGrnFilters(grnDB, { date: grnFilters.date, material: grnFilters.material }),
      'Invoice'
    );
    const dateValues = getUniqueValues(
      applyGrnFilters(grnDB, { invoice: grnFilters.invoice, material: grnFilters.material }),
      'Date'
    );
    const materialValues = getUniqueValues(
      applyGrnFilters(grnDB, { invoice: grnFilters.invoice, date: grnFilters.date }),
      'Material'
    );

    grnFilters.invoice = setGrnSelectOptions(grnInvoiceFilter, invoiceValues, grnFilters.invoice);
    grnFilters.date = setGrnSelectOptions(grnDateFilter, dateValues, grnFilters.date);
    grnFilters.material = setGrnSelectOptions(grnMaterialFilter, materialValues, grnFilters.material);

    const filtered = applyGrnFilters(grnDB, grnFilters);
    renderGrnTable(filtered);
    if (grnStatusEl) grnStatusEl.textContent = `พบข้อมูล ${filtered.length} รายการ`;
  }

  async function loadGrnData() {
    if (grnStatusEl) grnStatusEl.textContent = 'กำลังโหลดข้อมูล GRN...';
    try {
      const res = await fetch(GRN_URL);
      const data = await res.json();
      grnDB = Array.isArray(data) ? data.map(normalizeGrnRow) : [];
      updateGrnFiltersAndTable();
    } catch (err) {
      if (grnStatusEl) grnStatusEl.textContent = 'โหลดข้อมูล GRN ไม่สำเร็จ';
      grnDB = [];
      renderGrnTable([]);
    }
  }

  if (grnInvoiceFilter) {
    grnInvoiceFilter.addEventListener('change', () => {
      grnFilters.invoice = grnInvoiceFilter.value;
      updateGrnFiltersAndTable();
    });
  }
  if (grnDateFilter) {
    grnDateFilter.addEventListener('change', () => {
      grnFilters.date = grnDateFilter.value;
      updateGrnFiltersAndTable();
    });
  }
  if (grnMaterialFilter) {
    grnMaterialFilter.addEventListener('change', () => {
      grnFilters.material = grnMaterialFilter.value;
      updateGrnFiltersAndTable();
    });
  }

  if (grnStatusEl) {
    loadGrnData();
  }

  if (grnPrintBtn) {
    grnPrintBtn.addEventListener('click', () => {
      if (!grnTableBody) return;
      const selected = Array.from(grnTableBody.querySelectorAll('input.grn-checkbox:checked'));

      if (!selected.length) {
        alert('กรุณาเลือกอย่างน้อย 1 รายการเพื่อพิมพ์');
        return;
      }

      const data = selected.map(item => ({
        Material: item.dataset.material || '',
        Serial: item.dataset.serial || '',
        Description: item.dataset.description || '',
        Storagebin: item.dataset.storagebin || ''
      })).filter(row => row.Material && row.Serial);

      if (!data.length) {
        alert('ข้อมูลไม่ครบสำหรับพิมพ์');
        return;
      }

      spinner.style.display = 'flex';

      setTimeout(() => {
        buildSerialLabelsForPrint(data);

        setTimeout(() => {
          spinner.style.display = 'none';
          window.print();
        }, 300);
      }, 200);
    });
  }



  if (grnSelectAll) {
    grnSelectAll.addEventListener('change', () => {
      if (!grnTableBody) return;
      const checked = grnSelectAll.checked;
      grnTableBody.querySelectorAll('input.grn-checkbox').forEach((checkbox) => {
        checkbox.checked = checked;
      });
    });
  }

  const grnTableHead = document.querySelector('#grnTable thead');
  if (grnTableHead) {
    grnTableHead.addEventListener('click', async (event) => {
      const th = event.target.closest('th');
      if (!th) return;
      if (!grnTableBody) return;
      const headerText = th.textContent.trim();
      if (!headerText) return;
      const colIndex = Array.from(th.parentElement.children).indexOf(th);
      if (colIndex < 0) return;
      if (colIndex === 0) return;

      const values = Array.from(grnTableBody.querySelectorAll('tr'))
        .map(row => {
          const cell = row.children[colIndex];
          return cell ? cell.textContent.trim() : '';
        })
        .filter(val => val !== '');

      if (!values.length) {
        if (grnStatusEl) grnStatusEl.textContent = 'ไม่มีข้อมูลให้คัดลอก';
        return;
      }

      const text = values.join('\n');

      try {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          await navigator.clipboard.writeText(text);
        } else {
          const temp = document.createElement('textarea');
          temp.value = text;
          document.body.appendChild(temp);
          temp.select();
          document.execCommand('copy');
          temp.remove();
        }
        if (grnStatusEl) grnStatusEl.textContent = `คัดลอกคอลัมน์: ${headerText} (${values.length} รายการ)`;
      } catch (err) {
        if (grnStatusEl) grnStatusEl.textContent = 'คัดลอกไม่สำเร็จ';
      }
    });
  }
  const loginUserInput = document.getElementById('loginUserInput');
  const loginPassInput = document.getElementById('loginPassInput');
  const loginRemember = document.getElementById('loginRemember');
  const loginBtn = document.getElementById('loginBtn');
  const loginStatus = document.getElementById('loginStatus');
  const menuBtn = document.getElementById('menuBtn');
  const menuPanel = document.getElementById('menuPanel');
  const logoutBtn = document.getElementById('logoutBtn');
  const menuUserId = document.getElementById('menuUserId');
  const menuUserName = document.getElementById('menuUserName');
  const menuUserNickname = document.getElementById('menuUserNickname');
  const menuUserPlant = document.getElementById('menuUserPlant');
  const menuUserPlantCode = document.getElementById('menuUserPlantCode');
  const plantGate = document.getElementById('plantGate');
  const plantSelect = document.getElementById('plantSelect');
  const plantEnterBtn = document.getElementById('plantEnterBtn');
  const backToPlantBtn = document.getElementById('backToPlantBtn');
  const appShell = document.querySelector('.app-shell');
  const badgeText = document.querySelector('.badge span:last-child');

  const LOGIN_STORAGE_KEY = 'barcode-login';

  function getUserValue(row, keys) {
    for (const key of keys) {
      if (row[key] !== undefined && row[key] !== null) {
        const value = String(row[key]).trim();
        if (value) return value;
      }
    }
    return '';
  }

  function mapUserRow(row) {
    return {
      id: getUserValue(row, ['IDUser', 'IdUser', 'iduser', 'user', 'User']),
      name: getUserValue(row, ['Name', 'NAME', 'name']),
      nickname: getUserValue(row, ['Nicname', 'Nickname', 'NICKNAME', 'nickName', 'nick']),
      plant: getUserValue(row, ['Plant', 'PLANT', 'plant']),
      plantCode: getUserValue(row, ['PlantCode', 'Plant Code', 'PLANTCODE', 'plantcode'])
    };
  }

  function findUserById(id) {
    const target = String(id || '').trim();
    if (!target) return null;
    const row = userDB.find(item => {
      const value = getUserValue(item, ['IDUser', 'IdUser', 'iduser', 'user', 'User']);
      return value === target;
    });
    return row ? mapUserRow(row) : null;
  }

  function setMenuField(el, value) {
    if (el) el.textContent = value || '-';
  }

  function setMenuUserInfo(user) {
    setMenuField(menuUserId, user ? user.id : '-');
    setMenuField(menuUserName, user ? user.name : '-');
    setMenuField(menuUserNickname, user ? user.nickname : '-');
    setMenuField(menuUserPlant, user ? user.plant : '-');
    setMenuField(menuUserPlantCode, user ? user.plantCode : '-');
  }

  function getLast4(value) {
    const text = String(value || '').trim();
    return text.slice(-4);
  }

  function saveLogin(id, password) {
    if (!loginRemember || !loginRemember.checked) {
      localStorage.removeItem(LOGIN_STORAGE_KEY);
      return;
    }
    const payload = {
      id: id || '',
      password: password || ''
    };
    localStorage.setItem(LOGIN_STORAGE_KEY, JSON.stringify(payload));
  }

  function restoreLogin() {
    try {
      const raw = localStorage.getItem(LOGIN_STORAGE_KEY);
      if (!raw) return null;
      return JSON.parse(raw);
    } catch (err) {
      return null;
    }
  }

  function showLoginStatus(message, isError = false) {
    if (!loginStatus) return;
    loginStatus.textContent = message;
    if (isError) {
      loginStatus.classList.add('not-found');
    } else {
      loginStatus.classList.remove('not-found');
    }
  }

  function handleLogin() {
    const id = loginUserInput ? loginUserInput.value.trim() : '';
    const password = loginPassInput ? loginPassInput.value.trim() : '';

    if (!id || !password) {
      showLoginStatus('กรุณากรอก User และ Password', true);
      return;
    }

    const user = findUserById(id);
    if (!user) {
      showLoginStatus('ไม่พบ User นี้ในระบบ', true);
      return;
    }

    if (password !== getLast4(id)) {
      showLoginStatus('Password ไม่ถูกต้อง (ใช้ 4 ตัวท้ายของ IDUser)', true);
      return;
    }

    currentUser = user;
    setMenuUserInfo(user);
    showLoginStatus('');
    saveLogin(id, password);

    if (loginGate) loginGate.style.display = 'none';
    if (plantGate) plantGate.style.display = 'flex';
  }

  function toggleMenuPanel(forceOpen) {
    if (!menuPanel) return;
    const shouldOpen = typeof forceOpen === 'boolean'
      ? forceOpen
      : !menuPanel.classList.contains('open');
    menuPanel.classList.toggle('open', shouldOpen);
    menuPanel.setAttribute('aria-hidden', shouldOpen ? 'false' : 'true');
  }

  function handleLogout() {
    currentUser = null;
    localStorage.removeItem(LOGIN_STORAGE_KEY);
    if (loginRemember) loginRemember.checked = false;
    if (loginUserInput) loginUserInput.value = '';
    if (loginPassInput) loginPassInput.value = '';
    showLoginStatus('');
    setMenuUserInfo(null);
    toggleMenuPanel(false);
    resetAppState();
    if (plantSelect) plantSelect.value = '';
    if (plantGate) plantGate.style.display = 'none';
    if (appShell) appShell.classList.add('is-hidden');
    if (loginGate) loginGate.style.display = 'flex';
  }

  function setPlantAndStart() {
    if (!plantSelect) return;
    const plant = plantSelect.value;
    if (!plant || !PLANT_URLS[plant]) {
      alert('กรุณาเลือก Plant ก่อนเริ่มใช้งาน');
      return;
    }

    selectedPlant = plant;
    dataUrl = PLANT_URLS[plant];

    if (badgeText) {
      badgeText.textContent = `โหมดทำงาน: Barcode (${selectedPlant})`;
    }

    if (plantGate) plantGate.style.display = 'none';
    if (appShell) appShell.classList.remove('is-hidden');

    startApp();
  }

  async function startApp() {
    await loadMaterialData();
    await loadPartsData();
    loadEmployeeData();

    addNewRow();        // barcode tab
    addSerialRow();     // serial tab แถวแรก
    updateSerialSummary();
  }

  function resetAppState() {
    selectedPlant = '';
    dataUrl = '';
    materialDB = [];
    partsDB = [];
    partsMap = new Map();
    nextIndex = 1;
    serialNextIndex = 1;

    if (tableBody) tableBody.innerHTML = '';
    if (serialTableBody) serialTableBody.innerHTML = '';
    if (rowStatusEl) rowStatusEl.textContent = '';
    if (dataStatusEl) dataStatusEl.textContent = '';
    if (empCodeInput) empCodeInput.value = '';
    if (empStatus) {
      empStatus.textContent = 'กำลังโหลดข้อมูลพนักงาน...';
    }

    if (serialInvoiceInput) serialInvoiceInput.value = '';
    if (serialMatInput) serialMatInput.value = '';
    if (serialDescInput) serialDescInput.value = '';
    if (serialBinInput) serialBinInput.value = '';
    if (serialPrefixInput) serialPrefixInput.value = '';
    if (serialSuffixInput) serialSuffixInput.value = '';
    if (serialTopStatus) {
      serialTopStatus.textContent = 'กรอก Invoice และ Material เพื่อดึง Description / Storage bin จากฐานข้อมูล';
    }
    if (serialStatusEl) {
      serialStatusEl.textContent = 'แถวแรกถูกสร้างให้แล้ว กรอก Serial แล้วกด Enter เพื่อสร้างแถวถัดไป';
    }
    updateSerialSummary();
  }

  function returnToPlant() {
    resetAppState();
    if (plantSelect) plantSelect.value = '';
    if (badgeText) {
      badgeText.textContent = 'โหมดทำงาน: Barcode';
    }
    if (plantGate) plantGate.style.display = 'flex';
    if (appShell) appShell.classList.add('is-hidden');
  }

  if (appShell) appShell.classList.add('is-hidden');
  if (plantEnterBtn) {
    plantEnterBtn.addEventListener('click', setPlantAndStart);
  }
  if (backToPlantBtn) {
    backToPlantBtn.addEventListener('click', returnToPlant);
  }

  if (loginGate) loginGate.style.display = 'flex';
  if (plantGate) plantGate.style.display = 'none';
  setMenuUserInfo(null);

  if (menuBtn && menuPanel) {
    menuBtn.addEventListener('click', (event) => {
      event.stopPropagation();
      toggleMenuPanel();
    });
    menuPanel.addEventListener('click', (event) => {
      event.stopPropagation();
    });
    document.addEventListener('click', () => {
      toggleMenuPanel(false);
    });
  }
  if (logoutBtn) {
    logoutBtn.addEventListener('click', handleLogout);
  }

  if (loginBtn) {
    loginBtn.addEventListener('click', handleLogin);
  }
  if (loginPassInput) {
    loginPassInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        handleLogin();
      }
    });
  }
  if (loginUserInput) {
    loginUserInput.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        handleLogin();
      }
    });
  }

  loadUserData().then(() => {
    const saved = restoreLogin();
    if (saved && loginUserInput && loginPassInput) {
      loginUserInput.value = saved.id || '';
      loginPassInput.value = saved.password || '';
      if (loginRemember) loginRemember.checked = true;
      handleLogin();
    }
  });



