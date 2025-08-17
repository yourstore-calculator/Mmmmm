# Mmmmm
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8" />
  <title>حاسبة الأسعار المطورة</title>
  
  <!-- **إضافة المكتبات الجديدة اللازمة لإنشاء ملف وورد حقيقي** -->
  <script src="https://unpkg.com/html-docx-js/dist/html-docx.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>

  <style>
    body {
      font-family: 'Segoe UI', Arial, sans-serif;
      background-color: #1e1e1e;
      color: white;
      direction: rtl;
      padding: 30px;
    }
    .logo { text-align: center; margin-bottom: 20px; }
    .logo img { max-height: 100px; }
    select, input {
      padding: 8px; margin: 5px 0; width: 100%; border: none;
      border-radius: 4px; box-sizing: border-box; background-color: #333; color: white;
    }
    input[type="file"] { padding: 3px; }
    input:disabled { background-color: #555; cursor: not-allowed; color: #aaa; }
    .section {
      background-color: #252526; padding: 15px; border-radius: 8px;
      margin-bottom: 20px; border: 1px solid #333;
    }
    .addon-section { background-color: #2a2d2e; padding: 10px; border-radius: 6px; margin-top: 10px; }
    label { font-weight: bold; color: #00e676; }
    button {
      background-color: #007acc; color: white; padding: 12px 22px; border: none;
      margin-top: 10px; border-radius: 5px; cursor: pointer; font-size: 16px;
      transition: background-color 0.3s;
    }
    button:hover { background-color: #005a9e; }
    .btn-secondary { background-color: #555; }
    .btn-secondary:hover { background-color: #777; }
    .btn-danger { background-color: #c0392b; }
    .btn-danger:hover { background-color: #a52f22; }
    .results { background: #2b2b2b; padding: 20px; margin-top: 20px; border-radius: 8px; }
    .result-item {
      border-bottom: 1px solid #555; padding: 15px 0; line-height: 1.9;
    }
    .result-item:last-child { border-bottom: none; }
    .result-item h3 { color: #00e676; margin-top: 0; font-size: 1.4em; }
    .result-item ul { list-style-type: disc; padding-right: 20px; }
    .summary {
      font-weight: bold; color: #00e676; padding-top: 15px; border-top: 2px solid #00e676;
      margin-top: 15px; font-size: 1.2em; line-height: 2;
    }
    .final-total {
      background-color: #00e676; color: #1e1e1e; padding: 10px; border-radius: 5px;
      font-size: 1.5em; text-align: center; margin-top: 10px;
    }
  </style>
</head>
<body>

<div class="logo">
  <img src="https://i.imgur.com/ih5zjVj.png" alt="شعار الشركة">
</div>

<div class="section">
    <label for="customerName">اسم الزبون:</label>
    <input type="text" id="customerName" placeholder="مثال: خلف الفطيسي" />
    <label for="customerPhone">رقم هاتف الزبون:</label>
    <input type="text" id="customerPhone" placeholder="مثال: 99455082" />
</div>

<div class="section">
  <label for="mainCategory">الفئة الرئيسية:</label>
  <select id="mainCategory" onchange="loadSubTypes()"></select>
</div>

<div class="section">
  <label for="subType">النوع الفرعي:</label>
  <select id="subType" onchange="updateUIBasedOnType()"></select>
</div>

<div class="section">
  <label for="customItemName">الاسم المخصص للعنصر (اختياري):</label>
  <input type="text" id="customItemName" placeholder="اتركه فارغًا لاستخدام الاسم الافتراضي" />
  <label for="height">الطول (متر):</label>
  <input type="number" id="height" step="0.01" value="1" />
  <label for="width">العرض (متر):</label>
  <input type="number" id="width" step="0.01" value="1" />
  <label for="quantity">الكمية:</label>
  <input type="number" id="quantity" value="1" />
  <label for="styleImage">صورة الستايل:</label>
  <input type="file" id="styleImage" accept="image/*" />
  <div id="openingDirectionContainer" style="display: none; margin-top: 10px;">
    <label for="openingDirection">اتجاه الفتح (RH/LH):</label>
    <select id="openingDirection">
        <option value="left inward">Left Inward (يسار للداخل)</option>
        <option value="right inward">Right Inward (يمين للداخل)</option>
    </select>
  </div>
</div>

<div class="section" id="addonsHost"></div>

<div class="section">
    <label for="installationCost">تكلفة التركيب (اختياري):</label>
    <input type="number" id="installationCost" value="0" step="1" />
</div>

<button onclick="calculate()">➕ أضف للنتائج</button>
<button onclick="clearResults()" class="btn-danger">🗑️ مسح النتائج</button>
<button onclick="saveAsWord()" class="btn-secondary">💾 حفظ كـ Word</button>

<div class="results" id="results"></div>

<script>
const SHIPPING_RATE = 48;
const productData = { "نوافذ": { "دبل جلاس دبل فريم ثابتة": { price: 34, cbm: 0.13, method: 'per_meter', addons: ['curtain', 'net'] }, "دبل جلاس دبل فريم حركة": { price: 73, cbm: 0.13, method: 'per_meter', addons: ['curtain', 'net'] }, "دبل جلاس دبل فريم حركتين": { price: 92, cbm: 0.13, method: 'per_meter', addons: ['curtain', 'net'] }, "دبل جلاس سنجل فريم ثابتة": { price: 26, cbm: 0.07, method: 'per_meter', addons: ['curtain', 'net'] }, "دبل جلاس سنجل فريم حركة": { price: 46, cbm: 0.07, method: 'per_meter', addons: ['curtain', 'net'] }, "دبل جلاس سنجل فريم حركتين": { price: 58, cbm: 0.07, method: 'per_meter', addons: ['curtain', 'net'] }, "سنجل جلاس سنجل فريم ثابتة": { price: 20, cbm: 0.07, method: 'per_meter', addons: ['net'] }, "سنجل جلاس سنجل فريم حركة": { price: 43, cbm: 0.07, method: 'per_meter', addons: ['net'] }, "سنجل جلاس سنجل فريم حركتين": { price: 47, cbm: 0.07, method: 'per_meter', addons: ['net'] }, "نوافذ السلايدنج": { price: 10, cbm: 0.13, method: 'per_meter', special: 'add_10', addons: ['curtain'] }, "النوافذ الكهربائية": { price: 102, cbm: 0.13, method: 'per_meter' }, "سكاي لايت بدون مكينة": { price: 56, cbm: 0.13, method: 'per_meter' }, "سكاي لايت مع مكينة": { price: 145, cbm: 0.13, method: 'per_meter' }, "كارتن وول ثقيل": { price: 56, cbm: 0.15, method: 'per_meter' }, "كارتن وول خفيف": { price: 45, cbm: 0.15, method: 'per_meter' }, }, "أبواب": { "باب المدخل - زينك": { price: 66, cbm: 0.20, method: 'per_meter', special: 'add_10' }, "باب المدخل - ستينلس ستيل": { price: 120, cbm: 0.20, method: 'per_meter', special: 'add_10' }, "باب المدخل - كاست المنيوم": { price: 168, cbm: 0.20, method: 'per_meter', special: 'add_10' }, "WPC - فارغ": { price: 45, cbm: 0.11, method: 'per_unit', std_h: 2.2, std_w: 1.0, addons: ['telbisa'] }, "WPC - مع خشب": { price: 50, cbm: 0.11, method: 'per_unit', std_h: 2.2, std_w: 1.0, addons: ['telbisa'] }, "WPC - مع حشوة ضد الصوت": { price: 60, cbm: 0.11, method: 'per_unit', std_h: 2.2, std_w: 1.0, addons: ['telbisa'] }, "WPC - مع فريم ألمينيوم": { price: 67, cbm: 0.11, method: 'per_unit', std_h: 2.2, std_w: 1.0, addons: ['telbisa'] }, "ألمنيوم - فارغ": { price: 65, cbm: 0.11, method: 'per_unit', std_h: 2.2, std_w: 1.0 }, "ألمنيوم - مع خشب": { price: 75, cbm: 0.11, method: 'per_unit', std_h: 2.2, std_w: 1.0 }, "ألمنيوم - فل": { price: 85, cbm: 0.11, method: 'per_unit', std_h: 2.2, std_w: 1.0 }, "ألمنيوم - باب مخفي": { price: 110, cbm: 0.11, method: 'per_unit', std_h: 2.2, std_w: 1.0 }, "باب صبغ مخفي": { price: 110, cbm: 0.11, method: 'per_unit', std_h: 2.2, std_w: 1.0 }, "ألمنيوم - باب خارجي": { price: 61, cbm: 0.11, method: 'per_unit', std_h: 2.2, std_w: 1.0 }, "دورات مياه - النوع الجديد": { price: 55, cbm: 0.11, method: 'per_unit', std_h: 2.2, std_w: 0.8 }, "دورات مياه - النوع القديم": { price: 45, cbm: 0.11, method: 'per_unit', std_h: 2.2, std_w: 0.8 }, "دورات مياه - مخفي زجاجي": { price: 65, cbm: 0.11, method: 'per_unit', std_h: 2.2, std_w: 0.8 }, }, "أبواب سحب": { "داخلي - زجاج": { price: 38, cbm: 0.15, method: 'per_meter', addons: ['curtain'] }, "داخلي - متين": { price: 41, cbm: 0.15, method: 'per_meter' }, "خارجي - جزء مفتوح": { price: 55, cbm: 0.15, method: 'per_meter' }, "خارجي - جزئين مفتوحات": { price: 58, cbm: 0.15, method: 'per_meter' }, "WPC - سلايد": { price: 61, cbm: 0.15, method: 'per_meter' }, }, "أبواب فولدنج": { "داخلي": { price: 39, cbm: 0.15, method: 'per_meter', addons: ['curtain'] }, "خارجي": { price: 56, cbm: 0.15, method: 'per_meter' }, }, "شتر خارجي": { "شتر رول": { price: 28, cbm: 0.20, method: 'per_meter' } }, "أبواب حدائق": { "كاست المنيوم": { price: 91, cbm: 0.20, method: 'per_meter' } }, "حواجز": { "حاجز درج": { price: 43, cbm: 0.05, method: 'per_meter' }, "حاجز حمام": { price: 32, cbm: 0.05, method: 'per_meter' } } };
const addonPrices = { curtain: 26, net: { 'باب': 39, 'فولدنج': 18, 'سلايد': 14 }, telbisa: 10 };
let resultsList = [];

function toBase64(file) { return new Promise((resolve, reject) => { const reader = new FileReader(); reader.readAsDataURL(file); reader.onload = () => resolve(reader.result); reader.onerror = error => reject(error); }); }
function initializeApp() { const mainCat = document.getElementById("mainCategory"); mainCat.innerHTML = `<option value="">-- اختر الفئة --</option>`; Object.keys(productData).forEach(cat => mainCat.innerHTML += `<option value="${cat}">${cat}</option>`); loadSubTypes(); }
function loadSubTypes() { const mainCatVal = document.getElementById("mainCategory").value; const subType = document.getElementById("subType"); subType.innerHTML = ""; if (mainCatVal && productData[mainCatVal]) { Object.keys(productData[mainCatVal]).forEach(sub => subType.innerHTML += `<option value="${sub}">${sub}</option>`); } updateUIBasedOnType(); }
function updateUIBasedOnType() { const mainCatVal = document.getElementById("mainCategory").value; const openingContainer = document.getElementById('openingDirectionContainer'); const doorCategories = ["أبواب", "أبواب سحب", "أبواب فولدنج", "أبواب حدائق"]; openingContainer.style.display = doorCategories.includes(mainCatVal) ? 'block' : 'none'; if (!mainCatVal || !document.getElementById("subType").value) { document.getElementById("addonsHost").innerHTML = ""; return; } const data = productData[mainCatVal][document.getElementById("subType").value]; const heightInput = document.getElementById("height"), widthInput = document.getElementById("width"); const addonsHost = document.getElementById("addonsHost"); addonsHost.innerHTML = ""; heightInput.disabled = false; widthInput.disabled = false; if (data.method === 'per_unit') { heightInput.value = data.std_h; widthInput.value = data.std_w; } if (data.addons) { if (data.addons.includes('curtain')) addonsHost.innerHTML += `<div class="addon-section"><label><input type="checkbox" id="addCurtain" /> إضافة ستارة داخلية (26 ريال للمتر)</label></div>`; if (data.addons.includes('net')) addonsHost.innerHTML += `<div class="addon-section"><label for="netType">إضافة شبك:</label><select id="netType"><option value="">-- بدون --</option><option value="باب">باب</option><option value="فولدنج">فولدنج</option><option value="سلايد">سلايد</option></select></div>`; if (data.addons.includes('telbisa')) { addonsHost.innerHTML += `<div class="addon-section"><label><input type="checkbox" id="addTelbisa" /> إضافة تلبيسة للسقف</label><input type="number" id="telbisaLength" placeholder="طول التلبيسة بالمتر" style="display:none;" value="1" step="0.1" /></div>`; document.getElementById('addTelbisa').onchange = e => document.getElementById('telbisaLength').style.display = e.target.checked ? 'block' : 'none'; } } }
async function calculate() { const mainCatVal = document.getElementById("mainCategory").value; const subVal = document.getElementById("subType").value; if (!mainCatVal || !subVal) { alert("يرجى اختيار الفئة والنوع."); return; } const data = productData[mainCatVal][subVal]; const qty = parseInt(document.getElementById("quantity").value) || 1; const h = parseFloat(document.getElementById("height").value) || 0; const w = parseFloat(document.getElementById("width").value) || 0; if (h === 0 || w === 0) { alert("الأبعاد يجب أن تكون أكبر من صفر."); return; } const customName = document.getElementById('customItemName').value; const finalItemName = customName.trim() !== '' ? customName : subVal; const openingDir = document.getElementById('openingDirectionContainer').style.display !== 'none' ? document.getElementById("openingDirection").value : ''; const styleImageFile = document.getElementById('styleImage').files[0]; let imageBase64 = ''; if (styleImageFile) { imageBase64 = await toBase64(styleImageFile); } const area = h * w; let basePrice = 0, sizePenalty = 0; if (data.method === 'per_meter') basePrice = data.price * area; else if (data.method === 'per_unit') { basePrice = data.price; const stdArea = data.std_h * data.std_w; if (area > stdArea) sizePenalty = Math.ceil((area - stdArea) / 0.1) * 2; } let itemPrice = basePrice + sizePenalty; let addonsDetailsList = []; if (document.getElementById("addCurtain")?.checked) { const cost = addonPrices.curtain * area; itemPrice += cost; addonsDetailsList.push(`<li>إضافة ستارة</li>`); } const netType = document.getElementById("netType")?.value; if (netType) { const cost = (area * 0.5) * addonPrices.net[netType]; itemPrice += cost; addonsDetailsList.push(`<li>إضافة شبك ${netType}</li>`); } let telbisaLength = 0; if (document.getElementById("addTelbisa")?.checked) { telbisaLength = parseFloat(document.getElementById("telbisaLength").value) || 0; if (telbisaLength > 0) { const cost = Math.ceil(telbisaLength) * addonPrices.telbisa; itemPrice += cost; addonsDetailsList.push(`<li>تلبيسة سقف (${telbisaLength}م)</li>`); } } if (data.special === 'add_10') { itemPrice += 10; addonsDetailsList.push(`<li>إضافة خاصة: 10.00</li>`); } const addonsHTMLForWord = `<li>${finalItemName}</li>` + addonsDetailsList.join(''); let shippingCBM = ((telbisaLength > 0) ? ((h + telbisaLength) * w) : area) * data.cbm; const shippingCost = shippingCBM * SHIPPING_RATE; const totalItemCost = itemPrice + shippingCost; resultsList.push({ displayName: finalItemName, sub: subVal, h, w, qty, openingDirection: openingDir, styleImage: imageBase64, details: { basePrice, sizePenalty, addonsHTML: addonsHTMLForWord, addonsPrice: (itemPrice - basePrice - sizePenalty), shippingPerItem: shippingCost, totalCBM: shippingCBM * qty }, singleItemPrice: totalItemCost, final: totalItemCost * qty }); renderResults(); document.getElementById('styleImage').value = ''; document.getElementById('customItemName').value = ''; }
function renderResults() { const container = document.getElementById("results"); container.innerHTML = ""; let sum = 0; resultsList.forEach((r, index) => { sum += r.final; const d = r.details; container.innerHTML += `<div class="result-item"><h3>${index + 1}. ${r.displayName} (الكمية: ${r.qty})</h3><b>المقاس:</b> ${r.h} × ${r.w} متر<br>${r.openingDirection ? `<b>الاتجاه:</b> ${r.openingDirection}<br>` : ''}<ul><li>السعر الأساسي الإجمالي: ${(d.basePrice * r.qty).toFixed(2)}</li>${d.sizePenalty > 0 ? `<li>إجمالي تكلفة زيادة الحجم: ${(d.sizePenalty * r.qty).toFixed(2)}</li>` : ''}${d.addonsPrice > 0 ? `<li>إجمالي الإضافات: ${(d.addonsPrice * r.qty).toFixed(2)}</li>` : ''}<li>🚚 إجمالي الشحن: ${(d.shippingPerItem * r.qty).toFixed(2)}</li></ul><b>🏷️ سعر الحبة الواحدة: ${r.singleItemPrice.toFixed(2)} ريال</b><br><b>💵 السعر الإجمالي للبند: ${r.final.toFixed(2)} ريال</b></div>`; }); if (resultsList.length > 0) { const comm = sum * 0.04; const installationCost = parseFloat(document.getElementById('installationCost').value) || 0; const total = sum + comm + installationCost; container.innerHTML += `<div class="summary">المجموع الفرعي: ${sum.toFixed(2)} ريال<br>عمولة المكتب (4%): ${comm.toFixed(2)} ريال<br>${installationCost > 0 ? `التركيب: ${installationCost.toFixed(2)} ريال<br>` : ''}<div class="final-total">💰 الناتج النهائي: ${total.toFixed(2)} ريال</div></div>`; } }
function clearResults() { resultsList = []; document.getElementById("results").innerHTML = ""; document.getElementById("mainCategory").value = ""; document.getElementById("height").value = "1"; document.getElementById("width").value = "1"; document.getElementById("quantity").value = "1"; document.getElementById("customerName").value = ""; document.getElementById("customerPhone").value = ""; document.getElementById("installationCost").value = "0"; document.getElementById('styleImage').value = ''; document.getElementById('customItemName').value = ''; loadSubTypes(); }

// --- **الدالة النهائية والمُعدلة لحفظ ملف وورد حقيقي (.docx)** ---
async function saveAsWord() {
    if (resultsList.length === 0) { alert("لا توجد نتائج لحفظها."); return; }
    const customerName = document.getElementById('customerName').value || 'زبون';
    const customerPhone = document.getElementById('customerPhone').value || 'غير محدد';
    const installationCost = parseFloat(document.getElementById('installationCost').value) || 0;

    const headerHtml = `<div style="text-align: center; margin-bottom: 20px;"><div style="font-size: 14px; color: #555;">OMAN - MUSCAT</div><div style="font-size: 24px; font-weight: bold; color: #004a99;">BLUE WAVES SERVICES LLC</div><div style="font-size: 14px; color: #555;">SR. NO. : 1595256</div><div style="font-size: 14px; color: #555;">TEL: 77 22 45 11 – 90 99 88 10</div></div>`;
    const customerHtml = `<div style="background-color: #004a99; color: white; padding: 10px; font-weight: bold; font-size: 16px; margin-bottom: 15px;">${customerName} – ${customerPhone}</div>`;

    let tableHtml = `<table border="1" width="100%" style="border-collapse: collapse; text-align: center; font-size: 12px;"><thead><tr style="background-color: #004a99; color: white;"><th style="padding: 8px; width: 4%;">NO</th><th style="padding: 8px; width: 6%;">H</th><th style="padding: 8px; width: 6%;">W</th><th style="padding: 8px; width: 7%;">m²</th><th style="padding: 8px; width: 4%;">Q</th><th style="padding: 8px; width: 12%;">RH/LH</th><th style="padding: 8px; width: 20%;">STYLE</th><th style="padding: 8px; width: 8%;">PRICE</th><th style="padding: 8px; width: 8%;">TOTAL</th><th style="padding: 8px;">DESCRIPTION</th></tr></thead><tbody>`;
    
    let subtotalFromChina = 0; let totalCBM = 0;
    resultsList.forEach((r, index) => {
        subtotalFromChina += r.final;
        totalCBM += r.details.totalCBM;
        const directionCell = r.openingDirection ? `<div style="padding: 5px; border: 1px solid #ccc; margin: 5px; border-radius: 4px; font-size: 11px;">${r.openingDirection}</div>` : '';
        const descriptionCell = `<ul style="text-align: left; margin: 0; padding-right: 15px; list-style-position: inside;">${r.details.addonsHTML}</ul>`;
        const styleCell = r.styleImage ? `<img src="${r.styleImage}" style="max-width: 150px; height: auto; display: block; margin: 5px auto;">` : '';
        
        tableHtml += `<tr><td style="padding: 8px; font-weight: bold;">${index + 1}</td><td style="padding: 8px;">${r.h.toFixed(2)}</td><td style="padding: 8px;">${r.w.toFixed(2)}</td><td style="padding: 8px;">${(r.h * r.w).toFixed(3)}</td><td style="padding: 8px;">${r.qty}</td><td style="padding: 8px;">${directionCell}</td><td style="padding: 8px; vertical-align: middle;">${styleCell}</td><td style="padding: 8px; background-color: #f2f2f2;">${r.singleItemPrice.toFixed(0)}</td><td style="padding: 8px; background-color: #f2f2f2; font-weight: bold;">${r.final.toFixed(0)}</td><td style="padding: 8px; background-color: #f2f2f2;">${descriptionCell}</td></tr>`;
    });
    tableHtml += `</tbody></table>`;
    
    const commission = subtotalFromChina * 0.04;
    const totalShipping = totalCBM * SHIPPING_RATE;
    const grandTotal = subtotalFromChina + commission + totalShipping + installationCost;

    const summaryTable = `<table width="100%" style="border-collapse: collapse; margin-top: 20px; font-size: 14px; border: 1px solid #333;"><tr style="background-color: #f2f2f2;"><td style="padding: 10px; font-weight: bold; border: 1px solid #ccc;">المجموع الفرعي (سعر الشراء من الصين)</td><td style="padding: 10px; text-align: right; font-weight: bold; border: 1px solid #ccc;">${subtotalFromChina.toFixed(2)} ريال عماني</td></tr><tr><td style="padding: 10px; font-weight: bold; border: 1px solid #ccc;">عمولة المكتب (4%)</td><td style="padding: 10px; text-align: right; font-weight: bold; border: 1px solid #ccc;">${commission.toFixed(2)} ريال عماني</td></tr><tr style="background-color: #f2f2f2;"><td style="padding: 10px; font-weight: bold; border: 1px solid #ccc;">الشحن والتخليص الجمركي (${totalCBM.toFixed(3)} CBM x ${SHIPPING_RATE} ريال)</td><td style="padding: 10px; text-align: right; font-weight: bold; border: 1px solid #ccc;">${totalShipping.toFixed(2)} ريال عماني</td></tr>${installationCost > 0 ? `<tr><td style="padding: 10px; font-weight: bold; border: 1px solid #ccc;">التركيب</td><td style="padding: 10px; text-align: right; font-weight: bold; border: 1px solid #ccc;">${installationCost.toFixed(2)} ريال عماني</td></tr>` : ''}<tr style="background-color: #004a99; color: white;"><td style="padding: 12px; font-weight: bold; font-size: 18px; border: 1px solid #004a99;">الإجمالي النهائي</td><td style="padding: 12px; text-align: right; font-weight: bold; font-size: 18px; border: 1px solid #004a99;">${grandTotal.toFixed(2)} ريال عماني</td></tr></table>`;
    const notesSection = `<div style="margin-top: 25px; font-size: 12px;"><h4 style="margin-bottom: 5px;">ملاحظات:</h4><div style="border: 1px solid #ccc; min-height: 80px; padding: 10px;"></div></div>`;
    
    // **تغليف المحتوى بالكامل داخل هيكل HTML**
    const finalHtml = `<!DOCTYPE html><html dir="rtl"><head><meta charset='utf-8'></head><body style="font-family: Arial, sans-serif;">${headerHtml}${customerHtml}${tableHtml}${summaryTable}${notesSection}</body></html>`;

    // **استخدام المكتبة لتحويل HTML إلى Blob وإنشاء ملف .docx**
    var converted = htmlDocx.asBlob(finalHtml, {
        orientation: 'portrait',
        margins: { top: 720, right: 720, bottom: 720, left: 720 } // هوامش 1 إنش
    });
    
    // **استخدام مكتبة FileSaver لحفظ الملف**
    saveAs(converted, `كوتيشن الفاضل ${customerName}.docx`);
}

initializeApp();
</script>

</body>
</html>
