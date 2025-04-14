pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://unpkg.com/pdfjs-dist@3.4.120/build/pdf.worker.min.js';

let excelData = [];
let pdfData = [];

// Excel読み込み
document.getElementById('excel-upload').addEventListener('change', function(e) {
  const reader = new FileReader();
  reader.onload = (evt) => {
    const workbook = XLSX.read(evt.target.result, { type: 'binary' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    excelData = raw.slice(1).map(row => {
      const [month, day, jp, en, source] = row;
      if (!jp || !en || !source) return null;
      return {
        date: `${month}月${day}日`,
        jp: jp.toString().trim(),
        en: en.toString().trim(),
        source: source.toString().trim()
      };
    }).filter(x => x);
  };
  reader.readAsBinaryString(e.target.files[0]);
});

// PDF読み込み
document.getElementById('pdf-upload').addEventListener('change', async function(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = async function () {
    const typedArray = new Uint8Array(this.result);
    const pdf = await pdfjsLib.getDocument({ data: typedArray }).promise;
    const allLines = [];

    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const lines = content.items.map(item => item.str.trim()).filter(t => t.length > 0);
      allLines.push(...lines);
    }

    pdfData = extractSmartQuotes(allLines);
  };
  reader.readAsArrayBuffer(file);
});

// PDFの名言抽出（2日分セット）
function extractSmartQuotes(lines) {
  const results = [];
  let day = 1;

  for (let i = 6; i < lines.length; i++) {
    const year = lines[i - 1];
    const serial = lines[i];

    if (year === '2026' && /^\d{3}$/.test(serial)) {
      const window = lines.slice(i - 6, i - 1);
      let person = '', info = '', en = '', ja = '', matched = false;

      for (let j = 0; j <= 2; j++) {
        const maybePerson = window[j];
        const maybeInfo = window[j + 1];
        const maybeEn = window[j + 2];
        const maybeJa1 = window[j + 3];
        const maybeJa2 = window[j + 4];

        const isValidInfo = /^[（(][0-9B.C.～年・）\s\-～]+/.test(maybeInfo);
        const isEn = /[a-zA-Z]/.test(maybeEn);
        const isJa = /[ぁ-んァ-ン一-龯]/.test(maybeJa1 + maybeJa2);

        if (isValidInfo && isEn && isJa) {
          person = maybePerson;
          info = maybeInfo;
          en = maybeEn;
          ja = (maybeJa1 + ' ' + maybeJa2).replace(/\s+/g, ' ').trim();
          matched = true;
          break;
        }
      }

      if (matched) {
        results.push({ date: `1月${day}日`, ja, en, source: `${person}${info}` });
      } else {
        results.push({ date: `1月${day}日`, ja: 'error', en: 'error', source: 'error' });
      }

      day++;
    }
  }

  return results;
}

// 比較処理（DOMContentLoaded後にボタンに紐付け）
document.addEventListener("DOMContentLoaded", () => {
  document.getElementById("compare-btn").addEventListener("click", compareQuotes);
});

function compareQuotes() {
  const dmp = new diff_match_patch();
  const table = document.createElement('table');
  table.innerHTML = '<tr><th>日付</th><th>項目</th><th>Excel</th><th>PDF</th><th>判定</th></tr>';

  for (let i = 0; i < excelData.length; i++) {
    const ex = excelData[i];
    const pdf = pdfData[i];
    if (!pdf) {
      ['jp', 'en', 'source'].forEach(type => {
        const row = table.insertRow();
        row.innerHTML = `<td>${ex.date}</td><td>${label(type)}</td><td>${ex[type]}</td><td class="error">データなし</td><td class="error">エラー</td>`;
      });
      continue;
    }

    ['jp', 'en', 'source'].forEach(type => {
      const row = table.insertRow();
      const diffs = dmp.diff_main(ex[type], pdf[type]);
      dmp.diff_cleanupSemantic(diffs);
      const html = dmp.diff_prettyHtml(diffs);
      const match = ex[type] === pdf[type];
      row.innerHTML = `<td>${ex.date}</td><td>${label(type)}</td><td>${ex[type]}</td><td class="${match ? '' : 'mismatch'}">${match ? pdf[type] : html}</td><td>${match ? '一致' : '不一致'}</td>`;
    });
  }

  document.getElementById('output').innerHTML = '';
  document.getElementById('output').appendChild(table);
}

function label(type) {
  return type === 'jp' ? '日本語' : type === 'en' ? '英語' : '出典';
}
