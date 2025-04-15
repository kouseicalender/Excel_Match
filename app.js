// app.js（軽量版・差分一覧とシート追加/削除のみ）
let workbookA, workbookB;

function compareFiles() {
  const fileA = document.getElementById("fileA").files[0];
  const fileB = document.getElementById("fileB").files[0];
  if (!fileA || !fileB) {
    alert("両方のファイルを選択してください。");
    return;
  }

  const readerA = new FileReader();
  const readerB = new FileReader();

  readerA.onload = function(e) {
    workbookA = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
    readerB.readAsBinaryString(fileB);
  };

  readerB.onload = function(e) {
    workbookB = XLSX.read(e.target.result, { type: 'binary', cellDates: true });
    compareWorkbooks();
  };

  readerA.readAsBinaryString(fileA);
}

function formatValue(val, cell) {
  if (cell && cell.t === 'd' && val instanceof Date) {
    return val.toLocaleDateString('ja-JP');
  }
  return (val ?? "").toString().trim();
}

function compareWorkbooks() {
  const sheetNamesA = workbookA.SheetNames;
  const sheetNamesB = workbookB.SheetNames;
  const diffListDiv = document.getElementById("diffList");
  diffListDiv.innerHTML = "";

  const deletedSheets = sheetNamesA.filter(name => !sheetNamesB.includes(name));
  const addedSheets = sheetNamesB.filter(name => !sheetNamesA.includes(name));

  if (addedSheets.length > 0) {
    diffListDiv.innerHTML += `<h3>追加されたシート</h3><ul>${addedSheets.map(name => `<li>${name}</li>`).join('')}</ul>`;
  }
  if (deletedSheets.length > 0) {
    diffListDiv.innerHTML += `<h3>削除されたシート</h3><ul>${deletedSheets.map(name => `<li>${name}</li>`).join('')}</ul>`;
  }

  const commonSheets = sheetNamesA.filter(name => sheetNamesB.includes(name));

  commonSheets.forEach(sheetName => {
    const sheetA = workbookA.Sheets[sheetName];
    const sheetB = workbookB.Sheets[sheetName];
    const range = XLSX.utils.decode_range(sheetA['!ref'] || sheetB['!ref']);

    const diffRows = [];
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
        const cellA = sheetA[cellRef];
        const cellB = sheetB[cellRef];
        const rawA = (cellA && cellA.f) ? null : (cellA ? cellA.v : "");
        const rawB = (cellB && cellB.f) ? null : (cellB ? cellB.v : "");
        const valA = formatValue(rawA, cellA);
        const valB = formatValue(rawB, cellB);
        if (valA !== valB) {
          diffRows.push(`<tr>
            <td>${sheetName}</td>
            <td>${cellRef}</td>
            <td>${valA}</td>
            <td>${valB}</td>
          </tr>`);
        }
      }
    }

    if (diffRows.length > 0) {
      diffListDiv.innerHTML += `<h4>${sheetName} の差分</h4>
        <table>
          <tr><th>シート</th><th>セル</th><th>Aの値</th><th>Bの値</th></tr>
          ${diffRows.join("")}
        </table>`;
    }
  });
}
