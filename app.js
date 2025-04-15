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
    workbookA = XLSX.read(e.target.result, { type: 'binary' });
    readerB.readAsBinaryString(fileB);
  };

  readerB.onload = function(e) {
    workbookB = XLSX.read(e.target.result, { type: 'binary' });
    compareWorkbooks();
  };

  readerA.readAsBinaryString(fileA);
}

function compareWorkbooks() {
  const sheetNames = workbookA.SheetNames;
  const diffListDiv = document.getElementById("diffList");
  const resultDiv = document.getElementById("result");
  diffListDiv.innerHTML = "";
  resultDiv.innerHTML = "";

  sheetNames.forEach(sheetName => {
    const sheetA = workbookA.Sheets[sheetName];
    const sheetB = workbookB.Sheets[sheetName];
    if (!sheetB) return;

    const range = XLSX.utils.decode_range(sheetA['!ref'] || sheetB['!ref']);
    const tableA = document.createElement("table");
    const tableB = document.createElement("table");
    const diffLinks = [];

    for (let R = range.s.r; R <= range.e.r; ++R) {
      const rowA = tableA.insertRow();
      const rowB = tableB.insertRow();
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
        const cellA = sheetA[cellRef];
        const cellB = sheetB[cellRef];
        const valA = (cellA && cellA.f) ? null : (cellA ? cellA.v : "");
        const valB = (cellB && cellB.f) ? null : (cellB ? cellB.v : "");
        const tdA = rowA.insertCell();
        const tdB = rowB.insertCell();
        tdA.innerText = valA ?? "";
        tdB.innerText = valB ?? "";
        const cellId = `${sheetName}_${cellRef}`;
        tdA.id = `A_${cellId}`;
        tdB.id = `B_${cellId}`;
        if (valA !== valB) {
          tdA.classList.add("diff");
          tdB.classList.add("diff");
          diffLinks.push(`<tr>
            <td>${sheetName}</td>
            <td>${cellRef}</td>
            <td>${valA}</td>
            <td>${valB}</td>
            <td><a class=\"jump-link\" onclick=\"jumpToCell('${cellId}')\">ジャンプ</a></td>
          </tr>`);
        }
      }
    }

    // シート名ラベル
    const labelA = document.createElement("div");
    labelA.innerHTML = `<h3>${sheetName}（ファイルA）</h3>`;
    const labelB = document.createElement("div");
    labelB.innerHTML = `<h3>${sheetName}（ファイルB）</h3>`;

    const groupA = document.createElement("div");
    groupA.appendChild(labelA);
    groupA.appendChild(tableA);

    const groupB = document.createElement("div");
    groupB.appendChild(labelB);
    groupB.appendChild(tableB);

    resultDiv.appendChild(groupA);
    resultDiv.appendChild(groupB);

    // 差分リンクリスト作成
    if (diffLinks.length > 0) {
      diffListDiv.innerHTML += `<h4>${sheetName} の差分</h4>
      <table>
        <tr><th>シート</th><th>セル</th><th>Aの値</th><th>Bの値</th><th>操作</th></tr>
        ${diffLinks.join("")}
      </table>`;
    }
  });
}

function jumpToCell(cellId) {
  const targetA = document.getElementById(`A_${cellId}`);
  const targetB = document.getElementById(`B_${cellId}`);
  [targetA, targetB].forEach(cell => {
    if (cell) {
      cell.scrollIntoView({ behavior: "smooth", block: "center" });
      cell.classList.add("highlight");
      setTimeout(() => cell.classList.remove("highlight"), 1500);
    }
  });
}
