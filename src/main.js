const XLSX = require('xlsx');
const fs = require('fs');

const file1Input = document.getElementById('file1');
const file2Input = document.getElementById('file2');
const file1Name = document.getElementById('file1Name');
const file2Name = document.getElementById('file2Name');
const compareForm = document.getElementById('compareForm');
const resultMessage = document.getElementById('resultMessage');

file1Input.addEventListener('change', function () {
  file1Name.textContent = file1Input.files[0]
    ? file1Input.files[0].name
    : 'Файл не вибрано';
});

file2Input.addEventListener('change', function () {
  file2Name.textContent = file2Input.files[0]
    ? file2Input.files[0].name
    : 'Файл не вибрано';
});
compareForm.addEventListener('submit', function (e) {
  e.preventDefault();
  resultMessage.textContent = 'Порівняння триває...';
  // Відкриваємо два файли Excel
  const workbook1 = XLSX.readFile(file1Name);
  const workbook2 = XLSX.readFile(file2Name);

  // Беремо перший лист з кожного файлу
  const sheet1 = workbook1.Sheets[workbook1.SheetNames[0]];
  const sheet2 = workbook2.Sheets[workbook2.SheetNames[0]];

  // Перетворюємо листи в JSON формат
  const data1 = XLSX.utils.sheet_to_json(sheet1, { header: 1 });
  const data2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 });

  const result = [];

  // Порівнюємо вміст ячейок стовпця F у першому файлі зі стовпцем D у другому файлі
  data1.forEach((row1) => {
    const cell1 = row1[5]; // Значення з стовпця F
    if (!cell1) return; // Пропускаємо порожні ячейки у першому файлі

    data2.forEach((row2) => {
      const cell2 = row2[3]; // Значення з стовпця D
      if (!cell2) return; // Пропускаємо порожні ячейки у другому файлі

      if (cell1 === cell2) {
        result.push(row1);
      }
    });
  });

  // Створюємо новий Excel файл з результатами
  const newWorkbook = XLSX.utils.book_new();
  const newSheet = XLSX.utils.aoa_to_sheet(result);

  XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Results');

  // Зберігаємо новий файл
  XLSX.writeFile(newWorkbook, 'result.xlsx');

  resultMessage.textContent(
    'Порівняння завершено. Результати збережено в result.xlsx у папці програми'
  );
});
