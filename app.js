if (sessionStorage.getItem('teamcoAuthenticated') !== 'true') {
  location.replace('login.html');
}

let excelFile = null;
const fileInput = document.querySelector('#excelFile');
const fileName = document.querySelector('#fileName');
const button = document.querySelector('#yachtsButton');
const yachtsCheckbox = document.querySelector('#yachtsCheckbox');
const status = document.querySelector('#status');

function updateButtonState() {
  button.disabled = !excelFile || !yachtsCheckbox.checked;
}

fileInput.addEventListener('change', () => {
  excelFile = fileInput.files[0] || null;
  fileName.textContent = excelFile ? excelFile.name : 'Select an .xlsm or .xlsx file';
  updateButtonState();
  status.textContent = '';
});

yachtsCheckbox.addEventListener('change', updateButtonState);

button.addEventListener('click', async () => {
  if (!excelFile) return;
  if (location.protocol === 'file:') {
    status.textContent = 'Open the application from http://localhost to load the template automatically.';
    return;
  }
  button.disabled = true;
    status.textContent = 'Processing Yachts…';
  try {
    const values = await readYachts(excelFile);
    const output = await updatePowerPoint(values);
    download(output, 'Yachts Report Updated.pptx');
    status.textContent = 'Done: the updated PowerPoint was downloaded.';
  } catch (error) {
    console.error(error);
    status.textContent = `Could not generate the PowerPoint: ${error.message}`;
  } finally { updateButtonState(); }
});

async function readYachts(file) {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false });
  const sheet = workbook.Sheets.Yachts;
  if (!sheet) throw new Error('The Excel file does not contain a Yachts sheet.');
  const columns = [3, 6, 9, 12, 15, 18, 21]; // D, G, J, M, P, S, V
  const readRow = (number) => columns.map(column => {
    const cell = sheet[XLSX.utils.encode_cell({ r: number - 1, c: column })];
    const value = cell?.v;
    if (value === undefined || value === null || value === '') return 0;
    const textValue = String(value).trim();
    const hasPercent = textValue.includes('%');
    let n = typeof value === 'number' ? value : Number(textValue.replace('%', '').replace(',', '.'));
    // Excel normalmente entrega 22% como 0.22; si llega como 22 o "22%",
    // convertirlo a la representación decimal que usan las gráficas.
    if (hasPercent || n > 1) n /= 100;
    if (!Number.isFinite(n)) throw new Error(`Non-numeric value in ${XLSX.utils.encode_cell({ r: number - 1, c: column })}.`);
    return n;
  });
  const labels = ['Four Seasons', 'Ritz-Carlton', 'Silversea', 'Regent Seven Seas', 'Seabourn', 'Orient Express', 'Aman'];
  const sortWithLabels = (data) => data
    .map((value, index) => ({ value, label: labels[index] }))
    .sort((a, b) => b.value - a.value);
  const graphic1Original = readRow(5);
  const graphic1 = sortWithLabels(graphic1Original);
  const graphic2 = sortWithLabels(readRow(4));
  const graphic4 = sortWithLabels(readRow(32));
  const graphic5 = sortWithLabels(readRow(31));
  const graphic6 = sortWithLabels(readRow(36));
  const graphic7 = sortWithLabels(readRow(42));
  const graphic8 = sortWithLabels(readRow(41));
  const tableData = {};
  const excelColumns = ['B', 'C', 'D', 'E', 'F', 'G', 'H'];
  const excelPrefixes = ['fs', 'si', 're', 'se', 'ri', 'or', 'am'];
  excelPrefixes.forEach((prefix, columnIndex) => {
    tableData[prefix] = [];
    for (let rowNumber = 9; rowNumber <= 14; rowNumber++) {
      const rawValue = sheet[`${excelColumns[columnIndex]}${rowNumber}`]?.v ?? '';
      const numericValue = Number(rawValue);
      tableData[prefix].push(Number.isFinite(numericValue) ? Math.round(numericValue) : rawValue);
    }
  });
  const sopfData = {};
  const sopfColumns = ['J', 'K', 'L', 'M', 'N', 'O', 'P'];
  excelPrefixes.forEach((prefix, columnIndex) => {
    sopfData[prefix] = [];
    for (let rowNumber = 9; rowNumber <= 14; rowNumber++) {
      const cell = sheet[`${sopfColumns[columnIndex]}${rowNumber}`];
      const rawValue = cell?.v ?? '';
      const numericValue = Number(rawValue);
      sopfData[prefix].push(Number.isFinite(numericValue)
        ? `${Math.round(numericValue <= 1 ? numericValue * 100 : numericValue)}%`
        : (cell?.w ?? rawValue));
    }
  });
  const articleData = {};
  excelPrefixes.forEach((prefix, columnIndex) => {
    const rawValue = sheet[`${excelColumns[columnIndex]}15`]?.v ?? '';
    const numericValue = Number(rawValue);
    articleData[prefix] = Number.isFinite(numericValue) ? Math.round(numericValue) : rawValue;
  });
  const totalData = {};
  const totalColumns = ['J', 'K', 'L', 'M', 'N', 'O', 'P'];
  excelPrefixes.forEach((prefix, columnIndex) => {
    const rawValue = sheet[`${totalColumns[columnIndex]}16`]?.v ?? '';
    const numericValue = Number(rawValue);
    totalData[prefix] = Number.isFinite(numericValue)
      ? Math.round(numericValue <= 1 ? numericValue * 100 : numericValue)
      : rawValue;
  });
  const table2Data = buildTableData(sheet, 20, 26, 27);
  return {
    graphic1: graphic1.map(item => item.value),
    graphic1Original,
    graphic1Labels: graphic1.map(item => item.label),
    graphic2: graphic2.map(item => item.value),
    graphic2Labels: graphic2.map(item => item.label),
    graphic4: graphic4.map(item => item.value),
    graphic4Labels: graphic4.map(item => item.label),
    graphic5: graphic5.map(item => item.value),
    graphic5Labels: graphic5.map(item => item.label),
    graphic6: graphic6.map(item => item.value),
    graphic6Labels: graphic6.map(item => item.label),
    graphic7: graphic7.map(item => item.value),
    graphic7Labels: graphic7.map(item => item.label),
    graphic8: graphic8.map(item => item.value),
    graphic8Labels: graphic8.map(item => item.label),
    tableData,
    sopfData,
    articleData,
    totalData,
    table2Data,
    labels
  };
}

async function updatePowerPoint(values) {
  const response = await fetch('Yachts Report Template.pptx');
  if (!response.ok) throw new Error('Yachts Report Template.pptx was not found next to the application.');
  const templateBuffer = await response.arrayBuffer();
  const zip = await JSZip.loadAsync(templateBuffer);
  const slideFiles = Object.keys(zip.files).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name));
  let targetSlide, slideXml;
  for (const name of slideFiles) {
    const xml = await zip.file(name).async('string');
    if (xml.includes('name="Graphic_1"') && xml.includes('name="Graphic_2"')) { targetSlide = name; slideXml = xml; break; }
  }
  if (!targetSlide) throw new Error('Graphic_1 and Graphic_2 were not found on the same slide.');
  const chartIds = ['Graphic_1', 'Graphic_2'].map(name => {
    const shape = slideXml.match(new RegExp(`name="${name}"[\\s\\S]*?<c:chart[^>]*r:id="([^"]+)"`));
    if (!shape) throw new Error(`${name} was not found.`);
    return shape[1];
  });
  for (let i = 0; i < chartIds.length; i++) {
    const rels = await zip.file(`${targetSlide.replace('ppt/slides/', 'ppt/slides/_rels/')}.rels`).async('string');
    const rel = rels.match(new RegExp(`Id="${chartIds[i]}"[^>]*Target="([^"]+)"`));
    if (!rel) throw new Error(`The link for ${['Graphic_1','Graphic_2'][i]} was not found.`);
    const chartPath = normalizePath('ppt/slides', rel[1]);
    const chartXml = await zip.file(chartPath).async('string');
    const sortedValues = values[i === 0 ? 'graphic1' : 'graphic2'];
    const sortedLabels = values[i === 0 ? 'graphic1Labels' : 'graphic2Labels'];
    zip.file(chartPath, updateChartColors(updateChartCache(chartXml, sortedValues), sortedLabels));
    const chartRels = await zip.file(chartPath.replace('ppt/charts/', 'ppt/charts/_rels/') + '.rels').async('string');
    const embed = chartRels.match(/Type="[^"]*package"[^>]*Target="([^"]+)"/);
    if (!embed) throw new Error('The chart embedded workbook was not found.');
    const embeddedPath = normalizePath('ppt/charts', embed[1]);
    const chartBook = XLSX.read(await zip.file(embeddedPath).async('arraybuffer'), { type: 'array' });
    const chartSheet = chartBook.Sheets[chartBook.SheetNames[0]];
    chartSheet.A1 = { t: 's', v: 'Hotel' };
    chartSheet.B1 = { t: 's', v: 'Yachts' };
    // En la plantilla actual, Graphic_1 y Graphic_2 trazan realmente
    // Sheet1!B2:B8. Se actualiza también ese rango para que la gráfica
    // visible y Edit Data reflejen los valores nuevos.
    ['B2','B3','B4','B5','B6','B7','B8'].forEach((cell, index) => {
      chartSheet[cell] = { t: 'n', v: sortedValues[index], z: '0%' };
    });
    ['A2','A3','A4','A5','A6','A7','A8'].forEach((cell, index) => {
      chartSheet[cell] = { t: 's', v: sortedLabels[index] };
    });
    Object.keys(chartSheet).forEach(cell => {
      if (/^[D-Z](?:[1-9]|[1-7][0-9])$/.test(cell)) delete chartSheet[cell];
    });
    chartSheet['!ref'] = 'A1:B8';
    zip.file(embeddedPath, XLSX.write(chartBook, {
      bookType: 'xlsx',
      type: 'array',
      cellStyles: true
    }));
  }
  // Graphic_3 usa siempre la fila 5, correspondiente al mes anterior.
  for (const name of slideFiles) {
    if (name === targetSlide) continue;
    const xml = await zip.file(name).async('string');
    if (xml.includes('name="Graphic_3"')) {
      await updateGraphic3(zip, name, xml, values.graphic1Original, values.labels);
      break;
    }
  }
  // Reemplazar {{month}} en todas las diapositivas por el mes anterior.
  const previousMonthName = new Intl.DateTimeFormat('en-US', { month: 'long' })
    .format(new Date(new Date().getFullYear(), new Date().getMonth() - 1, 1));
  for (const name of slideFiles) {
    const xml = await zip.file(name).async('string');
    if (xml.includes('{{month}}')) {
      zip.file(name, xml.replace(/\{\{month\}\}/g, previousMonthName));
    }
  }
  let tableSlide = null;
  for (const name of slideFiles) {
    if ((await zip.file(name).async('string')).includes('Table_1')) {
      tableSlide = name;
      break;
    }
  }
  if (tableSlide) {
    await updateTable1(zip, [tableSlide], values.tableData, values.sopfData, values.articleData, values.totalData);
  }
  let table2Slide = null;
  for (const name of slideFiles) {
    if ((await zip.file(name).async('string')).includes('Table_2')) {
      table2Slide = name;
      break;
    }
  }
  if (table2Slide) {
    await updateTable1(zip, [table2Slide], values.table2Data.tableData, values.table2Data.sopfData, values.table2Data.articleData, values.table2Data.totalData);
  }
  for (const graphic of ['Graphic_4', 'Graphic_5', 'Graphic_6', 'Graphic_7', 'Graphic_8']) {
    for (const name of slideFiles) {
      const xml = await zip.file(name).async('string');
      if (xml.includes(`name="${graphic}"`)) {
        const key = graphic.replace('_', '').toLowerCase();
        await updateExtraGraphic(zip, name, xml, graphic, values[key], values[`${key}Labels`]);
        break;
      }
    }
  }
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
}

async function updateExtraGraphic(zip, slideFile, slideXml, graphicName, data, labels) {
  const shape = slideXml.match(new RegExp(`name="${graphicName}"[\\s\\S]*?<c:chart[^>]*r:id="([^"]+)"`));
  if (!shape) throw new Error(`The link for ${graphicName} was not found.`);
  const rels = await zip.file(`${slideFile.replace('ppt/slides/', 'ppt/slides/_rels/')}.rels`).async('string');
  const rel = rels.match(new RegExp(`Id="${shape[1]}"[^>]*Target="([^"]+)"`));
  if (!rel) throw new Error(`The link for ${graphicName} was not found.`);
  const chartPath = normalizePath('ppt/slides', rel[1]);
  zip.file(chartPath, updateChartColors(updateChartCache(await zip.file(chartPath).async('string'), data), labels));
  const chartRels = await zip.file(chartPath.replace('ppt/charts/', 'ppt/charts/_rels/') + '.rels').async('string');
  const embed = chartRels.match(/Type="[^"]*package"[^>]*Target="([^"]+)"/);
  if (!embed) throw new Error(`The embedded workbook for ${graphicName} was not found.`);
  const embeddedPath = normalizePath('ppt/charts', embed[1]);
  const book = XLSX.read(await zip.file(embeddedPath).async('arraybuffer'), { type: 'array' });
  const sheet = book.Sheets[book.SheetNames[0]];
  ['D2','G2','J2','M2','P2','S2','V2'].forEach((cell, i) => { sheet[cell] = { t: 'n', v: data[i], z: '0%' }; });
  ['B2','B3','B4','B5','B6','B7','B8'].forEach((cell, i) => { sheet[cell] = { t: 'n', v: data[i], z: '0%' }; });
  ['A2','A3','A4','A5','A6','A7','A8'].forEach((cell, i) => { sheet[cell] = { t: 's', v: labels[i] }; });
  zip.file(embeddedPath, XLSX.write(book, { bookType: 'xlsx', type: 'array', cellStyles: true }));
}

function buildTableData(sheet, firstRow, articleRow, totalRow) {
  const excelColumns = ['B', 'C', 'D', 'E', 'F', 'G', 'H'];
  const excelPrefixes = ['fs', 'si', 're', 'se', 'ri', 'or', 'am'];
  const tableData = {}, sopfData = {}, articleData = {}, totalData = {};
  const sopfColumns = ['J', 'K', 'L', 'M', 'N', 'O', 'P'];
  excelPrefixes.forEach((prefix, i) => {
    tableData[prefix] = [];
    sopfData[prefix] = [];
    for (let r = firstRow; r < firstRow + 6; r++) {
      tableData[prefix].push(Math.round(Number(sheet[`${excelColumns[i]}${r}`]?.v ?? 0)));
      const value = Number(sheet[`${sopfColumns[i]}${r}`]?.v ?? 0);
      sopfData[prefix].push(`${Math.round(value <= 1 ? value * 100 : value)}%`);
    }
    const article = Number(sheet[`${excelColumns[i]}${articleRow}`]?.v ?? 0);
    articleData[prefix] = Number.isFinite(article) ? Math.round(article) : '';
    const total = Number(sheet[`${sopfColumns[i]}${totalRow}`]?.v ?? 0);
    totalData[prefix] = Number.isFinite(total) ? Math.round(total <= 1 ? total * 100 : total) : '';
  });
  return { tableData, sopfData, articleData, totalData };
}

async function updateTable1(zip, slideFiles, values, sopfValues, articleValues, totalValues) {
  const powerPointPrefixes = ['fs', 're', 'si', 'ri', 'se', 'am', 'or'];
  const rowNames = ['he', 'ex', 'le', 'ex', 'vi', 'po'];
  for (const name of slideFiles) {
    let xml = await zip.file(name).async('string');
    for (const prefix of powerPointPrefixes) {
      for (let rowIndex = 0; rowIndex < rowNames.length; rowIndex++) {
        const volumePlaceholder = `${prefix}_${rowNames[rowIndex]}_vol`;
        const sopfPlaceholder = `${prefix}_${rowNames[rowIndex]}_sopf`;
        xml = replacePlaceholder(xml, volumePlaceholder, values[prefix]?.[rowIndex] ?? '');
        xml = replacePlaceholder(xml, sopfPlaceholder, sopfValues[prefix]?.[rowIndex] ?? '');
      }
      xml = replacePlaceholder(xml, `art_${prefix}`, articleValues[prefix] ?? '');
      xml = replacePlaceholder(xml, `${prefix}_total`, totalValues[prefix] ?? '');
    }
    zip.file(name, xml);
  }
}

function replacePlaceholder(xml, placeholder, value) {
  const escaped = placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const token = new RegExp(`\\{\\{(?:[^{}<]|<[^>]+>)*?${escaped}(?:[^{}<]|<[^>]+>)*?\\}\\}`);
  return xml.replace(token, String(value));
}

async function updateGraphic3(zip, slideFile, slideXml, data, labels) {
  // Graphic_3 usa el orden de columnas: D, G, P, J, M, S, V.
  const monthlyData = [data[0], data[1], data[4], data[2], data[3], data[5], data[6]];
  const shape = slideXml.match(/name="Graphic_3"[\s\S]*?<c:chart[^>]*r:id="([^"]+)"/);
  if (!shape) throw new Error('The link for Graphic_3 was not found.');
  const rels = await zip.file(`${slideFile.replace('ppt/slides/', 'ppt/slides/_rels/')}.rels`).async('string');
  const rel = rels.match(new RegExp(`Id="${shape[1]}"[^>]*Target="([^"]+)"`));
  if (!rel) throw new Error('The link for Graphic_3 was not found.');
  const chartPath = normalizePath('ppt/slides', rel[1]);
  const previousMonth = new Date();
  previousMonth.setMonth(previousMonth.getMonth() - 1);
  const monthIndex = previousMonth.getMonth();
  const monthRow = monthIndex + 2; // January is row 2
  const chartXml = updateMonthlyChartCache(
    await zip.file(chartPath).async('string'), monthlyData, monthIndex
  );
  zip.file(chartPath, updateMonthlyDataLabels(chartXml, monthIndex));
  const chartRels = await zip.file(chartPath.replace('ppt/charts/', 'ppt/charts/_rels/') + '.rels').async('string');
  const embed = chartRels.match(/Type="[^"]*package"[^>]*Target="([^"]+)"/);
  if (!embed) throw new Error('The embedded workbook for Graphic_3 was not found.');
  const embeddedPath = normalizePath('ppt/charts', embed[1]);
  const book = XLSX.read(await zip.file(embeddedPath).async('arraybuffer'), { type: 'array' });
  const sheet = book.Sheets[book.SheetNames[0]];
  ['Four Seasons','Ritz-Carlton','Seabourn','Silversea','Regent Seven Seas','Orient Express','Aman']
    .forEach((label, i) => { sheet[XLSX.utils.encode_cell({ r: 0, c: i + 1 })] = { t: 's', v: label }; });
  monthlyData.forEach((value, i) => {
    sheet[XLSX.utils.encode_cell({ r: monthRow - 1, c: i + 1 })] = { t: 'n', v: value, z: '0%' };
  });
  // Conservar el formato porcentual de toda la tabla histórica.
  for (let r = 1; r <= 12; r++) {
    for (let c = 1; c <= 7; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      if (cell && typeof cell.v === 'number') cell.z = '0.##%';
    }
  }
  zip.file(embeddedPath, XLSX.write(book, { bookType: 'xlsx', type: 'array', cellStyles: true }));
}

function updateMonthlyChartCache(xml, data, monthIndex) {
  xml = xml.replace(/<c:numFmt\s+formatCode="[^"]*"/g, '<c:numFmt formatCode="0%"');
  // Solicitar a PowerPoint que recargue la caché al abrir la presentación.
  if (xml.includes('<c:externalData')) {
    xml = xml.replace(/<c:autoUpdate[^>]*\/>/g, '<c:autoUpdate val="1"/>');
  }
  let seriesIndex = 0;
  return xml.replace(/<c:numCache>([\s\S]*?)<\/c:numCache>/g, (whole, cache) => {
    const value = data[seriesIndex++];
    const pointPattern = new RegExp(`(<c:pt\\s+idx="${monthIndex}"[^>]*>\\s*<c:v>)[^<]*(<\\/c:v>)`);
    let updated = cache.replace(pointPattern, `$1${value}$2`);
    if (updated === cache && value !== undefined) {
      updated = updated.replace(/<c:ptCount\s+val="(\d+)"\s*\/>/, `<c:ptCount val="12"/>`);
      updated += `<c:pt idx="${monthIndex}"><c:v>${value}</c:v></c:pt>`;
    }
    return '<c:numCache>' + updated + '</c:numCache>';
  });
}

const HOTEL_COLORS = {
  'Four Seasons': '3D441E',
  'Seabourn': '5F88BD',
  'Silversea': '939393',
  'Regent Seven Seas': '1F216C',
  'Ritz-Carlton': '1F335B',
  'Aman': 'DDC38B',
  'Orient Express': '4F0304'
};

function updateChartColors(xml, labels) {
  if (!labels?.length) return xml;

  // Column/bar charts store each category color in a c:dPt block.
  xml = xml.replace(/<c:dPt>([\s\S]*?)<\/c:dPt>/g, (whole, point) => {
    const index = point.match(/<c:idx\s+val="(\d+)"/);
    const color = index && HOTEL_COLORS[labels[Number(index[1])]];
    if (!color) return whole;
    return whole.replace(/(<c:spPr>[\s\S]*?<a:solidFill><a:srgbClr\s+val=")[^"]+/, `$1${color}`);
  });

  // Line/area charts store the color on each series instead.
  return xml.replace(/<c:ser>([\s\S]*?)<\/c:ser>/g, (whole, series) => {
    const index = series.match(/<c:idx\s+val="(\d+)"/);
    const color = index && HOTEL_COLORS[labels[Number(index[1])]];
    if (!color) return whole;
    return whole.replace(/(<c:spPr>[\s\S]*?<a:solidFill><a:srgbClr\s+val=")[^"]+/, `$1${color}`);
  });
}

function updateMonthlyDataLabels(xml, activeMonthIndex) {
  return xml.replace(/<c:dLbls>([\s\S]*?)<\/c:dLbls>/g, (whole, labelsXml) => {
    const updatedLabels = labelsXml.replace(/<c:dLbl>([\s\S]*?)<\/c:dLbl>/g, (labelXml, label) => {
      const indexMatch = label.match(/<c:idx\s+val="(\d+)"/);
      if (!indexMatch) return labelXml;
      const isActive = Number(indexMatch[1]) === activeMonthIndex;
      let updated = label.replace(/<c:delete\s+val="1"\s*\/>/g, '');
      if (!isActive) updated = updated.replace(/(<c:idx\s+val="\d+"\s*\/>)/, '$1<c:delete val="1"/>');
      return `<c:dLbl>${updated}</c:dLbl>`;
    });
    return `<c:dLbls>${updatedLabels}</c:dLbls>`;
  });
}

function updateChartCache(xml, data) {
  xml = xml.replace(/<c:numFmt\s+formatCode="[^"]*"/g, '<c:numFmt formatCode="0%"');
  let used = false;
  return xml.replace(/<c:numCache>([\s\S]*?)<\/c:numCache>/g, (whole, cache) => {
    let index = 0;
    const updated = cache.replace(/(<c:pt\s+idx="\d+"[^>]*>\s*<c:v>)[^<]*(<\/c:v>)/g, (point, open, close) => {
      if (index >= data.length) return point;
      used = true;
      return open + String(data[index++]) + close;
    });
    let result = updated;
    // Si la plantilla no tiene todos los puntos, agregarlos a la caché.
    for (let pointIndex = index; pointIndex < data.length; pointIndex++) {
      result += `<c:pt idx="${pointIndex}"><c:v>${data[pointIndex]}</c:v></c:pt>`;
    }
    result = result.replace(/<c:ptCount\s+val="\d+"\s*\/>/, `<c:ptCount val="${data.length}"/>`);
    return '<c:numCache>' + result + '</c:numCache>';
  });
}

function normalizePath(base, target) {
  const parts = `${base}/${target}`.split('/'); const out = [];
  for (const part of parts) { if (part === '..') out.pop(); else if (part !== '.') out.push(part); }
  return out.join('/');
}
function download(blob, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 1000); }
