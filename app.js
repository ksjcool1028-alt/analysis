/**
 * ==========================================================================
 * 자산관리 다차원 리포트 시스템 (app.js)
 * 1. 카드 제목: '가족 통합 포트폴리오 비중 (X축 90도 수직 표기)'
 * 2. '평가금액' tab: X축과 Y축 상호 변경 (indexAxis: 'y') -> Y축 항목명, X축 수치 ₩ 배치!
 * 3. X축 눈금 글자 90도 직각 수직 표기 (maxRotation: 90, minRotation: 90)
 * 4. 막대 우측 [평가액 (비중%) | 수익률%] 3가지 정보 노출 & 상하 스크롤 연동
 * 5. 2단계 동적 피벗 컨트롤 (Row 컬럼 다중 선택 + Row내 세부 Value 선택)
 * 6. 피벗 집계 결과표 오름차순/내림차순 컬럼 정렬 & F5 새로고침 데이터 보존
 * ==========================================================================
 */

const AppState = {
  rawDataset: [],
  // 엑셀 업로드 시 감지될 8개 다중 필터 대상 컬럼 명칭 (유동적 변경 반영)
  filterKeys: ['운용사', '성명', '구분', '구분2', '구분3', '국내외', '국가_상품', '투자상품'],

  // 시스템 내부 키 ↔ 업로드된 실제 Excel 컬럼명 동적 매핑 테이블
  // 엑셀 컬럼명이 변경될 경우 바뀐 헤더명을 기록하여 화면 표시에 유동 적용하기 위함
  columnMap: {
    평가일: '평가일',
    운용사: '운용사',
    성명: '성명',
    구분: '구분',
    구분2: '구분2',
    구분3: '구분3',
    국내외: '국내외',
    국가_상품: '국가_상품',
    투자상품: '투자상품',
    투자원금: '투자원금',
    평가금액: '평가금액',
    증감: '증감',
    수익률: '수익률',
    환매여부: '환매여부',
    최종평가금액: '최종평가금액',
    pick: 'pick',
    비고: '비고'
  },

  sec1: { selectedDate: 'LATEST', filters: {} },
  sec2: { endDate: 'LATEST', metric: '평가금액', filters: {} },
  sec3: { selectedDate: 'LATEST', filters: {} },

  pivotConfig: {
    rows: ['운용사'],
    rowValues: {},
    metric: '평가금액'
  },
  pivotSortState: {
    column: 'valuation',
    direction: 'desc'
  },

  chartTab: 'donut',
  sortState: { column: '평가일', direction: 'desc' },
  chartInstances: {
    sec2TimeSeries: null,
    familyAllocation: null,
    personCharts: {}
  }
};

const $ = (id) => document.getElementById(id);

document.addEventListener('DOMContentLoaded', () => {
  initEventListeners();
  if (!loadFromLocalStorage()) {
    loadSampleData();
  }
});

/* ==========================================================================
   1. 이벤트 리스너 & 피벗 정렬
   ========================================================================== */
function initEventListeners() {
  $('btnLoadSample').addEventListener('click', () => loadSampleData());
  $('toastCloseBtn').addEventListener('click', hideToast);

  const fileInput = $('fileInput');
  const dropZone = $('dropZone');

  dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
  dropZone.addEventListener('dragleave', () => { dropZone.classList.remove('dragover'); });
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) handleFileUpload(e.dataTransfer.files[0]);
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFileUpload(e.target.files[0]);
      e.target.value = ''; // ★ 동일 파일명/데이터 수정 재업로드 시 change 이벤트 발화를 위해 내부 값 리셋!
    }
  });

  fileInput.addEventListener('click', (e) => {
    e.target.value = ''; // ★ 클릭 시에도 이전 파일 경로 리셋
  });

  document.addEventListener('click', (e) => {
    if (!e.target.closest('.dropdown-multi')) {
      document.querySelectorAll('.multi-select-dropdown').forEach(el => el.classList.add('hidden'));
    }
  });

  // Section 1
  $('sec1DateSelect').addEventListener('change', (e) => { AppState.sec1.selectedDate = e.target.value; updateSection1(); });

  // Section 2
  $('sec2EndDateSelect').addEventListener('change', (e) => { AppState.sec2.endDate = e.target.value; updateSection2(); });
  $('sec2MetricSelect').addEventListener('change', (e) => { AppState.sec2.metric = e.target.value; updateSection2(); });

  // Section 3
  $('sec3DateSelect').addEventListener('change', (e) => { AppState.sec3.selectedDate = e.target.value; updateSection3(); });

  // Section 3 피벗 집계 항목 (Value) 연동
  $('pivotMetricSelect').addEventListener('change', (e) => {
    AppState.pivotConfig.metric = e.target.value;
    updateSection3ChartsAndPivot();
  });

  // 피벗 테이블 헤더 정렬
  document.querySelectorAll('#pivotDataTable th[data-pivot-sort]').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.getAttribute('data-pivot-sort');
      if (AppState.pivotSortState.column === col) {
        AppState.pivotSortState.direction = AppState.pivotSortState.direction === 'asc' ? 'desc' : 'asc';
      } else {
        AppState.pivotSortState.column = col;
        AppState.pivotSortState.direction = 'desc';
      }
      renderPivotTable();
    });
  });

  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      e.target.classList.add('active');
      AppState.chartTab = e.target.getAttribute('data-tab');
      renderFamilyAllocationChart();
    });
  });

  $('tableSearchInput').addEventListener('input', () => renderRawTable());

  // Raw Data 테이블 정렬
  document.querySelectorAll('#rawDataTable th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.getAttribute('data-sort');
      if (AppState.sortState.column === col) {
        AppState.sortState.direction = AppState.sortState.direction === 'asc' ? 'desc' : 'asc';
      } else {
        AppState.sortState.column = col;
        AppState.sortState.direction = 'desc';
      }
      renderRawTable();
    });
  });

  $('btnExportExcel').addEventListener('click', exportToExcel);
}

/* ==========================================================================
   2단계 동적 피벗 컨트롤
   ========================================================================== */

function initPivotRowMultiSelect() {
  const container = $('pivotRowMultiContainer');
  container.innerHTML = '';

  const items = AppState.filterKeys;

  const box = document.createElement('div');
  box.className = 'multi-select-box';
  box.id = 'pivotRowMultiBox';
  box.innerHTML = `<span class="multi-select-label">그룹 컬럼</span><i class="fa-solid fa-chevron-down"></i>`;

  const dropdown = document.createElement('div');
  dropdown.className = 'multi-select-dropdown hidden';
  dropdown.id = 'pivotRowMultiDropdown';

  items.forEach(item => {
    const lbl = document.createElement('label');
    lbl.className = 'multi-option';
    const isChecked = AppState.pivotConfig.rows.includes(item) ? 'checked' : '';
    lbl.innerHTML = `<input type="checkbox" value="${item}" ${isChecked}> ${item}`;
    dropdown.appendChild(lbl);
  });

  container.appendChild(box);
  container.appendChild(dropdown);

  box.addEventListener('click', (e) => {
    e.stopPropagation();
    const isHidden = dropdown.classList.contains('hidden');
    document.querySelectorAll('.multi-select-dropdown').forEach(el => el.classList.add('hidden'));
    if (isHidden) dropdown.classList.remove('hidden');
  });

  const checkboxes = dropdown.querySelectorAll('input[type="checkbox"]');
  checkboxes.forEach(cb => {
    cb.addEventListener('change', () => {
      const selected = Array.from(checkboxes).filter(c => c.checked).map(c => c.value);
      if (selected.length === 0) {
        cb.checked = true;
        return;
      }
      AppState.pivotConfig.rows = selected;
      updateMultiSelectBoxLabel(box, '그룹 컬럼', selected, items.length);

      updatePivotRowValuesSubFilters();
      updateSection3ChartsAndPivot();
    });
  });

  updateMultiSelectBoxLabel(box, '그룹 컬럼', AppState.pivotConfig.rows, items.length);
  updatePivotRowValuesSubFilters();
}

function updatePivotRowValuesSubFilters() {
  const container = $('pivotRowValuesContainer');
  container.innerHTML = '';

  const selectedRows = AppState.pivotConfig.rows;

  selectedRows.forEach(rowKey => {
    const uniqueValues = getUniqueList(rowKey);

    if (!AppState.pivotConfig.rowValues[rowKey]) {
      AppState.pivotConfig.rowValues[rowKey] = [...uniqueValues];
    } else {
      uniqueValues.forEach(uVal => {
        if (!AppState.pivotConfig.rowValues[rowKey].includes(uVal)) {
          AppState.pivotConfig.rowValues[rowKey].push(uVal);
        }
      });
    }

    const groupDiv = document.createElement('div');
    groupDiv.className = 'filter-group dropdown-multi';

    const label = document.createElement('label');
    label.className = 'text-xs text-muted';
    label.textContent = `[${rowKey}] 내 세부 Value`;

    const box = document.createElement('div');
    box.className = 'multi-select-box';
    box.id = `pivotVal_${rowKey}_box`;
    box.innerHTML = `<span class="multi-select-label">전체 ${rowKey}</span><i class="fa-solid fa-chevron-down"></i>`;

    const dropdown = document.createElement('div');
    dropdown.className = 'multi-select-dropdown hidden';

    groupDiv.appendChild(label);
    groupDiv.appendChild(box);
    groupDiv.appendChild(dropdown);
    container.appendChild(groupDiv);

    box.addEventListener('click', (e) => {
      e.stopPropagation();
      const isHidden = dropdown.classList.contains('hidden');
      document.querySelectorAll('.multi-select-dropdown').forEach(el => el.classList.add('hidden'));
      if (isHidden) dropdown.classList.remove('hidden');
    });

    populatePivotRowValueOptions(dropdown, box, rowKey, uniqueValues);
  });
}

function populatePivotRowValueOptions(dropdown, box, rowKey, values) {
  dropdown.innerHTML = '';

  const isAllChecked = values.length > 0 && AppState.pivotConfig.rowValues[rowKey] && values.every(v => AppState.pivotConfig.rowValues[rowKey].includes(v));
  const allLabel = document.createElement('label');
  allLabel.className = 'multi-option font-bold';
  allLabel.innerHTML = `<input type="checkbox" value="ALL" ${isAllChecked ? 'checked' : ''}> 전체 선택 (${values.length}개)`;
  dropdown.appendChild(allLabel);

  values.forEach(val => {
    const lbl = document.createElement('label');
    lbl.className = 'multi-option';
    const isChecked = AppState.pivotConfig.rowValues[rowKey].includes(val) ? 'checked' : '';
    lbl.innerHTML = `<input type="checkbox" value="${val}" ${isChecked}> ${val}`;
    dropdown.appendChild(lbl);
  });

  const checkboxes = dropdown.querySelectorAll('input[type="checkbox"]');
  checkboxes.forEach(cb => {
    cb.addEventListener('change', (e) => {
      const val = e.target.value;
      if (val === 'ALL') {
        checkboxes.forEach(c => c.checked = e.target.checked);
      } else {
        const allCb = dropdown.querySelector('input[value="ALL"]');
        const itemCbs = Array.from(checkboxes).filter(c => c.value !== 'ALL');
        allCb.checked = itemCbs.every(c => c.checked);
      }

      const selected = Array.from(checkboxes)
        .filter(c => c.checked && c.value !== 'ALL')
        .map(c => c.value);

      AppState.pivotConfig.rowValues[rowKey] = selected;
      updateMultiSelectBoxLabel(box, rowKey, selected, values.length);
      updateSection3ChartsAndPivot();
    });
  });

  updateMultiSelectBoxLabel(box, rowKey, AppState.pivotConfig.rowValues[rowKey], values.length);
}

function updateSection3ChartsAndPivot() {
  renderPivotTable();
  renderFamilyAllocationChart();
  renderPerPersonCharts();
}

/* ==========================================================================
   Chart.js 상시 비중(%), 평가금액, 수익률 데이터 라벨 플러그인
   ========================================================================== */
const alwaysShowLabelsPlugin = {
  id: 'alwaysShowLabels',
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    ctx.save();
    ctx.font = 'bold 11px Inter, sans-serif';

    chart.data.datasets.forEach((dataset, i) => {
      const meta = chart.getDatasetMeta(i);
      if (meta.hidden) return;

      meta.data.forEach((element, index) => {
        const val = dataset.data[index];
        if (val === undefined || val === null || val === 0) return;

        let labelText = '';

        if (dataset.richInfoList && dataset.richInfoList[index]) {
          const info = dataset.richInfoList[index];
          const valStr = formatCurrencyCompact(info.valuation);
          const ratioStr = `${info.ratio.toFixed(1)}%`;
          const retStr = (info.returnRate >= 0 ? '+' : '') + info.returnRate.toFixed(2) + '%';
          labelText = `${valStr} (${ratioStr}) | ${retStr}`;
        } else if (dataset.ratioList && dataset.ratioList[index] !== undefined) {
          labelText = `${dataset.ratioList[index].toFixed(1)}%`;
        } else if (typeof val === 'number') {
          labelText = Math.abs(val) > 100 ? formatCurrencyCompact(val) : val.toFixed(1) + '%';
        }

        const position = element.tooltipPosition();
        const textWidth = ctx.measureText(labelText).width;

        const isHorizontal = chart.options.indexAxis === 'y';
        ctx.textAlign = isHorizontal ? 'left' : 'center';

        const xPos = isHorizontal ? Math.max(position.x + 8, 20) : position.x;
        const yPos = isHorizontal ? position.y + 4 : Math.max(position.y - 12, 18);

        ctx.fillStyle = 'rgba(15, 23, 42, 0.85)';
        ctx.beginPath();
        if (isHorizontal) {
          ctx.roundRect(xPos - 2, yPos - 12, textWidth + 10, 16, 4);
        } else {
          ctx.roundRect(xPos - textWidth / 2 - 4, yPos - 14, textWidth + 8, 16, 4);
        }
        ctx.fill();

        ctx.fillStyle = dataset.richInfoList && dataset.richInfoList[index] && dataset.richInfoList[index].returnRate < 0 ? '#fca5a5' : '#38bdf8';
        ctx.fillText(labelText, xPos + (isHorizontal ? 3 : 0), yPos);
      });
    });
    ctx.restore();
  }
};

/* ==========================================================================
   3. 파서 & 데이터 업로드 & localStorage 영구 보존
   ========================================================================== */
function saveToLocalStorage(rawJsonArray, fileName, explicitHeaders = null) {
  try {
    localStorage.setItem('asset_dashboard_raw_json', JSON.stringify(rawJsonArray));
    localStorage.setItem('asset_dashboard_file_name', fileName || '');
    localStorage.setItem('asset_dashboard_column_map', JSON.stringify(AppState.columnMap));
    localStorage.setItem('asset_dashboard_filter_keys', JSON.stringify(AppState.filterKeys));
    if (explicitHeaders) {
      localStorage.setItem('asset_dashboard_raw_headers', JSON.stringify(explicitHeaders));
    }
  } catch (e) {
    console.warn('localStorage 저장 제한:', e);
  }
}

function loadFromLocalStorage() {
  try {
    const savedRawJson = localStorage.getItem('asset_dashboard_raw_json');
    const savedFileName = localStorage.getItem('asset_dashboard_file_name');
    const savedColumnMap = localStorage.getItem('asset_dashboard_column_map');
    const savedFilterKeys = localStorage.getItem('asset_dashboard_filter_keys');
    const savedRawHeaders = localStorage.getItem('asset_dashboard_raw_headers');

    let explicitHeaders = null;
    if (savedRawHeaders) {
      try { explicitHeaders = JSON.parse(savedRawHeaders); } catch(err){}
    }

    if (savedRawJson) {
      const parsed = JSON.parse(savedRawJson);
      if (Array.isArray(parsed) && parsed.length > 0) {
        processRawJsonData(parsed, savedFileName || '이전 저장 데이터', false, explicitHeaders);
        
        // localStorage에 저장되어 있던 최신 columnMap과 filterKeys를 최종 복원
        if (savedColumnMap) AppState.columnMap = JSON.parse(savedColumnMap);
        if (savedFilterKeys) AppState.filterKeys = JSON.parse(savedFilterKeys);

        // UI 필터 및 테이블 헤더 다시 렌더링 동기화
        updateRawDataTableHeader();
        init8MultiSelectFilters();
        initPivotRowMultiSelect();
        populateDateDropdowns();

        updateSection1();
        updateSection2();
        updateSection3();

        $('dashboardGrid').classList.remove('hidden');
        $('dropZone').classList.add('hidden');
        return true;
      }
    }
  } catch (e) {
    console.warn('localStorage 불러오기 실패:', e);
  }
  return false;
}

function showToast(status, message, detail = '') {
  const container = $('toastNotification');
  const card = $('toastContent');
  const icon = $('toastIcon');
  const title = $('toastTitle');
  const msgElem = $('toastMessage');
  const detailElem = $('toastDetail');

  if (status === 'success') {
    card.className = 'toast-card glass-panel success';
    icon.innerHTML = '<i class="fa-solid fa-circle-check"></i>';
    title.textContent = '업로드 성공!';
    title.style.color = '#10b981';
    detailElem.classList.add('hidden');
    if (window.confetti) confetti({ particleCount: 70, spread: 60, origin: { y: 0.2 } });
  } else {
    card.className = 'toast-card glass-panel error';
    icon.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i>';
    title.textContent = '업로드 실패';
    title.style.color = '#ef4444';
    if (detail) {
      detailElem.textContent = '실패 사유: ' + detail;
      detailElem.classList.remove('hidden');
    } else {
      detailElem.classList.add('hidden');
    }
  }

  msgElem.textContent = message;
  container.classList.remove('hidden');
  setTimeout(hideToast, 5000);
}

function hideToast() {
  $('toastNotification').classList.add('hidden');
}

function handleFileUpload(file) {
  if (!file) return;
  const fileName = file.name.toLowerCase();

  // 이전 오염된 LocalStorage 정보 및 피벗 필터 캐시 즉시 클리어
  localStorage.removeItem('asset_dashboard_column_map');
  localStorage.removeItem('asset_dashboard_filter_keys');
  localStorage.removeItem('asset_dashboard_raw_headers');
  AppState.pivotConfig.rowValues = {};

  if (fileName.endsWith('.csv')) {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: (results) => {
        if (results.errors && results.errors.length > 0) {
          showToast('error', 'CSV 파일 분석 중 오류가 발생했습니다.', results.errors[0].message);
          return;
        }
        const headers = results.meta && results.meta.fields ? results.meta.fields : Object.keys(results.data[0] || {});
        processRawJsonData(results.data, file.name, true, headers);
      },
      error: (err) => showToast('error', 'CSV 파싱 실패', err.message)
    });
  } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheet];

        // 2차원 배열(rows)로 1번째 행(Row 0) 실제 엑셀 헤더 정밀 확보
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        if (!rows || rows.length < 2) {
          showToast('error', '엑셀 분석 실패', '엑셀 파일 내 유효한 데이터 행이 부족합니다.');
          return;
        }

        const rawHeaders = rows[0].map(h => String(h || '').trim());
        const dataRows = rows.slice(1);

        const jsonArray = dataRows.map(r => {
          const obj = {};
          rawHeaders.forEach((h, idx) => {
            if (h) obj[h] = r[idx] !== undefined ? r[idx] : '';
          });
          return obj;
        });

        processRawJsonData(jsonArray, file.name, true, rawHeaders);
      } catch (err) {
        showToast('error', '엑셀 분석 실패', err.message || '파일이 손상되었거나 유효하지 않은 엑셀 형식입니다.');
      }
    };
    reader.readAsArrayBuffer(file);
  } else {
    showToast('error', '지원되지 않는 파일 형식', '.xlsx, .xls, .csv 확장자의 파일만 업로드할 수 있습니다.');
  }
}

function parseRobustNumber(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;

  let str = String(val).trim();
  if (!str) return 0;

  let isNegative = false;
  if (/[△▼∇▼]/.test(str)) {
    isNegative = true;
    str = str.replace(/[△▼∇▼]/g, '');
  }

  if (str.startsWith('(') && str.endsWith(')')) {
    isNegative = true;
    str = str.substring(1, str.length - 1);
  } else if (str.startsWith('-')) {
    isNegative = true;
  }

  let cleanStr = str.replace(/[^0-9.-]+/g, '');
  if (!cleanStr || cleanStr === '.' || cleanStr === '-') return 0;

  let num = parseFloat(cleanStr);
  return isNaN(num) ? 0 : (isNegative ? -Math.abs(num) : num);
}

/**
 * 엑셀 데이터의 단일 행(row)에서 candidate 가능 키 리스트 중 일치하는 실제 컬럼명과 그 값을 탐색합니다.
 * 섣부른 includes 부분 일치로 인한 타 컬럼 오진(예: '구분' 키워드가 '구분 반도체/전력'에 오진 매핑)을 방지합니다.
 */
function findRowKeyAndValue(row, possibleKeys) {
  if (!row) return { matchedKey: '', value: '' };
  const rowKeys = Object.keys(row);

  // 1차: 정확한 텍스트 및 공백 제거 소문자 완전 일치 비교 (최우선)
  for (const pKey of possibleKeys) {
    const cleanPKey = pKey.replace(/\s+/g, '').toLowerCase();
    for (const rKey of rowKeys) {
      if (rKey.replace(/\s+/g, '').toLowerCase() === cleanPKey && row[rKey] !== undefined && row[rKey] !== '') {
        return { matchedKey: rKey, value: row[rKey] };
      }
    }
  }

  // 2차: 부분 포함(includes) 비교 (pKey 길이가 3자 이상인 경우에만 적용하여 섣부른 오진 방지)
  for (const pKey of possibleKeys) {
    const cleanPKey = pKey.replace(/\s+/g, '').toLowerCase();
    if (cleanPKey.length < 3) continue; // '구분' 등 짧은 단어의 포함 오진 방지
    for (const rKey of rowKeys) {
      if (rKey.replace(/\s+/g, '').toLowerCase().includes(cleanPKey) && row[rKey] !== undefined && row[rKey] !== '') {
        return { matchedKey: rKey, value: row[rKey] };
      }
    }
  }
  return { matchedKey: '', value: '' };
}

function findRowValue(row, possibleKeys) {
  return findRowKeyAndValue(row, possibleKeys).value;
}

/**
 * Raw Data Table 헤더의 <th> 텍스트를 엑셀 업로드 시 감지된 바뀐 컬럼 명칭으로 동적 업데이트합니다.
 */
function updateRawDataTableHeader() {
  const thElements = document.querySelectorAll('#rawDataTable th[data-sort]');
  const map = AppState.columnMap;

  const sortKeyToField = {
    '평가일': '평가일',
    '운용사': '운용사',
    '성명': '성명',
    '구분': '구분',
    '구분2': '구분2',
    '구분3': '구분3',
    '국내외': '국내외',
    '국가_상품': '국가_상품',
    '투자상품': '투자상품',
    '투자원금': '투자원금',
    '평가금액': '평가금액',
    '증감': '증감',
    '수익률': '수익률',
    '환매여부': '환매여부'
  };

  thElements.forEach(th => {
    const dataSort = th.getAttribute('data-sort');
    const field = sortKeyToField[dataSort];
    if (field && map[field]) {
      th.innerHTML = `${map[field]} <i class="fa-solid fa-sort"></i>`;
    }
  });
}

function processRawJsonData(jsonArray, fileName = '', saveToStorage = true, rawHeadersInput = null) {
  if (!jsonArray || jsonArray.length === 0) {
    showToast('error', '데이터 로딩 실패', '파일 내에 유효한 데이터 행이 존재하지 않습니다.');
    return;
  }

  // 1. 실제 엑셀 Header 명칭 배열 확보 (A~Q열)
  let excelHeaders = [];
  if (Array.isArray(rawHeadersInput) && rawHeadersInput.length > 0) {
    excelHeaders = rawHeadersInput;
  } else {
    const firstRow = jsonArray[0] || {};
    excelHeaders = Object.keys(firstRow).filter(k => k !== 'id' && k !== 'raw');
  }

  // 엑셀 열 1대1 위치 직결 매핑
  // index 0(A): 평가일
  // index 1(B): 운용사 (필터1)
  // index 2(C): 성명 (필터2)
  // index 3(D): 구분 -> 엑셀 1행 D열 명칭 (예: '주식/ETF/채권') (필터3)
  // index 4(E): 구분2 -> 엑셀 1행 E열 명칭 (예: '성장/안정/배당') (필터4)
  // index 5(F): 구분3 -> 엑셀 1행 F열 명칭 (예: '구분 반도체/전력') (필터5)
  // index 6(G): 국내외 (필터6)
  // index 7(H): 국가_상품 -> 엑셀 1행 H열 명칭 (예: 'Currency') (필터7)
  // index 8(I): 투자상품 (필터8)

  const hDate = excelHeaders[0] || '평가일';
  const hManager = excelHeaders[1] || '운용사';
  const hName = excelHeaders[2] || '성명';
  const hCat1 = excelHeaders[3] || '주식/ETF/채권';
  const hCat2 = excelHeaders[4] || '성장/안정/배당';
  const hCat3 = excelHeaders[5] || '구분 반도체/전력';
  const hDomestic = excelHeaders[6] || '국내외';
  const hCountry = excelHeaders[7] || 'Currency';
  const hProduct = excelHeaders[8] || '투자상품';

  AppState.columnMap = {
    평가일: hDate,
    운용사: hManager,
    성명: hName,
    구분: hCat1,
    구분2: hCat2,
    구분3: hCat3,
    국내외: hDomestic,
    국가_상품: hCountry,
    투자상품: hProduct,
    투자원금: excelHeaders[9] || '투자원금',
    평가금액: excelHeaders[10] || '평가금액',
    증감: excelHeaders[11] || '증감',
    수익률: excelHeaders[12] || '수익률',
    환매여부: excelHeaders[13] || '환매여부',
    최종평가금액: excelHeaders[14] || '최종평가금액',
    pick: excelHeaders[15] || 'pick',
    비고: excelHeaders[16] || '비고'
  };

  // ★ 8개 다중 선택 필터 라벨을 엑셀 1행의 실제 컬럼명 텍스트로 100% 직결
  AppState.filterKeys = [hManager, hName, hCat1, hCat2, hCat3, hDomestic, hCountry, hProduct];

  if (!AppState.filterKeys.includes(AppState.pivotConfig.rows[0])) {
    AppState.pivotConfig.rows = [AppState.filterKeys[0]];
  }

  // ★ 새로 업로드된 엑셀 데이터의 모든 세부 항목("테크" 등)이 피벗에 100% 반영되도록 피벗 필터 초기화
  AppState.pivotConfig.rowValues = {};

  // 2. 각 행별 데이터 파싱 및 안전한 셀 값 바인딩
  const parsed = jsonArray.map((row, index) => {
    const getValByColNameOrIdx = (colName, colIdx) => {
      // 1순위: 엑셀 실제 컬럼명 탐색
      if (colName && row[colName] !== undefined && row[colName] !== null && row[colName] !== '') {
        return row[colName];
      }
      // 2순위: 엑셀 Index 위치 탐색
      if (excelHeaders[colIdx] && row[excelHeaders[colIdx]] !== undefined && row[excelHeaders[colIdx]] !== '') {
        return row[excelHeaders[colIdx]];
      }
      return '';
    };

    const dateRaw = getValByColNameOrIdx(hDate, 0);
    const managerRaw = getValByColNameOrIdx(hManager, 1);
    const nameRaw = getValByColNameOrIdx(hName, 2);
    
    // 3번 필터 (D열 - '주식/ETF/채권')
    const cat1Raw = getValByColNameOrIdx(hCat1, 3);
    // 4번 필터 (E열 - '성장/안정/배당')
    const cat2Raw = getValByColNameOrIdx(hCat2, 4);
    // 5번 필터 (F열 - '구분 반도체/전력')
    const cat3Raw = getValByColNameOrIdx(hCat3, 5);

    const domesticRaw = getValByColNameOrIdx(hDomestic, 6);
    const countryRaw = getValByColNameOrIdx(hCountry, 7); // H열 (Currency)
    const productRaw = getValByColNameOrIdx(hProduct, 8);
    const principalRaw = getValByColNameOrIdx(AppState.columnMap.투자원금, 9);
    const valuationRaw = getValByColNameOrIdx(AppState.columnMap.평가금액, 10);
    const diffRaw = getValByColNameOrIdx(AppState.columnMap.증감, 11);
    const returnRaw = getValByColNameOrIdx(AppState.columnMap.수익률, 12);

    const redemptionRaw = getValByColNameOrIdx(AppState.columnMap.환매여부, 13);
    const finalValRaw = getValByColNameOrIdx(AppState.columnMap.최종평가금액, 14);
    const pickRaw = getValByColNameOrIdx(AppState.columnMap.pick, 15);
    const remarkRaw = getValByColNameOrIdx(AppState.columnMap.비고, 16);

    const principal = parseRobustNumber(principalRaw);
    const valuation = parseRobustNumber(valuationRaw);
    let diff = parseRobustNumber(diffRaw);

    if (diff === 0 && (valuation !== 0 || principal !== 0)) {
      diff = valuation - principal;
    }

    let returnRate = 0;
    if (principal > 0) {
      returnRate = Number(((diff / principal) * 100).toFixed(2));
    } else {
      let parsedRet = parseRobustNumber(returnRaw);
      if (typeof returnRaw === 'string' && returnRaw.includes('%')) {
        returnRate = parsedRet;
      } else if (Math.abs(parsedRet) > 0 && Math.abs(parsedRet) <= 1.0) {
        returnRate = Number((parsedRet * 100).toFixed(2));
      } else {
        returnRate = parsedRet;
      }
    }

    const itemObj = {
      id: `item-${index}`,
      평가일: formatDate(dateRaw),
      운용사: String(managerRaw || '기타운용사').trim(),
      성명: String(nameRaw || '기본투자자').trim(),
      구분: String(cat1Raw || '-').trim(),   // D열 데이터
      구분2: String(cat2Raw || '-').trim(),  // E열 데이터
      구분3: String(cat3Raw || '-').trim(),  // F열 데이터
      국내외: String(domesticRaw || '국내').trim(),
      국가_상품: String(countryRaw || '한국').trim(),
      투자상품: String(productRaw || '무명 자산').trim(),
      투자원금: principal,
      평가금액: valuation,
      증감: diff,
      수익률: returnRate,
      환매여부: String(redemptionRaw || '보유').trim(),
      최종평가금액: parseRobustNumber(finalValRaw),
      pick: String(pickRaw || ''),
      비고: String(remarkRaw || '')
    };

    // 실제 엑셀 헤더명으로도 동적 데이터 접근이 가능하도록 바인딩
    Object.keys(AppState.columnMap).forEach(stdField => {
      const actualColName = AppState.columnMap[stdField];
      if (actualColName) {
        itemObj[actualColName] = itemObj[stdField];
      }
    });

    return itemObj;
  });

  AppState.rawDataset = parsed;

  if (saveToStorage) {
    saveToLocalStorage(jsonArray, fileName, excelHeaders);
  }

  // Raw Table 헤더 텍스트 변경 적용
  updateRawDataTableHeader();

  init8MultiSelectFilters();
  initPivotRowMultiSelect();
  populateDateDropdowns();
  
  updateSection1();
  updateSection2();
  updateSection3();

  $('dashboardGrid').classList.remove('hidden');
  $('dropZone').classList.add('hidden');

  showToast('success', `${fileName ? `'${fileName}' ` : ''}총 ${parsed.length}건 데이터 로딩 완료!`, `스마트 헤더 감지 매핑이 적용되었습니다.`);
}

function isRedeemed(val) {
  if (!val) return false;
  const str = String(val).trim();
  return str === '환매' || str === '환매완료' || str === '해지';
}

function formatDate(val) {
  if (!val) return '2024-02-25';
  let str = String(val).trim().replace(/\s+/g, '').replace(/\./g, '-').replace(/\//g, '-');
  if (/^\d{8}$/.test(str)) {
    return `${str.substring(0, 4)}-${str.substring(4, 6)}-${str.substring(6, 8)}`;
  }
  if (typeof val === 'number' && val > 30000 && val < 60000) {
    return new Date((val - (25567 + 2)) * 86400 * 1000).toISOString().split('T')[0];
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;

  const cleanDigits = str.replace(/[^0-9]/g, '');
  if (cleanDigits.length >= 8) {
    const y = cleanDigits.substring(0, 4);
    const m = cleanDigits.substring(4, 6);
    const d = cleanDigits.substring(6, 8);
    const yr = parseInt(y, 10);
    if (yr >= 1900 && yr <= 2100) return `${y}-${m}-${d}`;
  }
  return str.length >= 10 ? str.substring(0, 10) : str;
}

/* ==========================================================================
   4. 샘플 데이터
   ========================================================================== */
function loadSampleData() {
  const sampleList = [
    { '평가일': '20240225', '운용사': '한국투자', '성명': '김성중', '주식/ETF/채권': '현금', '성장/안정/배당': '안정형', '구분 반도체/전력': '일반', '국내외': '국내', 'Currency': 'KRW', '투자상품': 'CMA', '투자원금': 117343976, '평가금액': 117343976, '증감': 0, '수익률': '0.0%', '환매여부': '보유', '비고': '' },
    { '평가일': '20240225', '운용사': '한국투자', '성명': '김성중', '주식/ETF/채권': '채권', '성장/안정/배당': '배당형', '구분 반도체/전력': '전력', '국내외': '국내', 'Currency': 'KRW', '투자상품': '우리금융', '투자원금': 99999043, '평가금액': 96013108, '증감': '△3,985,935', '수익률': '△4.0%', '환매여부': '보유', '비고': '옵션' },
    { '평가일': '20240225', '운용사': '한국투자', '성명': '김성중', '주식/ETF/채권': 'ETF', '성장/안정/배당': '성장형', '구분 반도체/전력': '반도체', '국내외': '국내', 'Currency': 'KRW', '투자상품': 'KIS 투자조합', '투자원금': 135427536, '평가금액': 135008267, '증감': '△419,269', '수익률': '△0.3%', '환매여부': '보유', '비고': '' },
    { '평가일': '20240225', '운용사': '한국투자', '성명': '김선윤', '주식/ETF/채권': '현금', '성장/안정/배당': '안정형', '구분 반도체/전력': '일반', '국내외': '국내', 'Currency': 'KRW', '투자상품': 'CMA', '투자원금': 19503574, '평가금액': 19503574, '증감': 0, '수익률': '0.0%', '환매여부': '보유', '비고': '' },
    { '평가일': '20240225', '운용사': '삼성증권', '성명': '김민서', '주식/ETF/채권': '주식', '성장/안정/배당': '성장형', '구분 반도체/전력': '반도체', '국내외': '해외', 'Currency': 'USD', '투자상품': '애플', '투자원금': 15000000, '평가금액': 16500000, '증감': '1,500,000', '수익률': '10.0%', '환매여부': '보유', '비고': '' },
    { '평가일': '20240325', '운용사': '한국투자', '성명': '김성중', '주식/ETF/채권': '채권', '성장/안정/배당': '배당형', '구분 반도체/전력': '전력', '국내외': '국내', 'Currency': 'KRW', '투자상품': '우리금융', '투자원금': 99999043, '평가금액': 98500000, '증감': '△1,499,043', '수익률': '△1.5%', '환매여부': '보유', '비고': '옵션' },
    { '평가일': '20240325', '운용사': '한국투자', '성명': '김성중', '주식/ETF/채권': 'ETF', '성장/안정/배당': '성장형', '구분 반도체/전력': '반도체', '국내외': '국내', 'Currency': 'KRW', '투자상품': 'KIS 투자조합', '투자원금': 135427536, '평가금액': 137000000, '증감': '1,572,464', '수익률': '1.2%', '환매여부': '보유', '비고': '' },
    { '평가일': '20240325', '운용사': '삼성증권', '성명': '김민서', '주식/ETF/채권': '주식', '성장/안정/배당': '성장형', '구분 반도체/전력': '반도체', '국내외': '해외', 'Currency': 'USD', '투자상품': '애플', '투자원금': 15000000, '평가금액': 18000000, '증감': '3,000,000', '수익률': '20.0%', '환매여부': '보유', '비고': '' },
    { '평가일': '20241231', '운용사': '한국투자', '성명': '김성중', '주식/ETF/채권': '채권', '성장/안정/배당': '배당형', '구분 반도체/전력': '전력', '국내외': '국내', 'Currency': 'KRW', '투자상품': '우리금융', '투자원금': 99999043, '평가금액': 102000000, '증감': '2,000,957', '수익률': '2.0%', '환매여부': '보유', '비고': '2024년말' },
    { '평가일': '20250225', '운용사': '한국투자', '성명': '김성중', '주식/ETF/채권': '채권', '성장/안정/배당': '배당형', '구분 반도체/전력': '전력', '국내외': '국내', 'Currency': 'KRW', '투자상품': '우리금융', '투자원금': 99999043, '평가금액': 104500000, '증감': '4,500,957', '수익률': '4.5%', '환매여부': '보유', '비고': '2025년최신' }
  ];

  const sampleHeaders = ['평가일', '운용사', '성명', '주식/ETF/채권', '성장/안정/배당', '구분 반도체/전력', '국내외', 'Currency', '투자상품', '투자원금', '평가금액', '증감', '수익률', '환매여부', '최종평가금액', 'pick', '비고'];
  processRawJsonData(sampleList, '샘플 데이터', true, sampleHeaders);
}

/* ==========================================================================
   5. 8개 항목 독립 다중 선택 필터 시스템
   ========================================================================== */
function init8MultiSelectFilters() {
  ['sec1', 'sec2', 'sec3'].forEach(secKey => {
    AppState[secKey].filters = {};
    const container = $(`${secKey}FilterContainer`);
    container.innerHTML = '';

    AppState.filterKeys.forEach(fKey => {
      const items = getUniqueList(fKey);
      AppState[secKey].filters[fKey] = [...items];

      const groupDiv = document.createElement('div');
      groupDiv.className = 'filter-group dropdown-multi';

      const label = document.createElement('label');
      label.textContent = `${fKey} (복수선택)`;

      const box = document.createElement('div');
      box.className = 'multi-select-box';
      box.id = `${secKey}_${fKey}_box`;
      box.innerHTML = `<span class="multi-select-label">전체 ${fKey}</span><i class="fa-solid fa-chevron-down"></i>`;

      const dropdown = document.createElement('div');
      dropdown.className = 'multi-select-dropdown hidden';
      dropdown.id = `${secKey}_${fKey}_dropdown`;

      groupDiv.appendChild(label);
      groupDiv.appendChild(box);
      groupDiv.appendChild(dropdown);
      container.appendChild(groupDiv);

      box.addEventListener('click', (e) => {
        e.stopPropagation();
        const isHidden = dropdown.classList.contains('hidden');
        document.querySelectorAll('.multi-select-dropdown').forEach(el => el.classList.add('hidden'));
        if (isHidden) dropdown.classList.remove('hidden');
      });

      populateMultiSelectItems(dropdown, box, items, fKey, secKey);
    });
  });
}

function getUniqueList(key) {
  return Array.from(new Set(AppState.rawDataset.map(d => {
    let val = d[key];
    if (val === undefined) {
      const stdKey = Object.keys(AppState.columnMap).find(k => AppState.columnMap[k] === key);
      if (stdKey && d[stdKey] !== undefined) val = d[stdKey];
    }
    return String(val !== undefined ? val : '');
  }))).sort();
}

function populateMultiSelectItems(dropdown, box, items, filterKey, secKey) {
  dropdown.innerHTML = '';

  const allLabel = document.createElement('label');
  allLabel.className = 'multi-option font-bold';
  allLabel.innerHTML = `<input type="checkbox" value="ALL" checked> 전체 선택 (${items.length}개)`;
  dropdown.appendChild(allLabel);

  items.forEach(item => {
    const lbl = document.createElement('label');
    lbl.className = 'multi-option';
    lbl.innerHTML = `<input type="checkbox" value="${item}" checked> ${item}`;
    dropdown.appendChild(lbl);
  });

  const checkboxes = dropdown.querySelectorAll('input[type="checkbox"]');
  checkboxes.forEach(cb => {
    cb.addEventListener('change', (e) => {
      const val = e.target.value;
      if (val === 'ALL') {
        checkboxes.forEach(c => c.checked = e.target.checked);
      } else {
        const allCb = dropdown.querySelector('input[value="ALL"]');
        const itemCbs = Array.from(checkboxes).filter(c => c.value !== 'ALL');
        allCb.checked = itemCbs.every(c => c.checked);
      }

      const selected = Array.from(checkboxes)
        .filter(c => c.checked && c.value !== 'ALL')
        .map(c => c.value);

      AppState[secKey].filters[filterKey] = selected;
      updateMultiSelectBoxLabel(box, filterKey, selected, items.length);

      if (secKey === 'sec1') updateSection1();
      else if (secKey === 'sec2') updateSection2();
      else if (secKey === 'sec3') updateSection3();
    });
  });
}

function updateMultiSelectBoxLabel(box, filterKey, selectedList, totalCount) {
  const labelElem = box.querySelector('.multi-select-label');
  if (selectedList.length === 0) labelElem.textContent = `선택 없음`;
  else if (selectedList.length === totalCount) labelElem.textContent = `전체 ${filterKey}`;
  else if (selectedList.length === 1) labelElem.textContent = selectedList[0];
  else labelElem.textContent = `${selectedList[0]} 외 ${selectedList.length - 1}건`;
}

function populateDateDropdowns() {
  const dates = Array.from(new Set(AppState.rawDataset.map(d => d.평가일))).sort();
  const latestDate = dates[dates.length - 1] || '2024-02-25';

  fillDateSelect('sec1DateSelect', dates, latestDate, '마지막 날짜');
  AppState.sec1.selectedDate = latestDate;

  fillDateSelect('sec2EndDateSelect', dates, latestDate, '최종일');
  AppState.sec2.endDate = latestDate;

  fillDateSelect('sec3DateSelect', dates, latestDate, '선택 평가일');
  AppState.sec3.selectedDate = latestDate;
}

function fillDateSelect(selectId, dates, defaultDate, labelSuffix) {
  const sel = $(selectId);
  sel.innerHTML = '';
  dates.slice().reverse().forEach(date => {
    const opt = document.createElement('option');
    opt.value = date;
    opt.textContent = `${date} ${date === defaultDate ? `(${labelSuffix})` : ''}`;
    if (date === defaultDate) opt.selected = true;
    sel.appendChild(opt);
  });
}

function filterDatasetBySection(secKey) {
  const f = AppState[secKey].filters;
  return AppState.rawDataset.filter(item => {
    if (isRedeemed(item.환매여부)) return false;
    for (const fKey of AppState.filterKeys) {
      if (f[fKey] && f[fKey].length > 0) {
        let val = item[fKey];
        if (val === undefined) {
          const stdKey = Object.keys(AppState.columnMap).find(k => AppState.columnMap[k] === fKey);
          if (stdKey) val = item[stdKey];
        }
        if (!f[fKey].includes(String(val !== undefined ? val : ''))) {
          return false;
        }
      }
    }
    return true;
  });
}

function getItemValueByHeaderKey(item, key) {
  if (!item) return '';
  if (item[key] !== undefined && item[key] !== null && item[key] !== '') {
    return item[key];
  }
  const stdKey = Object.keys(AppState.columnMap).find(k => AppState.columnMap[k] === key);
  if (stdKey && item[stdKey] !== undefined && item[stdKey] !== null) {
    return item[stdKey];
  }
  return '';
}

function getSection3PivotDataset() {
  const sec3Data = filterDatasetBySection('sec3');
  const rowValuesFilter = AppState.pivotConfig.rowValues;

  return sec3Data.filter(item => {
    for (const rKey of AppState.pivotConfig.rows) {
      const allowedVals = rowValuesFilter[rKey];
      if (allowedVals && allowedVals.length > 0) {
        const itemVal = String(getItemValueByHeaderKey(item, rKey));
        if (!allowedVals.includes(itemVal)) {
          return false;
        }
      }
    }
    return true;
  });
}

/* ==========================================================================
   6. SECTION 1. 누적 성과
   ========================================================================== */
function updateSection1() {
  const data = filterDatasetBySection('sec1');
  if (data.length === 0) return;

  const dates = Array.from(new Set(data.map(d => d.평가일))).sort();
  const baseDate = '2024-02-25';
  const firstDate = dates.includes(baseDate) ? baseDate : dates[0];
  const selectedDate = AppState.sec1.selectedDate !== 'LATEST' ? AppState.sec1.selectedDate : dates[dates.length - 1];

  const firstDateItems = data.filter(d => d.평가일 === firstDate);
  const basePrincipal = firstDateItems.reduce((s, i) => s + i.투자원금, 0);

  const selectedItems = data.filter(d => d.평가일 === selectedDate);
  const totalValuation = selectedItems.reduce((s, i) => s + i.평가금액, 0);

  const totalDiff = totalValuation - basePrincipal;
  const totalReturn = basePrincipal > 0 ? ((totalDiff / basePrincipal) * 100) : 0;

  $('sec1TotalPrincipal').textContent = formatCurrency(basePrincipal);
  $('sec1TotalValuation').textContent = formatCurrency(totalValuation);
  $('sec1SelectedDateLabel').textContent = `${selectedDate} 기준`;
  $('sec1ValuationSub').textContent = `${selectedDate} 기준 ${selectedItems.length}개 보유 종목`;

  const diffElem = $('sec1TotalDiff');
  diffElem.textContent = (totalDiff >= 0 ? '+' : '') + formatCurrency(totalDiff);
  $('sec1DiffCard').className = `kpi-card glass-panel ${totalDiff >= 0 ? 'positive' : 'negative'}`;

  const returnElem = $('sec1TotalReturn');
  returnElem.textContent = (totalReturn >= 0 ? '+' : '') + totalReturn.toFixed(2) + '%';
  $('sec1ReturnCard').className = `kpi-card glass-panel ${totalReturn >= 0 ? 'positive' : 'negative'}`;
  $('sec1ReturnBadge').className = `kpi-badge ${totalReturn >= 0 ? 'positive' : 'negative'}`;
  $('sec1ReturnBadge').textContent = totalReturn >= 0 ? '▲ 누적 수익 달성' : '▼ 누적 손실 발생';

  renderSec1AnnualTable();
}

function renderSec1AnnualTable() {
  const tbody = $('sec1AnnualTableBody');
  tbody.innerHTML = '';

  // ★ Section 1의 8개 다중 선택 필터 조건이 100% 반영된 필터링 데이터셋 참조
  const data = filterDatasetBySection('sec1');
  if (!data || data.length === 0) return;

  const dates = Array.from(new Set(data.map(d => d.평가일))).sort();
  const years = Array.from(new Set(dates.map(d => d.substring(0, 4)))).sort();
  const currentYear = new Date().getFullYear().toString();

  let prevYearValuation = 0;

  years.forEach((yr, idx) => {
    const yrDates = dates.filter(d => d.startsWith(yr));
    const isCurrentYr = (yr === currentYear || idx === years.length - 1);
    
    // 당해 연도의 마지막 평가일자 (예: 2024-12-25, 2025-12-31, 2026-08-08)
    const yearEndDate = yrDates[yrDates.length - 1];
    const yearEndItems = data.filter(d => d.평가일 === yearEndDate);
    
    // ★ 선택된 필터 조건(예: 성명=김선우, 운용사=한국투자 등)에 맞추어 당해 연도말 평가금액 실시간 동적 재계산
    const currentYearValuation = yearEndItems.reduce((s, i) => s + (Number(i.평가금액) || 0), 0);

    let baseValuation = prevYearValuation;
    let baseLabel = '';

    if (idx === 0 || baseValuation === 0) {
      // 첫 해(2024년) 기준 금액: 선택된 필터 조건하에서의 최초 평가일자(예: 2024-02-25) 평가금액 합계
      const firstDate = yrDates[0];
      const firstDateItems = data.filter(d => d.평가일 === firstDate);
      baseValuation = firstDateItems.reduce((s, i) => s + (Number(i.평가금액) || 0), 0);
      baseLabel = ` (최초일 ${firstDate})`;
    }

    const diff = currentYearValuation - baseValuation;
    const returnRate = baseValuation > 0 ? ((diff / baseValuation) * 100) : 0;

    const tr = document.createElement('tr');
    const labelTitle = isCurrentYr ? `${yr}년말 기준 (${yearEndDate})` : `${yr}년말 기준 (${yearEndDate})`;
    const retClass = returnRate >= 0 ? 'positive-text' : 'negative-text';

    tr.innerHTML = `
      <td class="font-bold">${labelTitle}</td>
      <td class="text-right num-col">${formatCurrency(baseValuation)}${idx === 0 ? `<span class="text-xs text-muted block">${baseLabel}</span>` : ''}</td>
      <td class="text-right num-col font-bold text-cyan-400" style="font-size:1.05rem;">${formatCurrency(currentYearValuation)}</td>
      <td class="text-right num-col ${diff >= 0 ? 'positive-text' : 'negative-text'}">${diff >= 0 ? '+' : ''}${formatCurrency(diff)}</td>
      <td class="text-right num-col ${retClass}">${returnRate >= 0 ? '+' : ''}${returnRate.toFixed(2)}%</td>
      <td><span class="badge ${returnRate >= 0 ? '' : 'badge-redeemed'}">${returnRate >= 0 ? '연간 수익' : '연간 손실'}</span></td>
    `;
    tbody.appendChild(tr);

    prevYearValuation = currentYearValuation;
  });
}

/* ==========================================================================
   7. SECTION 2. 누적 시계열 Trend (X축 90도 직각 수직 정렬)
   ========================================================================== */
function updateSection2() {
  const ctx = $('sec2TimeSeriesChart').getContext('2d');
  if (AppState.chartInstances.sec2TimeSeries) AppState.chartInstances.sec2TimeSeries.destroy();

  const data = filterDatasetBySection('sec2');
  if (data.length === 0) return;

  const allDates = Array.from(new Set(data.map(d => d.평가일))).sort();
  const endDate = AppState.sec2.endDate !== 'LATEST' ? AppState.sec2.endDate : allDates[allDates.length - 1];
  const dates = allDates.filter(d => d <= endDate);

  const metricKey = AppState.sec2.metric;
  const names = Array.from(new Set(data.map(d => d.성명))).sort();

  const mainPerson = names.includes('김성중') ? '김성중' : names[0];
  const otherPersons = names.filter(n => n !== mainPerson);

  const palette = ['#6366f1', '#06b6d4', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6'];
  const datasets = [];

  const mainDataPoints = dates.map(date => {
    const items = data.filter(d => d.평가일 === date && d.성명 === mainPerson);
    return items.reduce((s, i) => s + (Number(i[metricKey]) || 0), 0);
  });

  datasets.push({
    label: `${mainPerson} (좌측 Y축 - 상단 Scale)`,
    data: mainDataPoints,
    borderColor: '#6366f1',
    backgroundColor: 'rgba(99, 102, 241, 0.15)',
    borderWidth: 3,
    yAxisID: 'y-left',
    tension: 0.3
  });

  let maxRightValue = 0;

  otherPersons.forEach((pName, idx) => {
    const color = palette[(idx + 1) % palette.length];
    const pPoints = dates.map(date => {
      const items = data.filter(d => d.평가일 === date && d.성명 === pName);
      return items.reduce((s, i) => s + (Number(i[metricKey]) || 0), 0);
    });

    const maxVal = Math.max(...pPoints, 0);
    if (maxVal > maxRightValue) maxRightValue = maxVal;

    datasets.push({
      label: `${pName} (우측 Y축)`,
      data: pPoints,
      borderColor: color,
      backgroundColor: hexToRgba(color, 0.1),
      borderWidth: 2.5,
      yAxisID: 'y-right',
      tension: 0.3
    });
  });

  const minLeftVal = Math.max(maxRightValue * 1.2, 50000000);

  AppState.chartInstances.sec2TimeSeries = new Chart(ctx, {
    type: 'line',
    data: { labels: dates, datasets: datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top', labels: { color: '#cbd5e1', font: { family: 'Inter', size: 12 } } },
        tooltip: {
          backgroundColor: 'rgba(15, 23, 42, 0.9)',
          callbacks: { label: (c) => `${c.dataset.label}: ${formatCurrency(c.parsed.y)}` }
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(255, 255, 255, 0.05)' },
          ticks: {
            color: '#cbd5e1',
            maxRotation: 90, // ★ X축 90도 직각 수직 정렬!
            minRotation: 90,
            font: { size: 11, weight: '600' }
          }
        },
        'y-left': {
          type: 'linear',
          position: 'left',
          suggestedMin: minLeftVal,
          grid: { color: 'rgba(99, 102, 241, 0.1)' },
          ticks: {
            color: '#818cf8',
            callback: (v) => Math.abs(v) >= 100000000 ? (v / 100000000).toFixed(1) + '억원' : (v / 10000).toFixed(0) + '만원'
          },
          title: { display: true, text: `${mainPerson} Scale (좌측)`, color: '#818cf8' }
        },
        'y-right': {
          type: 'linear',
          position: 'right',
          suggestedMax: maxRightValue * 1.1,
          grid: { drawOnChartArea: false },
          ticks: {
            color: '#38bdf8',
            callback: (v) => Math.abs(v) >= 100000000 ? (v / 100000000).toFixed(1) + '억원' : (v / 10000).toFixed(0) + '만원'
          },
          title: { display: true, text: '기타 성명 Scale (우측)', color: '#38bdf8' }
        }
      }
    }
  });
}

/* ==========================================================================
   8. SECTION 3. 평가일기준 평가 ('평가금액' Tab: X/Y축 변경 & X축 90도 수직 표기)
   ========================================================================== */
function updateSection3() {
  const data = filterDatasetBySection('sec3');
  const selDate = AppState.sec3.selectedDate;

  const evalData = data.filter(d => selDate === 'ALL' || d.평가일 === selDate);
  const activeData = evalData.filter(d => !isRedeemed(d.환매여부));

  const totalValuation = activeData.reduce((s, i) => s + i.평가금액, 0);
  const totalPrincipal = activeData.reduce((s, i) => s + i.투자원금, 0);
  const totalDiff = totalValuation - totalPrincipal;
  const totalReturn = totalPrincipal > 0 ? ((totalDiff / totalPrincipal) * 100) : 0;

  const redeemedCount = evalData.length - activeData.length;

  $('sec3TotalValuation').textContent = formatCurrency(totalValuation);
  $('sec3ValuationSub').textContent = `총 ${activeData.length}개 보유 자산${redeemedCount > 0 ? ` ('환매' ${redeemedCount}건 제외)` : ''}`;
  $('sec3TotalPrincipal').textContent = formatCurrency(totalPrincipal);

  const diffElem = $('sec3TotalDiff');
  diffElem.textContent = (totalDiff >= 0 ? '+' : '') + formatCurrency(totalDiff);
  $('sec3DiffCard').className = `kpi-card glass-panel ${totalDiff >= 0 ? 'positive' : 'negative'}`;

  const returnElem = $('sec3TotalReturn');
  returnElem.textContent = (totalReturn >= 0 ? '+' : '') + totalReturn.toFixed(2) + '%';
  $('sec3ReturnCard').className = `kpi-card glass-panel ${totalReturn >= 0 ? 'positive' : 'negative'}`;
  $('sec3ReturnBadge').className = `kpi-badge ${totalReturn >= 0 ? 'positive' : 'negative'}`;
  $('sec3ReturnBadge').textContent = totalReturn >= 0 ? '▲ 수익 달성' : '▼ 손실 발생';

  updateSection3ChartsAndPivot();
  renderRawTable();
}

function getCombinedGroupKey(item, rowKeys) {
  if (!rowKeys || rowKeys.length === 0) return '기타';
  const parts = rowKeys.map(k => {
    const val = getItemValueByHeaderKey(item, k);
    return String(val !== '' ? val : '-').trim();
  });
  return parts.join(' - ');
}

function renderPivotTable() {
  const tbody = $('pivotTableBody');
  tbody.innerHTML = '';

  const data = getSection3PivotDataset();
  const selDate = AppState.sec3.selectedDate;
  const activeData = data.filter(d => (selDate === 'ALL' || d.평가일 === selDate) && !isRedeemed(d.환매여부));

  if (activeData.length === 0) return;

  const selectedRows = AppState.pivotConfig.rows;
  const metricKey = AppState.pivotConfig.metric || '평가금액';
  const groups = {};
  let totalMetricSum = 0;

  activeData.forEach(item => {
    const key = getCombinedGroupKey(item, selectedRows);
    if (!groups[key]) groups[key] = { key, principal: 0, valuation: 0, diff: 0, metricVal: 0, count: 0 };
    groups[key].principal += item.투자원금;
    groups[key].valuation += item.평가금액;
    groups[key].diff += item.증감;
    
    let itemMetric = item.평가금액;
    if (metricKey === '투자원금') itemMetric = item.투자원금;
    else if (metricKey === '증감') itemMetric = item.증감;
    else if (metricKey === '수익률') itemMetric = item.수익률;

    groups[key].metricVal += itemMetric;
    groups[key].count += 1;
    totalMetricSum += Math.abs(itemMetric);
  });

  let groupList = Object.values(groups).map(g => ({
    ...g,
    returnRate: g.principal > 0 ? ((g.diff / g.principal) * 100) : 0,
    ratio: totalMetricSum > 0 ? ((Math.abs(g.metricVal) / totalMetricSum) * 100) : 0
  }));

  const sortCol = AppState.pivotSortState.column || 'valuation';
  const sortDir = AppState.pivotSortState.direction === 'asc' ? 1 : -1;

  groupList.sort((a, b) => {
    let valA = a[sortCol];
    let valB = b[sortCol];

    if (typeof valA === 'string') {
      return valA.localeCompare(valB) * sortDir;
    }
    return (valA - valB) * sortDir;
  });

  groupList.forEach(g => {
    const tr = document.createElement('tr');
    const retClass = g.returnRate >= 0 ? 'positive-text' : 'negative-text';

    tr.innerHTML = `
      <td class="font-bold text-cyan-300">${g.key}</td>
      <td class="text-right num-col">${formatCurrency(g.principal)}</td>
      <td class="text-right num-col font-bold">${formatCurrency(g.valuation)}</td>
      <td class="text-right num-col ${g.diff >= 0 ? 'positive-text' : 'negative-text'}">${g.diff >= 0 ? '+' : ''}${formatCurrency(g.diff)}</td>
      <td class="text-right num-col ${retClass}">${g.returnRate >= 0 ? '+' : ''}${g.returnRate.toFixed(2)}%</td>
      <td class="text-right num-col">
        <div class="flex flex-col items-end">
          <span class="font-bold text-cyan-400">${g.ratio.toFixed(1)}% (${metricKey})</span>
          <div class="progress-bar-bg"><div class="progress-bar-fill" style="width: ${Math.min(g.ratio, 100)}%"></div></div>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

/**
 * 3-A. 가족 통합 포트폴리오 비중 차트
 * - '평가금액' 탭: X축과 Y축 변경(indexAxis: 'y') -> Y축 항목명, X축 수치 ₩ 배치!
 *   X축 눈금 틱 수치도 90도 직각 수직 표기 (maxRotation: 90, minRotation: 90)
 * - '도넛비중' 탭: 우측 색인 범례 배치
 * - '수익률순위' 탭: Y축 항목 배치 (indexAxis: 'y'), 상하 스크롤 적용
 */
function renderFamilyAllocationChart() {
  const ctx = $('familyAllocationChart').getContext('2d');
  if (AppState.chartInstances.familyAllocation) AppState.chartInstances.familyAllocation.destroy();

  const data = getSection3PivotDataset();
  const selDate = AppState.sec3.selectedDate;
  const activeData = data.filter(d => (selDate === 'ALL' || d.평가일 === selDate) && !isRedeemed(d.환매여부));

  if (activeData.length === 0) return;

  const tab = AppState.chartTab;
  const selectedRows = AppState.pivotConfig.rows;
  const metricKey = AppState.pivotConfig.metric || '평가금액';
  const rowLabelText = selectedRows.join(' + ');

  const groups = {};
  let totalValuationSum = 0;

  activeData.forEach(item => {
    const key = getCombinedGroupKey(item, selectedRows);
    if (!groups[key]) groups[key] = { principal: 0, valuation: 0, diff: 0, metricVal: 0 };
    groups[key].principal += item.투자원금;
    groups[key].valuation += item.평가금액;
    groups[key].diff += item.증감;
    
    let itemMetric = item.평가금액;
    if (metricKey === '투자원금') itemMetric = item.투자원금;
    else if (metricKey === '증감') itemMetric = item.증감;
    else if (metricKey === '수익률') itemMetric = item.수익률;

    groups[key].metricVal += itemMetric;
    totalValuationSum += item.평가금액;
  });

  // 평가금액 내림차순 정렬
  let labels = Object.keys(groups);
  if (tab === 'bar' || tab === 'donut') {
    labels.sort((a, b) => groups[b].valuation - groups[a].valuation);
  }

  const palette = [
    '#6366f1', '#06b6d4', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6', '#f97316', '#3b82f6',
    '#84cc16', '#14b8a6', '#a855f7', '#f43f5e', '#0284c7', '#d97706', '#059669', '#4f46e5'
  ];

  const ratios = labels.map(l => totalValuationSum > 0 ? ((groups[l].valuation / totalValuationSum) * 100) : 0);
  const richInfoList = labels.map((l, idx) => ({
    valuation: groups[l].valuation,
    ratio: ratios[idx],
    returnRate: groups[l].principal > 0 ? ((groups[l].diff / groups[l].principal) * 100) : 0
  }));

  const wrapper = $('familyChartInnerWrapper');

  if (tab === 'donut') {
    wrapper.style.width = labels.length > 20 ? `${Math.max(650, labels.length * 14)}px` : '100%';
    wrapper.style.height = labels.length > 15 ? `${Math.max(380, labels.length * 20)}px` : '380px';
  } else if (tab === 'bar') {
    // X축/Y축 상호 변경 (indexAxis: 'y'): 항목이 세로로 배치되므로 높이(세로) 연장으로 상하 스크롤 지원!
    wrapper.style.width = '100%';
    wrapper.style.height = labels.length > 8 ? `${Math.max(380, labels.length * 36)}px` : '380px';
  } else if (tab === 'returnRank') {
    wrapper.style.width = '100%';
    wrapper.style.height = labels.length > 10 ? `${Math.max(380, labels.length * 30)}px` : '380px';
  }

  const layoutPadding = {
    top: 35,
    bottom: tab === 'bar' ? 35 : 20, // X축 90도 수직 글자 표기를 위한 여백 확보
    left: 20,
    right: tab === 'bar' ? 140 : (tab === 'donut' ? 35 : 20)
  };

  if (tab === 'donut') {
    AppState.chartInstances.familyAllocation = new Chart(ctx, {
      type: 'doughnut',
      plugins: [alwaysShowLabelsPlugin],
      data: {
        labels: labels,
        datasets: [{
          label: `${metricKey} 기준 (${rowLabelText})`,
          data: labels.map(l => Math.abs(groups[l].metricVal)),
          backgroundColor: palette.slice(0, Math.max(labels.length, palette.length)),
          ratioList: ratios
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: layoutPadding },
        plugins: {
          legend: {
            position: 'right',
            labels: {
              color: '#cbd5e1',
              font: { size: labels.length > 15 ? 10 : 11 },
              boxWidth: 12
            }
          },
          title: { display: true, text: `가족 합산 [${rowLabelText}]별 [${metricKey}] 비중`, color: '#38bdf8', font: { size: 13 } }
        }
      }
    });
  } else if (tab === 'bar') {
    // 사용자 명확 요구사항:
    // 1) 평가금액 tab의 X축과 Y축 변경 (indexAxis: 'y') -> Y축 항목명, X축 평가금액 수치
    // 2) X축 글자 수직 90도 표기 (maxRotation: 90, minRotation: 90)
    AppState.chartInstances.familyAllocation = new Chart(ctx, {
      type: 'bar',
      plugins: [alwaysShowLabelsPlugin],
      data: {
        labels: labels, // Y축: 항목명
        datasets: [
          { label: '투자원금', data: labels.map(l => groups[l].principal), backgroundColor: 'rgba(148, 163, 184, 0.5)' },
          { 
            label: '평가금액', 
            data: labels.map(l => groups[l].valuation), // X축 수치: 평가금액
            backgroundColor: 'rgba(99, 102, 241, 0.85)', 
            richInfoList: richInfoList // 막대 우측 3-in-1 리치 데이터 라벨 [평가액 (비중%) | 수익률%]
          }
        ]
      },
      options: {
        indexAxis: 'y', // ★ X축과 Y축 상호 변경!
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: layoutPadding },
        plugins: {
          legend: { position: 'top', labels: { color: '#cbd5e1' } },
          title: { display: true, text: `[${rowLabelText}] 평가금액 비교 (내림차순 정렬)`, color: '#38bdf8', font: { size: 13 } }
        },
        scales: {
          x: {
            grid: { color: 'rgba(255, 255, 255, 0.05)' },
            ticks: {
              color: '#94a3b8',
              maxRotation: 90, // ★ X축 글자 수직 90도 직각 표기!
              minRotation: 90,
              font: { size: 11, weight: '600' },
              callback: (v) => formatCurrency(v)
            }
          },
          y: {
            grid: { color: 'rgba(255, 255, 255, 0.05)' },
            ticks: { color: '#cbd5e1', autoSkip: false, font: { size: 11, weight: '600' } }
          }
        }
      }
    });
  } else if (tab === 'returnRank') {
    const rankList = labels.map(l => ({ label: l, returnRate: groups[l].principal > 0 ? ((groups[l].diff / groups[l].principal) * 100) : 0 })).sort((a, b) => b.returnRate - a.returnRate);
    AppState.chartInstances.familyAllocation = new Chart(ctx, {
      type: 'bar',
      plugins: [alwaysShowLabelsPlugin],
      data: {
        labels: rankList.map(r => r.label),
        datasets: [{ label: '수익률 (%)', data: rankList.map(r => r.returnRate), backgroundColor: rankList.map(r => r.returnRate >= 0 ? 'rgba(16, 185, 129, 0.85)' : 'rgba(239, 68, 68, 0.85)') }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: layoutPadding },
        plugins: {
          legend: { display: false },
          title: { display: true, text: `[${rowLabelText}] 그룹별 수익률 랭킹`, color: '#38bdf8', font: { size: 13 } }
        },
        scales: {
          x: {
            ticks: {
              color: '#94a3b8',
              maxRotation: 90, // ★ X축 수치 글자 90도 직각 수직 표기
              minRotation: 90,
              callback: (v) => v + '%'
            }
          },
          y: { ticks: { color: '#cbd5e1' } }
        }
      }
    });
  }
}

function renderPerPersonCharts() {
  const container = $('personChartsGrid');
  container.innerHTML = '';

  Object.keys(AppState.chartInstances.personCharts).forEach(key => {
    if (AppState.chartInstances.personCharts[key]) AppState.chartInstances.personCharts[key].destroy();
  });
  AppState.chartInstances.personCharts = {};

  const data = getSection3PivotDataset();
  const selDate = AppState.sec3.selectedDate;
  const activeData = data.filter(d => (selDate === 'ALL' || d.평가일 === selDate) && !isRedeemed(d.환매여부));

  const names = Array.from(new Set(activeData.map(d => d.성명))).sort();

  const selectedRows = AppState.pivotConfig.rows;
  const metricKey = AppState.pivotConfig.metric || '평가금액';
  const rowLabelText = selectedRows.join(' + ');

  names.forEach(pName => {
    const pItems = activeData.filter(d => d.성명 === pName);
    const pPrincipal = pItems.reduce((s, i) => s + i.투자원금, 0);
    const pValuation = pItems.reduce((s, i) => s + i.평가금액, 0);
    const pDiff = pValuation - pPrincipal;
    const pReturn = pPrincipal > 0 ? ((pDiff / pPrincipal) * 100) : 0;

    const card = document.createElement('div');
    card.className = 'glass-panel person-card';

    card.innerHTML = `
      <div class="person-card-header">
        <span class="person-name"><i class="fa-solid fa-user-gear"></i> ${pName} 자산 분석 (${rowLabelText})</span>
        <span class="badge ${pReturn >= 0 ? '' : 'badge-redeemed'}">${pReturn >= 0 ? '+' : ''}${pReturn.toFixed(2)}%</span>
      </div>
      <div class="flex justify-between items-center text-sm mb-2">
        <span class="text-muted">원금: ${formatCurrency(pPrincipal)}</span>
        <span class="font-bold text-white">${metricKey}: ${formatCurrency(metricKey === '평가금액' ? pValuation : pItems.reduce((s, i) => s + (Number(i[metricKey]) || 0), 0))}</span>
      </div>
      <div class="chart-box person-chart-box">
        <canvas id="personChart_${pName}"></canvas>
      </div>
    `;
    container.appendChild(card);

    const ctx = document.getElementById(`personChart_${pName}`).getContext('2d');
    const groupMap = {};
    let pMetricSum = 0;

    pItems.forEach(i => {
      const gKey = getCombinedGroupKey(i, selectedRows);
      let mVal = i.평가금액;
      if (metricKey === '투자원금') mVal = i.투자원금;
      else if (metricKey === '증감') mVal = i.증감;
      else if (metricKey === '수익률') mVal = i.수익률;

      groupMap[gKey] = (groupMap[gKey] || 0) + mVal;
      pMetricSum += Math.abs(mVal);
    });

    const labels = Object.keys(groupMap);
    const values = Object.values(groupMap);
    const ratios = values.map(v => pMetricSum > 0 ? ((Math.abs(v) / pMetricSum) * 100) : 0);
    const palette = ['#6366f1', '#06b6d4', '#10b981', '#f59e0b', '#ec4899', '#8b5cf6'];

    AppState.chartInstances.personCharts[pName] = new Chart(ctx, {
      type: 'bar',
      plugins: [alwaysShowLabelsPlugin],
      data: {
        labels: labels,
        datasets: [{ label: `${metricKey}`, data: values, backgroundColor: palette.slice(0, labels.length), ratioList: ratios }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: { top: 30, bottom: 40, left: 10, right: 10 } },
        plugins: {
          legend: { display: false },
          title: { display: true, text: `[${rowLabelText}] 조합 분포`, color: '#cbd5e1', font: { size: 11 } }
        },
        scales: {
          x: {
            ticks: {
              color: '#cbd5e1',
              maxRotation: 90, // ★ 성명별 개별 차트 X축 90도 직각 수직 정렬!
              minRotation: 90,
              autoSkip: false,
              font: { size: 11, weight: '600' }
            }
          },
          y: { ticks: { color: '#94a3b8', font: { size: 10 }, callback: (v) => formatCurrency(v) } }
        }
      }
    });
  });
}

function renderRawTable() {
  const tbody = $('rawDataTableBody');
  tbody.innerHTML = '';

  const data = filterDatasetBySection('sec3');
  const selDate = AppState.sec3.selectedDate;
  let evalData = data.filter(d => selDate === 'ALL' || d.평가일 === selDate);

  const map = AppState.columnMap;

  const query = $('tableSearchInput').value.toLowerCase().trim();
  if (query) {
    evalData = evalData.filter(d => 
      String(d[map['투자상품']] || d.투자상품 || '').toLowerCase().includes(query) ||
      String(d[map['운용사']] || d.운용사 || '').toLowerCase().includes(query) ||
      String(d[map['성명']] || d.성명 || '').toLowerCase().includes(query) ||
      String(d[map['비고']] || d.비고 || '').toLowerCase().includes(query) ||
      String(d[map['구분']] || d.구분 || '').toLowerCase().includes(query) ||
      String(d[map['구분2']] || d.구분2 || '').toLowerCase().includes(query) ||
      String(d[map['구분3']] || d.구분3 || '').toLowerCase().includes(query) ||
      String(d[map['국가_상품']] || d.국가_상품 || '').toLowerCase().includes(query) ||
      String(d.환매여부 || '').toLowerCase().includes(query)
    );
  }

  const col = AppState.sortState.column;
  const dir = AppState.sortState.direction === 'asc' ? 1 : -1;

  evalData.sort((a, b) => {
    let valA = a[col] !== undefined ? a[col] : (map[col] ? a[map[col]] : '');
    let valB = b[col] !== undefined ? b[col] : (map[col] ? b[map[col]] : '');
    return (typeof valA === 'string' ? valA.localeCompare(valB) : (valA - valB)) * dir;
  });

  const getVal = (stdKey) => {
    const actualKey = map[stdKey];
    if (actualKey && itemValExists(item, actualKey)) return item[actualKey];
    return item[stdKey] !== undefined ? item[stdKey] : '';
  };

  function itemValExists(item, key) {
    return item[key] !== undefined && item[key] !== null;
  }

  evalData.forEach(item => {
    const tr = document.createElement('tr');
    const isRedeemedItem = isRedeemed(item.환매여부);
    if (isRedeemedItem) tr.className = 'redeemed-row';

    const retClass = item.수익률 >= 0 ? 'positive-text' : 'negative-text';
    const redemptionBadge = isRedeemedItem 
      ? `<span class="badge badge-redeemed"><i class="fa-solid fa-ban"></i> 환매 (계산제외)</span>` 
      : item.환매여부;

    const rowVal = (stdKey) => {
      const actualKey = map[stdKey];
      if (actualKey && itemValExists(item, actualKey)) return item[actualKey];
      return item[stdKey] !== undefined ? item[stdKey] : '';
    };

    tr.innerHTML = `
      <td>${rowVal('평가일')}</td>
      <td>${rowVal('운용사')}</td>
      <td>${rowVal('성명')}</td>
      <td><span class="badge">${rowVal('구분')}</span></td>
      <td>${rowVal('구분2')}</td>
      <td>${rowVal('구분3')}</td>
      <td>${rowVal('국내외')}</td>
      <td>${rowVal('국가_상품')}</td>
      <td class="font-bold">${rowVal('투자상품')}</td>
      <td class="text-right num-col">${formatCurrency(item.투자원금)}</td>
      <td class="text-right num-col font-bold">${formatCurrency(item.평가금액)}</td>
      <td class="text-right num-col ${item.증감 >= 0 ? 'positive-text' : 'negative-text'}">${item.증감 >= 0 ? '+' : ''}${formatCurrency(item.증감)}</td>
      <td class="text-right num-col ${retClass}">${item.수익률 >= 0 ? '+' : ''}${item.수익률.toFixed(2)}%</td>
      <td>${redemptionBadge}</td>
      <td class="text-muted">${rowVal('비고')}</td>
    `;
    tbody.appendChild(tr);
  });

  const activeCount = evalData.filter(d => !isRedeemed(d.환매여부)).length;
  const redeemedCount = evalData.length - activeCount;

  $('tableRecordCount').textContent = `조회 데이터: 총 ${evalData.length}건 (운용 자산: ${activeCount}건, '환매' 제외: ${redeemedCount}건)`;
}

function exportToExcel() {
  const data = filterDatasetBySection('sec3');
  const selDate = AppState.sec3.selectedDate;
  const evalData = data.filter(d => selDate === 'ALL' || d.평가일 === selDate);

  if (evalData.length === 0) {
    showToast('error', '내보내기 실패', '내보낼 자산 데이터가 존재하지 않습니다.');
    return;
  }

  const map = AppState.columnMap;

  const exportRows = evalData.map(d => {
    const rowObj = {};
    rowObj[map['평가일'] || '평가일'] = d.평가일;
    rowObj[map['운용사'] || '운용사'] = d.운용사;
    rowObj[map['성명'] || '성명'] = d.성명;
    rowObj[map['구분'] || '구분'] = d.구분;
    rowObj[map['구분2'] || '구분2'] = d.구분2;
    rowObj[map['구분3'] || '구분3'] = d.구분3;
    rowObj[map['국내외'] || '국내외'] = d.국내외;
    rowObj[map['국가_상품'] || '국가_상품'] = d.국가_상품;
    rowObj[map['투자상품'] || '투자상품'] = d.투자상품;
    rowObj[map['투자원금'] || '투자원금'] = d.투자원금;
    rowObj[map['평가금액'] || '평가금액'] = d.평가금액;
    rowObj[map['증감'] || '증감'] = d.증감;
    rowObj[map['수익률'] ? `${map['수익률']}(%)` : '수익률(%)'] = d.수익률;
    rowObj[map['환매여부'] || '환매여부'] = d.환매여부;
    rowObj['집계제외여부'] = isRedeemed(d.환매여부) ? '계산제외(환매)' : '포함';
    rowObj[map['비고'] || '비고'] = d.비고;
    return rowObj;
  });

  const worksheet = XLSX.utils.json_to_sheet(exportRows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, '자산분석리포트');

  const todayStr = new Date().toISOString().split('T')[0];
  XLSX.writeFile(workbook, `자산관리_다차원분석리포트_${todayStr}.xlsx`);
  showToast('success', '엑셀 리포트 내보내기 완료!');
}

function formatCurrency(num) {
  if (num === null || num === undefined || isNaN(num)) return '₩0';
  return '₩' + Math.round(num).toLocaleString('ko-KR');
}

function formatCurrencyCompact(num) {
  if (num === null || num === undefined || isNaN(num)) return '₩0';
  const abs = Math.abs(num);
  if (abs >= 100000000) {
    return '₩' + (num / 100000000).toFixed(2) + '억';
  } else if (abs >= 10000) {
    return '₩' + (num / 10000).toFixed(0) + '만';
  }
  return '₩' + Math.round(num).toLocaleString('ko-KR');
}

function hexToRgba(hex, alpha) {
  let c = hex.replace('#', '');
  if (c.length === 3) c = c.split('').map(x => x + x).join('');
  const num = parseInt(c, 16);
  return `rgba(${(num >> 16) & 255}, ${(num >> 8) & 255}, ${num & 255}, ${alpha})`;
}
