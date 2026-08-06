export default function reconcileFlow(component) {
const root = component.parentElement;
const { data: serverData, setTriggerValue } = component;
const state = root.__poowardState || {
  bundle: null,
  module: 'order',
  basis: 'rmb',
  customer: null,
  lineKey: null,
  documentSignature: null,
  lastResponseId: null,
};
root.__poowardState = state;
const $ = selector => root.querySelector(selector);

const moduleUi = {
  order: {
    badge: '接单差异核对',
    eyebrow: 'ORDER RECONCILIATION',
    resultEyebrow: 'ORDER',
    title: '接单差异核对',
    systemDetail: '系统接单数据',
    uploadTitle: '上传接单核对三份文件',
    action: '开始接单核对',
    progress: '正在读取接单模块三份文件，并生成人民币与原币结果…',
    rmbRule: '系统金额 = 接单金额(RMB) + 出货运费(RMB)；文控金额 = “接单”子表的 VAT PRICE。',
    originalRule: '系统金额 = 交易金额 + 出货运费(原币)；文控金额 = “接单”子表的 TP-CPO。',
  },
  shipment: {
    badge: '出货差异核对',
    eyebrow: 'SHIPMENT RECONCILIATION',
    resultEyebrow: 'SHIPMENT',
    title: '出货差异核对',
    systemDetail: '系统出货数据',
    uploadTitle: '上传出货核对三份文件',
    action: '开始出货核对',
    progress: '正在读取出货模块三份文件，并生成人民币与原币结果…',
    rmbRule: '系统金额 = 出货金额(RMB) + 出货运费(RMB)；文控金额 = “出货”子表的 VAT PRICE。',
    originalRule: '系统金额 = 出货金额 + 出货运费(原币)；文控金额 = “出货”子表的 TP-CPO。',
  },
};

const basisUi = {
  rmb: {
    badge: '人民币口径 · RMB',
    tolerance: '差异容差（元）',
    orderAmount: '读取接单金额(RMB)',
    orderFreight: '追加出货运费(RMB)',
    shipmentAmount: '读取出货金额(RMB)',
    shipmentFreight: '追加出货运费(RMB)',
  },
  original: {
    badge: '原币口径 · ORIGINAL',
    tolerance: '差异容差（原币单位）',
    orderAmount: '读取交易金额',
    orderFreight: '追加出货运费(原币)',
    shipmentAmount: '读取出货金额',
    shipmentFreight: '追加出货运费(原币)',
  },
};

const currentData = () => state.bundle?.modules?.[state.module]?.[state.basis] || null;

function updateContext({ render = true } = {}) {
  const module = moduleUi[state.module];
  const basis = basisUi[state.basis];
  $('#moduleBadge').textContent = module.badge;
  $('#basisBadge').textContent = basis.badge;
  $('#moduleEyebrow').textContent = module.eyebrow;
  $('#moduleTitle').textContent = module.title;
  $('#uploadTitle').textContent = module.uploadTitle;
  $('#analyzeButton').querySelector('span').textContent = module.action;
  $('#progressText').textContent = module.progress;
  $('#basisRule').textContent = state.basis === 'rmb' ? module.rmbRule : module.originalRule;
  $('#toleranceLabel').textContent = basis.tolerance;
  $('#orderAmountRule').textContent = basis.orderAmount;
  $('#orderFreightRule').textContent = basis.orderFreight;
  $('#shipmentAmountRule').textContent = basis.shipmentAmount;
  $('#shipmentFreightRule').textContent = basis.shipmentFreight;
  root.dataset.basis = state.basis;
  root.querySelectorAll('[data-module]').forEach(button => button.classList.toggle('active', button.dataset.module === state.module));
  root.querySelectorAll('[data-source-module]').forEach(group => group.classList.toggle('active', group.dataset.sourceModule === state.module));
  root.querySelectorAll('[data-source-module] input[type="file"]').forEach(input => {
    input.required = input.closest('[data-source-module]').dataset.sourceModule === state.module;
  });
  if (render && currentData()) {
    renderResults({ scroll: false });
  } else if (render) {
    $('#results').classList.add('hidden');
  }
}

root.querySelectorAll('[data-module]').forEach(button => {
  button.onclick = () => {
    state.module = button.dataset.module;
    state.customer = null;
    state.lineKey = null;
    updateContext();
    if (!currentData()) $('#uploadPanel').scrollIntoView({behavior: 'smooth', block: 'start'});
  };
});

root.querySelectorAll('input[name="basis"]').forEach(input => {
  input.onchange = () => {
    state.basis = input.value;
    state.customer = null;
    state.lineKey = null;
    updateContext();
  };
});

const isExcelFile = file => /\.(xlsx|xlsm)$/i.test(file?.name || '');

function updateFileCard(input) {
  const card = input.closest('.drop-card');
  const filename = card.querySelector('.filename');
  const file = input.files[0];
  filename.textContent = file?.name || '尚未选择';
  filename.title = file?.name || '';
  filename.setAttribute('aria-live', 'polite');
  card.classList.toggle('selected', Boolean(file));
}

function clearDragStates() {
  root.querySelectorAll('.drop-card.drag-over').forEach(card => card.classList.remove('drag-over'));
}

root.ondragover = event => {
  if ([...(event.dataTransfer?.types || [])].includes('Files')) event.preventDefault();
};

root.ondrop = event => {
  if ([...(event.dataTransfer?.types || [])].includes('Files')) event.preventDefault();
  clearDragStates();
};

root.onmouseleave = clearDragStates;

root.querySelectorAll('input[type="file"]').forEach(input => {
  const card = input.closest('.drop-card');
  let dragDepth = 0;

  input.onchange = () => updateFileCard(input);

  card.ondragenter = event => {
    if (![...(event.dataTransfer?.types || [])].includes('Files')) return;
    event.preventDefault();
    dragDepth += 1;
    card.classList.add('drag-over');
  };

  card.ondragover = event => {
    if (![...(event.dataTransfer?.types || [])].includes('Files')) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = 'copy';
  };

  card.ondragleave = () => {
    dragDepth = Math.max(0, dragDepth - 1);
    if (!dragDepth) card.classList.remove('drag-over');
  };

  card.ondrop = event => {
    event.preventDefault();
    event.stopPropagation();
    dragDepth = 0;
    card.classList.remove('drag-over');
    const files = [...(event.dataTransfer?.files || [])];
    if (!files.length) return;
    if (files.length > 1) {
      showError('每个上传框一次只能放入一个 Excel 文件。');
      return;
    }
    if (!isExcelFile(files[0])) {
      showError('文件格式不支持，请上传 .xlsx 或 .xlsm 文件。');
      return;
    }
    if (typeof DataTransfer === 'undefined') {
      showError('当前浏览器不支持拖拽上传，请点击上传框选择文件。');
      return;
    }
    const transfer = new DataTransfer();
    transfer.items.add(files[0]);
    input.files = transfer.files;
    $('#errorBox').classList.add('hidden');
    input.dispatchEvent(new Event('change', { bubbles: true }));
  };
});

const escapeHtml = value => String(value ?? '').replace(/[&<>'"]/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[char]));
const number = value => Number(value || 0).toLocaleString('zh-CN', {
  minimumFractionDigits: 2,
  maximumFractionDigits: state.basis === 'original' ? 4 : 2,
});
const integer = value => Number(value || 0).toLocaleString('zh-CN');
const moneyClass = value => Number(value) < 0 ? 'negative' : Number(value) > 0 ? 'positive' : '';
const statusClass = value => ({'有差异':'difference','仅文控表':'doc-only','仅系统表':'sys-only','一致':'ok'}[value] || 'difference');
const statusTag = value => `<span class="status ${statusClass(value)}">${escapeHtml(value)}</span>`;

function showError(message) {
  $('#errorBox').textContent = message;
  $('#errorBox').classList.remove('hidden');
}

async function encodeFile(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  let binary = '';
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return {
    name: file.name,
    size: file.size,
    last_modified: file.lastModified,
    data: btoa(binary),
  };
}

$('#uploadForm').onsubmit = async event => {
  event.preventDefault();
  const form = event.currentTarget;
  const button = $('#analyzeButton');
  $('#errorBox').classList.add('hidden');

  const documentFile = form.elements.document.files[0];
  const amountFile = form.elements[`${state.module}_amount`].files[0];
  const freightFile = form.elements[`${state.module}_freight`].files[0];
  if (!documentFile || !amountFile || !freightFile) {
    showError('请完整上传文控登记表、金额明细和运费明细三份文件。');
    return;
  }
  if (![documentFile, amountFile, freightFile].every(isExcelFile)) {
    showError('文件格式不支持，请上传 .xlsx 或 .xlsm 文件。');
    return;
  }

  $('#progress').classList.remove('hidden');
  button.disabled = true;
  $('#results').classList.add('hidden');
  try {
    const signature = `${documentFile.name}|${documentFile.size}|${documentFile.lastModified}`;
    if (state.documentSignature && state.documentSignature !== signature) state.bundle = null;
    state.documentSignature = signature;
    const requestId = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
    setTriggerValue('analyze', {
      request_id: requestId,
      module: state.module,
      tolerance: Number(form.elements.tolerance.value || 1),
      files: {
        document: await encodeFile(documentFile),
        amount: await encodeFile(amountFile),
        freight: await encodeFile(freightFile),
      },
    });
  } catch (error) {
    showError(error?.message || '文件读取失败，请重新选择后再试。');
    $('#progress').classList.add('hidden');
    button.disabled = false;
  }
};

function renderResults({ scroll = false } = {}) {
  const data = currentData();
  if (!data) return;
  const s = data.statistics;
  const module = moduleUi[state.module];
  const prefix = state.basis === 'rmb' ? '¥ ' : '';
  const suffix = state.basis === 'rmb' ? '' : '（原币）';
  $('#resultEyebrow').textContent = `${module.resultEyebrow} · ${state.basis === 'rmb' ? 'RMB' : 'ORIGINAL'}`;
  $('#resultTitle').textContent = `${module.title}结果`;
  $('#systemDetailTitle').textContent = module.systemDetail;
  $('#downloadButton').textContent = `下载${module.title} Excel结果`;
  const cachedModules = Object.keys(state.bundle?.modules || {});
  $('#cacheStatusText').textContent = cachedModules.length > 1
    ? '接单、出货模块及各自双口径结果均已缓存，可直接切换。'
    : `${module.title}的人民币与原币结果均已缓存，可直接切换口径。`;
  const cards = [
    [`系统总额${suffix}`, `${prefix}${number(s['系统总额'])}`, 'system-total'],
    [`文控总额${suffix}`, `${prefix}${number(s['文控总额'])}`, 'document-total'],
    [`总差异（系统 − 文控）${suffix}`, `${prefix}${number(s['总差异'])}`, 'accent'],
    ['差异客户 / 流水号', `${integer(s['差异客户数'])} / ${integer(s['差异流水号数'])}`],
  ];
  $('#kpis').innerHTML = cards.map(([label, value, cls='']) => `<div class="kpi ${cls}"><span>${label}</span><strong>${value}</strong></div>`).join('');
  const warning = $('#warningBox');
  if (data.warnings.length) {
    warning.innerHTML = `<details><summary>数据提示（${data.warnings.length}）</summary><ul>${data.warnings.map(item => `<li>${escapeHtml(item)}</li>`).join('')}</ul></details>`;
    warning.classList.remove('hidden');
  } else {
    warning.classList.add('hidden');
  }
  $('#downloadButton').onclick = () => {
    const binary = atob(data.download_base64 || '');
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    const url = URL.createObjectURL(new Blob(
      [bytes],
      { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
    ));
    const link = document.createElement('a');
    link.href = url;
    link.download = data.download_name || '核对结果.xlsx';
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };
  $('#customerSearch').value = '';
  state.customer = null;
  state.lineKey = null;
  renderCustomers();
  $('#linePanel').classList.add('hidden');
  $('#detailPanel').classList.add('hidden');
  $('#results').classList.remove('hidden');
  if (scroll) $('#results').scrollIntoView({behavior: 'smooth', block: 'start'});
}

function renderCustomers() {
  const data = currentData();
  if (!data) return;
  const query = $('#customerSearch').value.trim().toUpperCase();
  const rows = data.customer_differences.filter(row => Math.abs(Number(row['差异'] || 0)) > 1e-9 && (!query || row['客户代码'].includes(query)));
  const labels = data.labels;
  const headers = ['客户代码', labels.main, labels.freight, labels.system, labels.document, labels.difference, '状态'];
  const body = rows.map(row => `<tr class="${state.customer === row['客户代码'] ? 'selected-row' : ''}">
    <td><button class="code-button" data-customer="${escapeHtml(row['客户代码'])}">${escapeHtml(row['客户代码'])}</button></td>
    <td>${number(row['主金额'])}</td><td>${number(row['运费'])}</td><td class="compare-system">${number(row['系统金额'])}</td><td class="compare-document">${number(row['文控金额'])}</td>
    <td class="money ${moneyClass(row['差异'])}">${number(row['差异'])}</td><td>${statusTag(row['状态'])}</td></tr>`).join('');
  $('#customerTable').innerHTML = `<thead><tr>${headers.map((item, index) => `<th class="${index === 3 ? 'compare-system' : index === 4 ? 'compare-document' : ''}">${escapeHtml(item)}</th>`).join('')}</tr></thead><tbody>${body || '<tr><td colspan="7" class="empty">没有符合条件的差异客户</td></tr>'}</tbody>`;
  root.querySelectorAll('[data-customer]').forEach(button => button.addEventListener('click', () => selectCustomer(button.dataset.customer)));
}

$('#customerSearch').oninput = renderCustomers;

function selectCustomer(customer) {
  const data = currentData();
  state.customer = customer;
  state.lineKey = null;
  renderCustomers();
  const lines = data.line_differences.filter(row => row['客户代码'] === customer && Math.abs(Number(row['差异'] || 0)) > 1e-9);
  $('#lineCustomer').textContent = customer;
  $('#lineCount').textContent = `${lines.length} 条差异`;
  const labels = data.labels;
  const headers = ['订单流水号', labels.main, labels.freight, labels.system, labels.document, labels.difference, '状态'];
  const body = lines.map(row => `<tr>
    <td><button class="code-button" data-key="${escapeHtml(row['匹配键'])}">${escapeHtml(row['订单流水号'])}</button></td>
    <td>${number(row['主金额'])}</td><td>${number(row['运费'])}</td><td class="compare-system">${number(row['系统金额'])}</td><td class="compare-document">${number(row['文控金额'])}</td>
    <td class="money ${moneyClass(row['差异'])}">${number(row['差异'])}</td><td>${statusTag(row['状态'])}</td></tr>`).join('');
  $('#lineTable').innerHTML = `<thead><tr>${headers.map((item, index) => `<th class="${index === 3 ? 'compare-system' : index === 4 ? 'compare-document' : ''}">${escapeHtml(item)}</th>`).join('')}</tr></thead><tbody>${body || '<tr><td colspan="7" class="empty">该客户没有超过容差的流水号差异</td></tr>'}</tbody>`;
  root.querySelectorAll('[data-key]').forEach(button => button.addEventListener('click', () => selectLine(button.dataset.key)));
  $('#linePanel').classList.remove('hidden');
  $('#detailPanel').classList.add('hidden');
  $('#linePanel').scrollIntoView({behavior: 'smooth', block: 'start'});
}

const documentPreferred = ['DAY','BIL','CO','CPO NO','JO','PART-DWG','QTY','CURR','UP-CPO','TP-CPO','EX-CH','VAT  PRICE','VAT PRICE','INVOICE','匹配方式','关联流水号'];
const orderPreferred = ['数据来源','接单时间','客户代码','订单流水号','客户PO','零件名称','订单数量','单价','交易金额','交易币种','接单汇率','接单金额(RMB)','出货运费(原币)','出货运费(RMB)','出货装箱单号','实际出厂日期','出货通知单号'];
const shipmentPreferred = ['数据来源','出货日期','实际出货日期','实际出厂日期','客户代码','订单流水号','客户PO','零件名称','出货数量','订单数量','出货金额','实际出货金额','交易币种','出货金额(RMB)','实际出货金额(RMB)','出货运费(原币)','出货运费(RMB)','出货装箱单号','出货通知单号'];
const internalFields = new Set(['客户代码_标准','订单流水号_标准','匹配键','来源表','来源行号']);

function orderedLabels(row, preferred) {
  const available = Object.keys(row).filter(label => !internalFields.has(label) && row[label] !== null && row[label] !== '' && row[label] !== undefined);
  const first = preferred.filter(label => available.includes(label));
  return [...first, ...available.filter(label => !first.includes(label))];
}

function rawCard(row, preferred, side) {
  const fields = orderedLabels(row, preferred)
    .map(label => `<div class="field"><label>${escapeHtml(label)}</label><span title="${escapeHtml(row[label])}">${escapeHtml(row[label])}</span></div>`).join('');
  const sourceType = side === 'sys' ? row['数据来源'] : row['匹配方式'];
  return `<div class="raw-card"><div class="raw-title"><span>${escapeHtml(row['来源表'] || '')} · 第 ${escapeHtml(row['来源行号'] || '')} 行</span><span>${escapeHtml(sourceType || '')}</span></div><div class="field-grid">${fields}</div></div>`;
}

function selectLine(key) {
  const data = currentData();
  state.lineKey = key;
  const line = data.line_differences.find(row => row['客户代码'] === state.customer && row['匹配键'] === key);
  if (!line) return;
  const docs = data.document_rows.filter(row => row['客户代码_标准'] === state.customer && row['匹配键'] === key);
  const systems = data.system_rows.filter(row => row['客户代码_标准'] === state.customer && row['匹配键'] === key);
  $('#detailOrder').textContent = line['订单流水号'];
  const prefix = state.basis === 'rmb' ? '¥ ' : '';
  $('#detailStatus').innerHTML = `${statusTag(line['状态'])}　${escapeHtml(data.labels.difference)} ${prefix}${number(line['差异'])}`;
  $('#docCount').textContent = `${docs.length} 行`;
  $('#sysCount').textContent = `${systems.length} 行`;
  $('#docDetails').innerHTML = docs.length ? docs.map(row => rawCard(row, documentPreferred, 'doc')).join('') : '<div class="empty">文控登记表无对应行</div>';
  const preferred = state.module === 'order' ? orderPreferred : shipmentPreferred;
  $('#sysDetails').innerHTML = systems.length ? systems.map(row => rawCard(row, preferred, 'sys')).join('') : '<div class="empty">系统明细无对应行</div>';
  $('#detailPanel').classList.remove('hidden');
  $('#detailPanel').scrollIntoView({behavior: 'smooth', block: 'start'});
}

updateContext({ render: false });

const response = serverData?.response;
if (response && response.request_id && response.request_id !== state.lastResponseId) {
  state.lastResponseId = response.request_id;
  $('#progress').classList.add('hidden');
  $('#analyzeButton').disabled = false;
  if (response.status === 'ok') {
    state.module = response.module || state.module;
    state.bundle = {
      ...state.bundle,
      modules: {
        ...(state.bundle?.modules || {}),
        ...(response.modules || {}),
      },
    };
    state.customer = null;
    state.lineKey = null;
    $('#errorBox').classList.add('hidden');
    updateContext({ render: false });
    renderResults({ scroll: true });
  } else {
    showError(response.error || '核对失败，请检查文件后重试。');
  }
}
}
