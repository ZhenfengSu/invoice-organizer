const folderInput = document.querySelector('#folder');
const fileInput = document.querySelector('#files');
const dropzone = document.querySelector('#dropzone');
const selected = document.querySelector('#selected');
const run = document.querySelector('#run');
const status = document.querySelector('#status');
const results = document.querySelector('#results');
let chosen = [];

function setFiles(files) {
  chosen = [...files].filter(file => /\.(pdf|zip)$/i.test(file.name));
  selected.textContent = chosen.length ? `已选择 ${chosen.length} 个文件 / 压缩包` : '没有找到 PDF 或 ZIP 文件';
  results.hidden = true;
}

folderInput.addEventListener('change', event => setFiles(event.target.files));
fileInput.addEventListener('change', event => setFiles(event.target.files));
for (const name of ['dragenter', 'dragover']) {
  dropzone.addEventListener(name, event => { event.preventDefault(); dropzone.classList.add('dragging'); });
}
for (const name of ['dragleave', 'drop']) {
  dropzone.addEventListener(name, event => { event.preventDefault(); dropzone.classList.remove('dragging'); });
}
dropzone.addEventListener('drop', event => setFiles(event.dataTransfer.files));

function stat(label, value) {
  const card = document.createElement('div');
  card.className = 'stat';
  const number = document.createElement('strong');
  number.textContent = value;
  const caption = document.createElement('span');
  caption.textContent = label;
  card.append(number, caption);
  return card;
}

run.addEventListener('click', async () => {
  if (!chosen.length) { status.textContent = '请先选择文件夹、PDF 或 ZIP'; return; }
  run.disabled = true;
  status.textContent = '正在识别与生成报告…';
  results.hidden = true;
  const form = new FormData();
  for (const file of chosen) form.append('files', file, file.webkitRelativePath || file.name);
  form.append('buyer_names', document.querySelector('#buyer-names').value);
  form.append('tax_ids', document.querySelector('#tax-ids').value);
  form.append('use_llm', document.querySelector('#use-llm').checked);
  try {
    const response = await fetch('/api/process', { method: 'POST', body: form });
    const data = await response.json();
    if (!response.ok) throw new Error(data.detail || '处理失败');
    const summary = data.summary;
    const stats = document.querySelector('#stats');
    stats.replaceChildren(
      stat('已处理 PDF', summary.processed),
      stat('国内有效发票', summary.domestic),
      stat('国内合计 · CNY', `¥ ${summary.domestic_total}`),
      stat('海外有效收据', summary.overseas),
      stat('待检查', summary.errors),
    );
    for (const [currency, amount] of Object.entries(summary.overseas_totals)) stats.append(stat(`海外合计 · ${currency}`, amount));
    const errors = document.querySelector('#errors');
    errors.replaceChildren();
    for (const item of summary.error_details.slice(0, 20)) {
      const row = document.createElement('p');
      row.textContent = `${item.file} — ${item.reasons.join('；')}`;
      errors.append(row);
    }
    if (summary.error_details.length > 20) {
      const row = document.createElement('p');
      row.textContent = `还有 ${summary.error_details.length - 20} 项，请查看错误报告。`;
      errors.append(row);
    }
    document.querySelector('#download').href = `/api/download/${data.job_id}`;
    results.hidden = false;
    status.textContent = `完成 · ${summary.processed} 份 PDF`;
    results.scrollIntoView({ behavior: 'smooth', block: 'start' });
  } catch (error) {
    status.textContent = error.message;
  } finally {
    run.disabled = false;
  }
});
