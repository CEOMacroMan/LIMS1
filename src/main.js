import { fetchTable } from './sharepointClient.js';
import { generateForm, inferSchema } from './formGenerator.js';

async function init() {
  const errorBanner = document.getElementById('error');
  const formContainer = document.getElementById('form-container');
  const preview = document.getElementById('preview');

  async function load() {
    errorBanner.classList.add('hidden');
    formContainer.innerHTML = '';
    preview.innerHTML = '';
    try {
      const configResp = await fetch('../config.json');
      const config = await configResp.json();
      const table = await fetchTable(config);
      const schema = inferSchema(table);
      generateForm(schema, formContainer);
      renderPreview(schema, preview);
    } catch (err) {
      showError(err);
    }
  }

  function showError(err) {
    errorBanner.textContent = `Fehler beim Laden: ${err.message}`;
    const btn = document.createElement('button');
    btn.textContent = 'Erneut versuchen';
    btn.onclick = load;
    errorBanner.appendChild(document.createElement('br'));
    errorBanner.appendChild(btn);
    errorBanner.classList.remove('hidden');
  }

  function renderPreview(schema, el) {
    const pre = document.createElement('pre');
    pre.textContent = JSON.stringify(schema, null, 2);
    el.appendChild(pre);
  }

  await load();
}

init();
