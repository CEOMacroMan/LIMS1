// Functions to infer form schema and render the form

function detectType(name, sampleValues, validation) {
  const lower = name.toLowerCase();
  const required = /\*|\[required\]/i.test(name);
  let type = 'text';
  const info = [];

  const hasValidationList = validation && validation.rule && validation.rule.inCellDropDown;
  if (hasValidationList) {
    type = 'select';
    info.push('Liste aus Datenvalidierung');
  } else if (/notiz|beschreibung|kommentar/i.test(lower)) {
    type = 'textarea';
  } else if (/foto|bild|image|attachment|datei/i.test(lower)) {
    type = 'file';
  } else {
    const values = sampleValues.filter(v => v !== null && v !== undefined);
    const allNumbers = values.every(v => !isNaN(v));
    const allInts = allNumbers && values.every(v => Number.isInteger(Number(v)));
    const allDates = values.every(v => /\d{2}\.\d{2}\.\d{4}/.test(v));
    const allBools = values.every(v => /^(ja|nein|true|false|bool|check)$/i.test(String(v)));

    if (allDates || /datum/i.test(lower)) {
      type = 'date';
    } else if (allBools) {
      type = 'checkbox';
    } else if (allNumbers) {
      type = 'number';
      if (!allInts) info.push('dezimal');
    }
  }

  return { name, type, required, validation, info };
}

export function inferSchema(table) {
  const schema = table.columns.map((col, idx) => {
    const sampleValues = table.rows.map(r => r[idx]);
    return detectType(col.name, sampleValues, col.validation);
  });
  return schema;
}

export function generateForm(schema, container) {
  const form = document.createElement('form');
  schema.forEach(field => {
    const wrapper = document.createElement('div');
    const label = document.createElement('label');
    label.textContent = field.name;
    let input;
    switch (field.type) {
      case 'date':
        input = document.createElement('input');
        input.type = 'date';
        input.placeholder = 'dd.MM.yyyy';
        break;
      case 'number':
        input = document.createElement('input');
        input.type = 'number';
        if (field.info.includes('dezimal')) input.step = '0.01';
        break;
      case 'checkbox':
        input = document.createElement('input');
        input.type = 'checkbox';
        break;
      case 'select':
        input = document.createElement('select');
        const opts = field.validation && field.validation.formula1;
        if (opts) {
          opts.replace(/[=\"]?/g, '').split(',').forEach(o => {
            const option = document.createElement('option');
            option.value = o.trim();
            option.textContent = o.trim();
            input.appendChild(option);
          });
        }
        break;
      case 'textarea':
        input = document.createElement('textarea');
        break;
      case 'file':
        input = document.createElement('input');
        input.type = 'file';
        input.multiple = true;
        break;
      default:
        input = document.createElement('input');
        input.type = 'text';
    }
    if (field.required) input.required = true;
    wrapper.appendChild(label);
    wrapper.appendChild(input);
    form.appendChild(wrapper);
  });
  container.appendChild(form);
}
