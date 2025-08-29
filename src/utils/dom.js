export function el(tag, props = {}, ...children) {
  const element = document.createElement(tag);
  Object.assign(element, props);
  for (const child of children) element.appendChild(child);
  return element;
}

export function clear(node) {
  while (node.firstChild) node.removeChild(node.firstChild);
}

export function fragment() {
  return document.createDocumentFragment();
}

export function setStatus(msg) {
  const el = document.getElementById('status');
  if (el) el.textContent = msg || '';
}
