export function el(tag, attrs = {}, children = []) {
  const node = document.createElement(tag);
  Object.entries(attrs).forEach(([k, v]) => node.setAttribute(k, v));
  children.forEach(ch => node.appendChild(ch));
  return node;
}

export function clear(node) {
  while (node.firstChild) node.removeChild(node.firstChild);
}

export function fragment(children = []) {
  const f = document.createDocumentFragment();
  children.forEach(ch => f.appendChild(ch));
  return f;
}

export function setStatus(msg = '') {
  const el = document.getElementById('status');
  if (el) el.textContent = msg;
}
