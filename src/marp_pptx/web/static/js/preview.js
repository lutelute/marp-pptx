const range = document.getElementById('fs-range');
const val = document.getElementById('fs-val');
range.addEventListener('input', () => { val.textContent = parseFloat(range.value).toFixed(2); });
