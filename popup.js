function scrapeOrders() {
  const rows = [];
  const orderItems = document.querySelectorAll('.order-item');

  orderItems.forEach((order) => {
    // Date
    const dateEl = order.querySelector('.order-item-header-right-info div');
    let date = '';
    if (dateEl) {
      date = dateEl.textContent.replace('Date:', '').trim();
    }

    // Order Number (Ref. Number)
    let orderNumber = '';
    const infoDivs = order.querySelectorAll('.order-item-header-right-info div');
    infoDivs.forEach((div) => {
      if (div.textContent.includes('Ref. Number')) {
        orderNumber = div.textContent.replace('Ref. Number:', '').replace('Copy', '').trim();
      }
    });

    // Status
    const statusEl = order.querySelector('.order-item-header-status-text');
    const status = statusEl ? statusEl.textContent.trim() : '';

    // Order total amount
    let currency = '';
    let amount = '';
    const totalEl = order.querySelector('[data-pl="order_item_content_price_total"]');
    if (totalEl) {
      const priceWrap = totalEl.querySelector('.es--wrap--1Hlfkoj');
      if (priceWrap) {
        const fullText = priceWrap.textContent.replace(/\s+/g, ' ').trim();
        const match = fullText.match(/^([^\d]*)([\d.,]*)$/);
        if (match) {
          currency = match[1].trim();
          amount = match[2].trim();
        } else {
          amount = fullText;
        }
      }
    }

    // Each product line (an order can contain multiple products)
    const contentBodies = order.querySelectorAll('.order-item-content-body');
    if (contentBodies.length === 0) {
      rows.push({ date, orderNumber, status, productName: '', currency, amount, quantity: '' });
      return;
    }

    contentBodies.forEach((body) => {
      const nameEl = body.querySelector('.order-item-content-info-name span');
      const productName = nameEl ? nameEl.getAttribute('title') || nameEl.textContent.trim() : '';

      const numberWrap = body.querySelector('.order-item-content-info-number');
      let quantity = '';
      if (numberWrap) {
        const qtyEl = numberWrap.querySelector('.order-item-content-info-number-quantity');
        if (qtyEl) {
          quantity = qtyEl.textContent.trim();
        }
      }

      rows.push({ date, orderNumber, status, productName, currency, amount, quantity });
    });
  });

  return rows;
}

function toCSV(rows) {
  const headers = ['Date', 'Order Number', 'Status', 'Product Name', 'Currency', 'Amount', 'Quantity'];
  const escape = (val) => {
    const str = String(val ?? '');
    if (/[",\n]/.test(str)) {
      return '"' + str.replace(/"/g, '""') + '"';
    }
    return str;
  };

  const lines = [headers.map(escape).join(',')];
  rows.forEach((r) => {
    const orderNumberCell = '="' + r.orderNumber + '"';
    lines.push([r.date, orderNumberCell, r.status, r.productName, r.currency, r.amount, r.quantity].map(escape).join(','));
  });
  return lines.join('\n');
}

// Toggle format option highlight
document.querySelectorAll('.format-option').forEach((opt) => {
  opt.addEventListener('click', () => {
    document.querySelectorAll('.format-option').forEach((o) => o.classList.remove('selected'));
    opt.classList.add('selected');
    opt.querySelector('input').checked = true;
  });
});

function toXLSX(rows) {
  const headers = ['Date', 'Order Number', 'Status', 'Product Name', 'Currency', 'Amount', 'Quantity'];
  const escape = (val) => String(val ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  let table = '<table><tr>' + headers.map((h) => `<th>${escape(h)}</th>`).join('') + '</tr>';
  rows.forEach((r) => {
    const cells = [r.date, r.orderNumber, r.status, r.productName, r.currency, r.amount, r.quantity];
    table += '<tr>' + cells.map((v, i) => {
      const style = i === 1 ? ' style="mso-number-format:\'\\@\';"' : '';
      return `<td${style}>${escape(v)}</td>`;
    }).join('') + '</tr>';
  });
  table += '</table>';

  return `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="utf-8"></head><body>${table}</body></html>`;
}

function filterRows(rows, dateFrom, dateTo, status) {
  return rows.filter((r) => {
    if (status && status !== 'all' && r.status.toLowerCase() !== status.toLowerCase().replace('-', ' ')) {
      return false;
    }
    if (dateFrom || dateTo) {
      const rowDate = new Date(r.date);
      if (!isNaN(rowDate)) {
        if (dateFrom && rowDate < new Date(dateFrom)) return false;
        if (dateTo && rowDate > new Date(dateTo)) return false;
      }
    }
    return true;
  });
}

document.getElementById('exportBtn').addEventListener('click', async () => {
  const statusEl = document.getElementById('status');
  statusEl.textContent = 'Scraping...';

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  if (!tab.url || !tab.url.includes('aliexpress.com')) {
    statusEl.textContent = 'Please open the AliExpress Orders page first.';
    return;
  }

  const [{ result: allRows }] = await chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: scrapeOrders,
  });

  if (!allRows || allRows.length === 0) {
    statusEl.textContent = 'No orders found on this page.';
    return;
  }

  const dateFrom = document.getElementById('dateFrom').value;
  const dateTo = document.getElementById('dateTo').value;
  const statusFilter = document.getElementById('statusFilter').value;
  const format = document.querySelector('input[name="format"]:checked').value;

  const rows = filterRows(allRows, dateFrom, dateTo, statusFilter);

  if (rows.length === 0) {
    statusEl.textContent = 'No orders match the selected filters.';
    return;
  }

  const date = new Date();
  const stamp = date.toISOString().slice(0, 10);

  let blob, filename;
  if (format === 'csv') {
    const csv = '﻿' + toCSV(rows);
    blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    filename = `aliexpress_orders_${stamp}.csv`;
  } else {
    const xlsx = toXLSX(rows);
    blob = new Blob([xlsx], { type: 'application/vnd.ms-excel' });
    filename = `aliexpress_orders_${stamp}.xls`;
  }

  const url = URL.createObjectURL(blob);

  await chrome.downloads.download({
    url,
    filename,
    saveAs: true,
  });

  statusEl.textContent = `Exported ${rows.length} row(s).`;
});
