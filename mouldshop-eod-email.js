#!/usr/bin/env node
// Mouldshop End-of-Day Sales Summary Email
// Queries Business Central via client_credentials, sends HTML email via Microsoft Graph

const { ProxyAgent, fetch } = require('undici');
const dispatcher = new ProxyAgent(process.env.HTTPS_PROXY || 'http://127.0.0.1:40037');

const TENANT   = 'f17044ae-2b66-4aa4-951b-38908b13d18b';
const CLIENT   = '94cf8276-e615-4d6e-a825-fb71f5568b8b';
const SECRET   = 'PCw8Q~JToVo-lSDikqEtwl4G7tXyDbBuYBsS5aab';
const BC_ROOT  = `https://api.businesscentral.dynamics.com/v2.0/${TENANT}/Production/ODataV4/Company('Live%20Mouldshop')/`;
const TO_EMAIL = 'matt@mouldshop.com';
const DAILY_TARGET = 16678;

async function getToken(scope) {
  const res = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type:    'client_credentials',
      client_id:     CLIENT,
      client_secret: SECRET,
      scope
    }).toString(),
    dispatcher
  });
  if (!res.ok) throw new Error(`Token error ${res.status}: ${await res.text()}`);
  return (await res.json()).access_token;
}

async function bcGet(token, entity, query) {
  const url = `${BC_ROOT}${entity}?${query}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
    dispatcher
  });
  if (!res.ok) throw new Error(`BC error ${res.status} on ${entity}: ${await res.text()}`);
  return (await res.json()).value;
}

async function bcGetAll(token, entity, query) {
  // Follows @odata.nextLink to retrieve all pages
  let url = `${BC_ROOT}${entity}?${query}`;
  const results = [];
  while (url) {
    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
      dispatcher
    });
    if (!res.ok) throw new Error(`BC error ${res.status} on ${entity}: ${await res.text()}`);
    const json = await res.json();
    results.push(...(json.value || []));
    url = json['@odata.nextLink'] || null;
  }
  return results;
}

function fmt(n) {
  return '£' + Number(n).toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function statusBadge(status) {
  const colours = {
    'Open':                { bg: '#253470', fg: '#9aafd4' },
    'Released':            { bg: '#052d14', fg: '#4ade80' },
    'Pending Prepayment':  { bg: '#271a02', fg: '#fbbf24' },
    'Pending Approval':    { bg: '#1e1a02', fg: '#fbbf24' },
  };
  const c = colours[status] || { bg: '#1e2a58', fg: '#f4f4f4' };
  return `<span style="background:${c.bg};color:${c.fg};padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600;">${status}</span>`;
}

async function main() {
  const today = new Date().toISOString().slice(0, 10);
  const dayName = new Date().toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });

  console.log(`Fetching BC token...`);
  const bcToken = await getToken(`https://api.businesscentral.dynamics.com/.default`);

  console.log(`Querying Sales Orders for ${today}...`);
  const orders = await bcGet(bcToken, 'SalesOrder',
    `$select=No,Sell_to_Customer_Name,Salesperson_Code,Status,Order_Date` +
    `&$filter=Order_Date eq ${today}` +
    `&$orderby=No asc&$top=500`
  );

  if (orders.length === 0) {
    console.log('No orders today — skipping email.');
    return;
  }

  // --- Fetch line amounts for all today's orders in parallel ---
  console.log(`Fetching line amounts for ${orders.length} orders...`);
  const amountByOrder = {};
  await Promise.all(orders.map(async o => {
    const lines = await bcGetAll(bcToken, 'SalesOrderSalesLines',
      `$select=Document_No,Line_Amount&$filter=Document_No eq '${o.No}'`
    );
    amountByOrder[o.No] = lines.reduce((sum, l) => sum + (parseFloat(l.Line_Amount) || 0), 0);
  }));

  // --- Aggregate ---
  let grandTotal = 0;
  const byStatus = {};
  const bySalesperson = {};

  for (const o of orders) {
    const amt = amountByOrder[o.No] || 0;
    o.Amount = amt;
    grandTotal += amt;

    const st = o.Status || 'Unknown';
    byStatus[st] = byStatus[st] || { count: 0, total: 0 };
    byStatus[st].count++;
    byStatus[st].total += amt;

    const sp = o.Salesperson_Code || '—';
    bySalesperson[sp] = bySalesperson[sp] || { count: 0, total: 0 };
    bySalesperson[sp].count++;
    bySalesperson[sp].total += amt;
  }

  const pctOfTarget = (grandTotal / DAILY_TARGET) * 100;

  // --- Build HTML ---
  const orderRows = orders.map(o => `
    <tr>
      <td style="padding:8px 10px;border-bottom:1px solid #253470;font-family:monospace;font-size:13px;color:#9aafd4;">${o.No}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #253470;color:#f4f4f4;">${o.Sell_to_Customer_Name || '—'}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #253470;text-align:right;color:#f4f4f4;font-variant-numeric:tabular-nums;">${fmt(o.Amount)}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #253470;text-align:center;color:#9aafd4;font-weight:600;">${o.Salesperson_Code || '—'}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #253470;">${statusBadge(o.Status)}</td>
    </tr>`).join('');

  const statusRows = Object.entries(byStatus)
    .sort((a,b) => b[1].total - a[1].total)
    .map(([st, d]) => `
    <tr>
      <td style="padding:7px 10px;border-bottom:1px solid #253470;width:40%;">${statusBadge(st)}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #253470;text-align:center;color:#9aafd4;width:15%;font-variant-numeric:tabular-nums;">${d.count}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #253470;text-align:right;color:#f4f4f4;width:25%;font-variant-numeric:tabular-nums;">${fmt(d.total)}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #253470;text-align:right;color:#607098;width:20%;font-variant-numeric:tabular-nums;font-size:12px;">${((d.total/grandTotal)*100).toFixed(1)}%</td>
    </tr>`).join('');

  const spRows = Object.entries(bySalesperson)
    .sort((a,b) => b[1].total - a[1].total)
    .map(([sp, d]) => `
    <tr>
      <td style="padding:7px 10px;border-bottom:1px solid #253470;color:#f4f4f4;font-weight:600;width:40%;">${sp}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #253470;text-align:center;color:#9aafd4;width:15%;font-variant-numeric:tabular-nums;">${d.count}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #253470;text-align:right;color:#f4f4f4;width:25%;font-variant-numeric:tabular-nums;">${fmt(d.total)}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #253470;text-align:right;color:#607098;width:20%;font-variant-numeric:tabular-nums;font-size:12px;">${((d.total/grandTotal)*100).toFixed(1)}%</td>
    </tr>`).join('');

  const html = `<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#0e1733;font-family:'Segoe UI',Arial,sans-serif;">
<div style="max-width:700px;margin:0 auto;padding:24px 16px;">

  <div style="background:#111b3e;border-top:4px solid #5b8db8;border-radius:8px;padding:20px 24px;margin-bottom:20px;">
    <img src="data:image/png;base64,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" alt="Mouldshop" style="height:auto;width:auto;display:block;margin-bottom:12px;">
    <div style="font-size:22px;font-weight:700;color:#f4f4f4;">End of Day Sales Summary</div>
    <div style="font-size:13px;color:#9aafd4;margin-top:4px;">${dayName}</div>
  </div>

  <div style="display:flex;gap:12px;margin-bottom:20px;">
    <div style="flex:1;background:#182048;border:1px solid #253470;border-radius:8px;padding:16px 20px;">
      <div style="font-size:11px;color:#607098;text-transform:uppercase;letter-spacing:0.08em;">Total Orders</div>
      <div style="font-size:28px;font-weight:700;color:#f4f4f4;margin-top:4px;">${orders.length}</div>
    </div>
    <div style="flex:2;background:#182048;border:1px solid #5b8db8;border-radius:8px;padding:16px 20px;">
      <div style="font-size:11px;color:#607098;text-transform:uppercase;letter-spacing:0.08em;">Total Sales excl. VAT</div>
      <div style="font-size:28px;font-weight:700;color:#5b8db8;margin-top:4px;">${fmt(grandTotal)}</div>
      <div style="font-size:12px;color:${pctOfTarget >= 100 ? '#4ade80' : pctOfTarget >= 75 ? '#fbbf24' : '#f87171'};margin-top:6px;font-weight:600;">${pctOfTarget.toFixed(1)}% of ${fmt(DAILY_TARGET)} target</div>
    </div>
  </div>

  <div style="background:#182048;border:1px solid #253470;border-radius:8px;margin-bottom:20px;overflow:hidden;">
    <div style="padding:14px 16px;border-bottom:1px solid #253470;font-size:13px;font-weight:600;color:#9aafd4;text-transform:uppercase;letter-spacing:0.06em;">Sales Orders</div>
    <table style="width:100%;border-collapse:collapse;">
      <thead>
        <tr style="background:#1e2a58;">
          <th style="padding:8px 10px;text-align:left;font-size:11px;color:#607098;text-transform:uppercase;letter-spacing:0.06em;">Order No.</th>
          <th style="padding:8px 10px;text-align:left;font-size:11px;color:#607098;text-transform:uppercase;letter-spacing:0.06em;">Customer</th>
          <th style="padding:8px 10px;text-align:right;font-size:11px;color:#607098;text-transform:uppercase;letter-spacing:0.06em;">Amount excl. VAT</th>
          <th style="padding:8px 10px;text-align:center;font-size:11px;color:#607098;text-transform:uppercase;letter-spacing:0.06em;">Salesperson</th>
          <th style="padding:8px 10px;text-align:left;font-size:11px;color:#607098;text-transform:uppercase;letter-spacing:0.06em;">Status</th>
        </tr>
      </thead>
      <tbody>${orderRows}</tbody>
      <tfoot>
        <tr style="background:#1e2a58;">
          <td colspan="2" style="padding:10px 10px;font-weight:700;color:#f4f4f4;">Total (${orders.length} orders)</td>
          <td style="padding:10px 10px;text-align:right;font-weight:700;color:#5b8db8;font-variant-numeric:tabular-nums;">${fmt(grandTotal)}</td>
          <td colspan="2"></td>
        </tr>
      </tfoot>
    </table>
  </div>

  <div style="background:#182048;border:1px solid #253470;border-radius:8px;margin-bottom:20px;overflow:hidden;">
    <div style="padding:14px 16px;border-bottom:1px solid #253470;font-size:13px;font-weight:600;color:#9aafd4;text-transform:uppercase;letter-spacing:0.06em;">By Status</div>
    <table style="width:100%;border-collapse:collapse;table-layout:fixed;">
      <thead>
        <tr style="background:#1e2a58;">
          <th style="padding:8px 10px;text-align:left;font-size:11px;color:#607098;text-transform:uppercase;letter-spacing:0.06em;width:40%;">Status</th>
          <th style="padding:8px 10px;text-align:center;font-size:11px;color:#607098;text-transform:uppercase;letter-spacing:0.06em;width:15%;">Orders</th>
          <th style="padding:8px 10px;text-align:right;font-size:11px;color:#607098;text-transform:uppercase;letter-spacing:0.06em;width:25%;">Total excl. VAT</th>
          <th style="padding:8px 10px;text-align:right;font-size:11px;color:#607098;text-transform:uppercase;letter-spacing:0.06em;width:20%;">% of Total</th>
        </tr>
      </thead>
      <tbody>${statusRows}</tbody>
    </table>
  </div>

  <div style="background:#182048;border:1px solid #253470;border-radius:8px;margin-bottom:20px;overflow:hidden;">
    <div style="padding:14px 16px;border-bottom:1px solid #253470;font-size:13px;font-weight:600;color:#9aafd4;text-transform:uppercase;letter-spacing:0.06em;">By Salesperson</div>
    <table style="width:100%;border-collapse:collapse;table-layout:fixed;">
      <thead>
        <tr style="background:#1e2a58;">
          <th style="padding:8px 10px;text-align:left;font-size:11px;color:#607098;text-transform:uppercase;letter-spacing:0.06em;width:40%;">Salesperson</th>
          <th style="padding:8px 10px;text-align:center;font-size:11px;color:#607098;text-transform:uppercase;letter-spacing:0.06em;width:15%;">Orders</th>
          <th style="padding:8px 10px;text-align:right;font-size:11px;color:#607098;text-transform:uppercase;letter-spacing:0.06em;width:25%;">Total excl. VAT</th>
          <th style="padding:8px 10px;text-align:right;font-size:11px;color:#607098;text-transform:uppercase;letter-spacing:0.06em;width:20%;">% of Total</th>
        </tr>
      </thead>
      <tbody>${spRows}</tbody>
    </table>
  </div>

  <div style="font-size:11px;color:#607098;text-align:center;padding-top:8px;">
    Generated automatically from Business Central · Mouldshop Live
  </div>
</div>
</body></html>`;

  // --- Send via Microsoft Graph ---
  console.log('Getting Graph token...');
  const graphToken = await getToken('https://graph.microsoft.com/.default');

  const subject = `Mouldshop Sales Summary — ${dayName} — ${fmt(grandTotal)} (${orders.length} orders)`;

  const payload = {
    message: {
      subject,
      body: { contentType: 'HTML', content: html },
      toRecipients: [{ emailAddress: { address: TO_EMAIL } }]
    },
    saveToSentItems: true
  };

  console.log('Sending email via Graph...');
  const sendRes = await fetch(`https://graph.microsoft.com/v1.0/users/${TO_EMAIL}/sendMail`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${graphToken}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(payload),
    dispatcher
  });

  if (sendRes.status === 202) {
    console.log('✅ Email sent successfully.');
  } else {
    const err = await sendRes.text();
    throw new Error(`Graph sendMail failed ${sendRes.status}: ${err}`);
  }
}

main().catch(e => { console.error('❌', e.message); process.exit(1); });
