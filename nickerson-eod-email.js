#!/usr/bin/env node
// Nickerson End-of-Day Sales Summary Email
// Queries Business Central via client_credentials, sends HTML email via Microsoft Graph

const { ProxyAgent, fetch } = require('undici');
const dispatcher = new ProxyAgent(process.env.HTTPS_PROXY || 'http://127.0.0.1:40037');

const TENANT   = 'f17044ae-2b66-4aa4-951b-38908b13d18b';
const CLIENT   = '94cf8276-e615-4d6e-a825-fb71f5568b8b';
const SECRET   = 'PCw8Q~JToVo-lSDikqEtwl4G7tXyDbBuYBsS5aab';
const BC_ROOT  = `https://api.businesscentral.dynamics.com/v2.0/${TENANT}/Production/ODataV4/Company('Nickerson%20(Live)')/`;
const TO_EMAIL = 'matt@mouldshop.com';
const DAILY_TARGET = 7840;

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
    'Open':                { bg: '#1e3260', fg: '#8fa3c8' },
    'Released':            { bg: '#052d1f', fg: '#34d399' },
    'Pending Prepayment':  { bg: '#271a02', fg: '#fbbf24' },
    'Pending Approval':    { bg: '#271a02', fg: '#fbbf24' },
  };
  const c = colours[status] || { bg: '#162444', fg: '#e8edf5' };
  return `<span style="background:${c.bg};color:${c.fg};padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700;">${status}</span>`;
}

async function main() {
  const today = new Date().toISOString().slice(0, 10);
  const dayName = new Date().toLocaleDateString('en-GB', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });

  console.log(`Fetching BC token...`);
  const bcToken = await getToken(`https://api.businesscentral.dynamics.com/.default`);

  console.log(`Querying Sales Orders for ${today}...`);
  const orders = await bcGet(bcToken, 'SalesOrder',
    `$select=No,Sell_to_Customer_Name,Status,Order_Date` +
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

  for (const o of orders) {
    const amt = amountByOrder[o.No] || 0;
    o.Amount = amt;
    grandTotal += amt;

    const st = o.Status || 'Unknown';
    byStatus[st] = byStatus[st] || { count: 0, total: 0 };
    byStatus[st].count++;
    byStatus[st].total += amt;
  }

  const pctOfTarget = (grandTotal / DAILY_TARGET) * 100;

  // --- Build HTML ---
  const orderRows = orders.map(o => `
    <tr>
      <td style="padding:8px 10px;border-bottom:1px solid #1e3260;font-family:monospace;font-size:13px;color:#8fa3c8;">${o.No}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #1e3260;color:#e8edf5;">${o.Sell_to_Customer_Name || '—'}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #1e3260;text-align:right;color:#e8edf5;font-variant-numeric:tabular-nums;">${fmt(o.Amount)}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #1e3260;">${statusBadge(o.Status)}</td>
    </tr>`).join('');

  const statusRows = Object.entries(byStatus)
    .sort((a,b) => b[1].total - a[1].total)
    .map(([st, d]) => `
    <tr>
      <td style="padding:7px 10px;border-bottom:1px solid #1e3260;width:40%;">${statusBadge(st)}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #1e3260;text-align:center;color:#8fa3c8;width:15%;font-variant-numeric:tabular-nums;">${d.count}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #1e3260;text-align:right;color:#e8edf5;width:25%;font-variant-numeric:tabular-nums;">${fmt(d.total)}</td>
      <td style="padding:7px 10px;border-bottom:1px solid #1e3260;text-align:right;color:#5a7099;width:20%;font-variant-numeric:tabular-nums;font-size:12px;">${((d.total/grandTotal)*100).toFixed(1)}%</td>
    </tr>`).join('');

  const html = `<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#0a1428;font-family:'Segoe UI',Arial,sans-serif;">
<div style="max-width:700px;margin:0 auto;padding:24px 16px;">

  <div style="background:#0f1e38;border-top:4px solid #f0c60a;border-radius:8px;padding:20px 24px;margin-bottom:20px;">
    <img src="data:image/png;base64,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" alt="Nickerson" style="height:auto;width:auto;display:block;margin-bottom:12px;">
    <div style="font-size:22px;font-weight:700;color:#e8edf5;">End of Day Sales Summary</div>
    <div style="font-size:13px;color:#8fa3c8;margin-top:4px;">${dayName}</div>
  </div>

  <div style="display:flex;gap:12px;margin-bottom:20px;">
    <div style="flex:1;background:#0f1e38;border:1px solid #1e3260;border-radius:8px;padding:16px 20px;">
      <div style="font-size:11px;color:#5a7099;text-transform:uppercase;letter-spacing:0.08em;">Total Orders</div>
      <div style="font-size:28px;font-weight:700;color:#e8edf5;margin-top:4px;">${orders.length}</div>
    </div>
    <div style="flex:2;background:#0f1e38;border:1px solid #f0c60a;border-radius:8px;padding:16px 20px;">
      <div style="font-size:11px;color:#5a7099;text-transform:uppercase;letter-spacing:0.08em;">Total Sales excl. VAT</div>
      <div style="font-size:28px;font-weight:700;color:#f0c60a;margin-top:4px;">${fmt(grandTotal)}</div>
      <div style="font-size:12px;color:${pctOfTarget >= 100 ? '#34d399' : pctOfTarget >= 75 ? '#fbbf24' : '#f87171'};margin-top:6px;font-weight:600;">${pctOfTarget.toFixed(1)}% of ${fmt(DAILY_TARGET)} target</div>
    </div>
  </div>

  <div style="background:#0f1e38;border:1px solid #1e3260;border-radius:8px;margin-bottom:20px;overflow:hidden;">
    <div style="padding:14px 16px;border-bottom:1px solid #1e3260;font-size:13px;font-weight:600;color:#8fa3c8;text-transform:uppercase;letter-spacing:0.06em;">Sales Orders</div>
    <table style="width:100%;border-collapse:collapse;">
      <thead>
        <tr style="background:#162444;">
          <th style="padding:8px 10px;text-align:left;font-size:11px;color:#5a7099;text-transform:uppercase;letter-spacing:0.06em;">Order No.</th>
          <th style="padding:8px 10px;text-align:left;font-size:11px;color:#5a7099;text-transform:uppercase;letter-spacing:0.06em;">Customer</th>
          <th style="padding:8px 10px;text-align:right;font-size:11px;color:#5a7099;text-transform:uppercase;letter-spacing:0.06em;">Amount excl. VAT</th>
          <th style="padding:8px 10px;text-align:left;font-size:11px;color:#5a7099;text-transform:uppercase;letter-spacing:0.06em;">Status</th>
        </tr>
      </thead>
      <tbody>${orderRows}</tbody>
      <tfoot>
        <tr style="background:#162444;">
          <td colspan="2" style="padding:10px 10px;font-weight:700;color:#e8edf5;">Total (${orders.length} orders)</td>
          <td style="padding:10px 10px;text-align:right;font-weight:700;color:#f0c60a;font-variant-numeric:tabular-nums;">${fmt(grandTotal)}</td>
          <td></td>
        </tr>
      </tfoot>
    </table>
  </div>

  <div style="background:#0f1e38;border:1px solid #1e3260;border-radius:8px;margin-bottom:20px;overflow:hidden;">
    <div style="padding:14px 16px;border-bottom:1px solid #1e3260;font-size:13px;font-weight:600;color:#8fa3c8;text-transform:uppercase;letter-spacing:0.06em;">By Status</div>
    <table style="width:100%;border-collapse:collapse;table-layout:fixed;">
      <thead>
        <tr style="background:#162444;">
          <th style="padding:8px 10px;text-align:left;font-size:11px;color:#5a7099;text-transform:uppercase;letter-spacing:0.06em;width:40%;">Status</th>
          <th style="padding:8px 10px;text-align:center;font-size:11px;color:#5a7099;text-transform:uppercase;letter-spacing:0.06em;width:15%;">Orders</th>
          <th style="padding:8px 10px;text-align:right;font-size:11px;color:#5a7099;text-transform:uppercase;letter-spacing:0.06em;width:25%;">Total excl. VAT</th>
          <th style="padding:8px 10px;text-align:right;font-size:11px;color:#5a7099;text-transform:uppercase;letter-spacing:0.06em;width:20%;">% of Total</th>
        </tr>
      </thead>
      <tbody>${statusRows}</tbody>
    </table>
  </div>

  <div style="font-size:11px;color:#5a7099;text-align:center;padding-top:8px;">
    Generated automatically from Business Central · Nickerson Live
  </div>
</div>
</body></html>`;

  // --- Send via Microsoft Graph ---
  console.log('Getting Graph token...');
  const graphToken = await getToken('https://graph.microsoft.com/.default');

  const subject = `Nickerson Sales Summary — ${dayName} — ${fmt(grandTotal)} (${orders.length} orders)`;

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
