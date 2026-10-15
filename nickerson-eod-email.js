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
    <img src="data:image/png;base64,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" alt="Nickerson" style="height:48px;width:auto;display:block;margin-bottom:12px;">
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
