#!/usr/bin/env node
// Mouldshop End-of-Day Sales Summary Email
// Queries Business Central via client_credentials, sends HTML email via Microsoft Graph

// Uses Node.js built-in fetch (Node 18+); HTTPS_PROXY is picked up automatically via the agent proxy.

require('dotenv').config();

const TENANT = process.env.BC_TENANT;
const CLIENT = process.env.BC_CLIENT_ID;
const SECRET = process.env.BC_CLIENT_SECRET;

if (!TENANT || !CLIENT || !SECRET) {
  console.error('❌ Missing required environment variables: BC_TENANT, BC_CLIENT_ID, BC_CLIENT_SECRET');
  process.exit(1);
}

const BC_ROOT  = `https://api.businesscentral.dynamics.com/v2.0/${TENANT}/Production/ODataV4/Company('Live%20Mouldshop')/`;
const TO_EMAIL  = 'matt@mouldshop.com';
const TO_EMAILS = [
  'matt@mouldshop.com',
  'mark@mouldshop.com',
  'Liam@mouldshop.com',
  'mattb@mouldshop.com',
];
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
  });
  if (!res.ok) throw new Error(`Token error ${res.status}: ${await res.text()}`);
  return (await res.json()).access_token;
}

async function bcGet(token, entity, query) {
  const url = `${BC_ROOT}${entity}?${query}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' },
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
      toRecipients: TO_EMAILS.map(a => ({ emailAddress: { address: a } }))
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
  });

  if (sendRes.status === 202) {
    console.log('✅ Email sent successfully.');
  } else {
    const err = await sendRes.text();
    throw new Error(`Graph sendMail failed ${sendRes.status}: ${err}`);
  }
}

main().catch(e => { console.error('❌', e.message); process.exit(1); });
