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
    <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAsAAAADSCAYAAABErnZWAAAD7klEQVR42u3dyw2DQBBEwcX5x8u1HYA52Hzkha6KwILL02g8LEkyAACgxMsjAABAAAMAgAAGAAABDAAAAhgAAAQwAAAIYAAAEMAAACCAAQBAAAMAgAAGAAABDAAAAhgAAAEMAAACGAAABDAAAAhgAAAQwAAAIIABAEAAAwCAAAYAAAEMAAACGAAABDAAAAhgAAAEMAAACGAAABDAAAAggAEAQAADAIAABgAAAQwAAAIYAAAEMAAACGAAABDAAAAggAEAEMAAACCAAQBAAAMAgAAGAAABDAAAAhgAAAQwAAAIYAAAEMAAACCAAQBAAAMAgAAGAAABDACAAAYAAAEMAAACGAAABDAAAE+33vnHL0niHQIA0MIEGAAAAQwAAAIYAAAEMAAACOCrrF4XAABHuQIBAEAVKxAAAAhgAAAQwAAAIIABAEAAAwDAtBe8XIEAAKCKCTAAAAIYAAAEMAAACGAAABDAAAAggAEAQAADAIAABgAAAQwAAAIYAAAEMAAACGAAABDAAAAIYAAAEMAAACCAAQBAAAMAgAAGAAABDAAAAhgAONXqEcCnJUk8BgAAWpgAAwAggAEAQAADAIAABgAAAQyczz+5AeAkrkAAAFDFBBgAAAEMAAACGIBf2NsGmJQdYAAAqpgAAwAggAGYnhULgJ2sQAAAUMUEGAAAAQwAAAIYAAAEMAAACGAAABDAAAAggAEAQAAD8AUfwgDYyYcwAACoYgIMAIAABuAwKwoAk7ICAQBAFRNgAAAEMAAACGDgn+yTAsBJ7AADAFDFBBgAAAEMAAACGAAABDAAAAhgAGAOLsjABlcgAACoYgIMAIAABgAAAQwAAAIYAAAEMAAACGAAABDAAAAggAEAQAADAIAABgAAAQwAAAIYAAAEMAAAAhgAAATwE6xeNQAAY4yxJInHAABACysQAAAIYAAAEMAAACCAAQBg07RHCPwJDgCAKibAAAAIYAAAEMAAACCAAQBAAF/F54wBADjMFQgAAKpYgQAAQAADAIAABgAAAQwAAAIYAIBnu/V1LlcgAACoYgIMAIAABgAAAQwAAAIYAAAEMAAACGAAABDAAAAggAEAQAADAIAABgAAAQwAAAIYAAAEMAAAAhgAAAQwAAAIYAAAEMAAACCAAQBAAAMAgAAGAAABDAAAAhgAAAQwAAAIYAAAEMAAAAhgAAAQwAAAIIABAEAAAwCAAAYAAAEMAAACGAAABDAAAAhgAAAQwAAAIIABAEAAAwAggAEAQAADAIAABgAAAQwAAAIYAAAEMAAACGAAABDAAAAggAEAQAADAIAABgAAAQwAAAIYAAABDAAAAhgAAAQwAADczhvQ3k+/lWX3RwAAAABJRU5ErkJggg==" alt="Mouldshop" style="height:auto;width:auto;display:block;margin-bottom:12px;">
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
