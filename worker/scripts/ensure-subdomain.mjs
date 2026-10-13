// Registers the account's *.workers.dev subdomain on first deploy (the dashboard step Cloudflare otherwise requires).
const { CLOUDFLARE_API_TOKEN: token, CLOUDFLARE_ACCOUNT_ID: account } = process.env;
const api = `https://api.cloudflare.com/client/v4/accounts/${account}/workers/subdomain`;
const headers = { authorization: `Bearer ${token}`, 'content-type': 'application/json' };

const current = await fetch(api, { headers }).then(r => r.json()).catch(() => null);
if (current && current.success && current.result && current.result.subdomain) {
  console.log(`workers.dev subdomain: ${current.result.subdomain}.workers.dev`);
  process.exit(0);
}

const suffix = Math.random().toString(36).slice(2, 6);
for (const name of ['luckswan', 'luckswan-resort', 'luckswan-sales', `luckswan-${suffix}`]) {
  const r = await fetch(api, { method: 'PUT', headers, body: JSON.stringify({ subdomain: name }) }).then(x => x.json()).catch(e => ({ success: false, errors: [{ message: String(e) }] }));
  if (r.success) { console.log(`Registered workers.dev subdomain: ${name}.workers.dev`); process.exit(0); }
  console.log(`"${name}" not available: ${(r.errors || []).map(e => e.message).join('; ')}`);
}
console.log('::error::Could not register a workers.dev subdomain. Open dash.cloudflare.com → Workers & Pages and set one up, then run this workflow again.');
process.exit(1);
