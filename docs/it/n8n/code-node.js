// LUX → LINE: one message per event of a function sheet or a quotation. Turn an event off with false.
// owner, 13 Nov 2026: only the approval steps for now (รออนุมัติ · อนุมัติ · ตีกลับ), for function sheets and quotations
const SEND = {
  created: false, sent: false, pending: true, approved: true, rejected: true, ack: false,        // ใบฟังก์ชัน
  quote_pending: true, quote_approved: true, quote_rejected: true,                              // ใบเสนอราคา
  test: true                                                                                     // ปุ่ม ทดสอบแจ้งเตือน LINE ใน LUX
};
const GROUP_ID = 'Cdfa4f3ef64472a17e5176618a318a81a'; // the LINE group (from the "เก็บ Group ID" workflow)

const TH_M = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
const d = s => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ''); return m ? `${Number(m[3])} ${TH_M[Number(m[2]) - 1]} ${Number(m[1]) + 543}` : ''; };
const range = (a, b) => !a ? '' : (!b || a === b ? d(a) : `${d(a)} – ${d(b)}`);
const out = [];
for (const item of $input.all()) {
  const e = item.json.body || item.json;
  if (!e || !SEND[e.event]) continue;
  const b = e.beo || {}, q = e.quote || {};
  if (e.event === 'test') {
    out.push({ json: { to: GROUP_ID, messages: [{ type: 'text', text: ['🔔 ทดสอบการแจ้งเตือนจาก LUX', e.by ? 'โดย ' + e.by : '', 'ถ้าเห็นข้อความนี้ แปลว่า LUX → n8n → LINE ทำงานครบ'].filter(Boolean).join('\n') }] } });
    continue;
  }
  if (e.kind === 'quote') {
    const qn = `ใบเสนอราคา ${q.id}`;
    const qh = { quote_pending: `⏳ ${qn} ส่งขออนุมัติ`, quote_approved: `✅ ${qn} อนุมัติแล้ว`, quote_rejected: `↩️ ${qn} ถูกตีกลับ` }[e.event];
    const ql = [qh, [q.title, q.client].filter(Boolean).join(' · '), [range(q.start, q.end), q.pax ? q.pax + ' ท่าน' : ''].filter(Boolean).join(' · ')];
    if (q.owner) ql.push('พนักงานขาย: ' + q.owner);
    if (e.event === 'quote_rejected' && e.reason) ql.push('เหตุผล: ' + e.reason);
    if (e.by) ql.push('โดย ' + e.by);
    if (e.link) ql.push(e.link);
    out.push({ json: { to: GROUP_ID, messages: [{ type: 'text', text: ql.filter(Boolean).join('\n').slice(0, 4900) }] } });
    continue;
  }
  const name = `ใบฟังก์ชัน ${b.id}${b.no ? ' (' + b.no + ')' : ''}`;
  const head = {
    created: `🆕 สร้าง${name}`,
    sent: b.rev ? `📋 ${name} แก้ไขครั้งที่ ${b.rev} — ส่งให้แผนกแล้ว` : `📋 ${name} — ส่งให้แผนกแล้ว`,
    pending: `⏳ ${name} มีการแก้ไข รออนุมัติ`,
    approved: `✅ ${name} อนุมัติแล้ว`,
    rejected: `↩️ ${name} ถูกตีกลับ`,
    ack: `👍 ${(e.acked || []).join(', ')} รับทราบ${name}`
  }[e.event];
  const lines = [head, [q.title, q.client].filter(Boolean).join(' · '), [range(q.start, q.end), b.pax || (q.pax ? q.pax + ' ท่าน' : '')].filter(Boolean).join(' · ')];
  if (e.event === 'rejected' && e.reason) lines.push('เหตุผล: ' + e.reason);
  if ((e.event === 'sent' || e.event === 'ack') && (e.waiting || []).length) lines.push('รอรับทราบ: ' + e.waiting.join(', '));
  if (e.event === 'ack' && !(e.waiting || []).length) lines.push('ครบทุกแผนกแล้ว');
  if (e.by) lines.push('โดย ' + e.by);
  if (e.link) lines.push(e.link);
  out.push({ json: { to: GROUP_ID, messages: [{ type: 'text', text: lines.filter(Boolean).join('\n').slice(0, 4900) }] } });
}
return out;