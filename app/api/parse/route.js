import Anthropic from '@anthropic-ai/sdk';

const client = new Anthropic();

const PROMPT = `Parse this invoice/packing list from a Chinese supplier. Translate ALL names to Russian. Separate goods from expenses.
Return ONLY valid JSON:
{"items":[{"name":"товар на русском","qty":0,"unit_price_cny":0,"total_cny":0}],"expenses":[{"name":"расход на русском","total_cny":0}],"total_cny":0,"invoice_number":null,"supplier":null}
Rules:
- items = physical goods only (чашки, тарелки etc)
- expenses = delivery, packaging, pallets, insurance, any service fees
- qty must be numeric
- All prices numeric
- supplier = company name as-is`;

function getMimeType(file) {
  if (file.type && file.type !== 'application/octet-stream') return file.type;
  const name = file.name || '';
  if (name.match(/\.(jpg|jpeg)$/i)) return 'image/jpeg';
  if (name.match(/\.png$/i)) return 'image/png';
  if (name.match(/\.pdf$/i)) return 'application/pdf';
  return 'image/jpeg';
}

export async function POST(request) {
  try {
    const formData = await request.formData();
    const file = formData.get('file');
    const text = formData.get('text');
    let messages;
    if (text) {
      messages = [{ role: 'user', content: `${PROMPT}\n\nFile:\n${text.slice(0, 8000)}` }];
    } else {
      const bytes = await file.arrayBuffer();
      const base64 = Buffer.from(bytes).toString('base64');
      const mime = getMimeType(file);
      const block = mime.startsWith('image/')
        ? { type: 'image', source: { type: 'base64', media_type: mime, data: base64 } }
        : { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: base64 } };
      messages = [{ role: 'user', content: [block, { type: 'text', text: PROMPT }] }];
    }
    const response = await client.messages.create({ model: 'claude-sonnet-4-5', max_tokens: 1000, messages });
    const raw = response.content.map(b => b.text || '').join('');
    const parsed = JSON.parse(raw.replace(/```json|```/g, '').trim());
    if (!parsed.expenses) parsed.expenses = [];
    const expenseKeywords = ['доставка', 'упаковка', 'поддон', 'транспорт', 'фрахт', 'тара', 'паллет', 'страховка'];
    const isExp = (name) => expenseKeywords.some(k => name?.toLowerCase().includes(k.toLowerCase()));
    if (parsed.items) {
      const exp = parsed.items.filter(i => isExp(i.name));
      exp.forEach(e => parsed.expenses.push({ name: e.name, total_cny: e.total_cny || 0 }));
      parsed.items = parsed.items.filter(i => !isExp(i.name));
    }
    return Response.json(parsed);
  } catch (e) {
    console.error('Parse error:', e);
    return Response.json({ error: e.message }, { status: 500 });
  }
}
