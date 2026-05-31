import Anthropic from '@anthropic-ai/sdk';

const client = new Anthropic();

const PROMPT = `Parse this invoice/packing list from a Chinese supplier.
Translate ALL item names to Russian.
Separate GOODS (products going to warehouse) from EXPENSES (delivery, packaging, pallets, insurance etc).

Return ONLY valid JSON, no markdown:
{
  "items": [{"name":"название товара на русском","qty":0,"unit_price_cny":0,"total_cny":0,"is_expense":false}],
  "expenses": [{"name":"название расхода на русском","total_cny":0}],
  "total_cny": 0,
  "invoice_number": null,
  "supplier": null
}

Rules:
- is_expense=true for: delivery/shipping (运费), packaging (包装), pallets (木托), insurance, any service fees
- is_expense=false for actual physical goods
- qty = actual quantity number
- All numbers must be numeric
- If total_cny missing, sum all items + expenses
- supplier = company name as-is`;

function getMimeType(file) {
  if (file.type && file.type !== 'application/octet-stream') return file.type;
  const name = file.name || '';
  if (name.match(/\.(jpg|jpeg)$/i)) return 'image/jpeg';
  if (name.match(/\.png$/i)) return 'image/png';
  if (name.match(/\.gif$/i)) return 'image/gif';
  if (name.match(/\.webp$/i)) return 'image/webp';
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
      messages = [{ role: 'user', content: `${PROMPT}\n\nFile contents:\n${text.slice(0, 8000)}` }];
    } else {
      const bytes = await file.arrayBuffer();
      const base64 = Buffer.from(bytes).toString('base64');
      const mime = getMimeType(file);

      let block;
      if (mime.startsWith('image/')) {
        block = { type: 'image', source: { type: 'base64', media_type: mime, data: base64 } };
      } else {
        block = { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: base64 } };
      }

      messages = [{ role: 'user', content: [block, { type: 'text', text: PROMPT }] }];
    }

    const response = await client.messages.create({
      model: 'claude-sonnet-4-5',
      max_tokens: 1000,
      messages,
    });

    const raw = response.content.map(b => b.text || '').join('');
    const parsed = JSON.parse(raw.replace(/```json|```/g, '').trim());
    
    // Move expense items to expenses array if not already separated
    if (parsed.items) {
      const expenseKeywords = ['доставка', 'упаковка', 'поддон', 'транспорт', 'фрахт', 'тара', 'паллет', 'страховка', '运费', '木托', '包装'];
      const isExp = (name) => expenseKeywords.some(k => name?.toLowerCase().includes(k.toLowerCase()));
      
      if (!parsed.expenses) parsed.expenses = [];
      const expFromItems = parsed.items.filter(i => i.is_expense || isExp(i.name));
      const goodsOnly = parsed.items.filter(i => !i.is_expense && !isExp(i.name));
      
      expFromItems.forEach(e => parsed.expenses.push({ name: e.name, total_cny: e.total_cny || 0 }));
      parsed.items = goodsOnly;
    }
    
    return Response.json(parsed);
  } catch (e) {
    console.error('Parse error:', e);
