import Anthropic from '@anthropic-ai/sdk';

const client = new Anthropic();

const PROMPT = `Parse this invoice/packing list from a Chinese supplier.
Return ONLY valid JSON, no markdown, no preamble:
{"items":[{"name":"...","qty":0,"unit_price_cny":0,"total_cny":0}],"total_cny":0,"invoice_number":null,"supplier":null}
All numbers must be numeric. If total_cny is missing, sum the items.`;

export async function POST(request) {
  try {
    const formData = await request.formData();
    const file = formData.get('file');
    const text = formData.get('text'); // for Excel text extracted client-side

    let messages;

    if (text) {
      messages = [{ role: 'user', content: `${PROMPT}\n\nFile contents:\n${text.slice(0, 8000)}` }];
    } else {
      const bytes = await file.arrayBuffer();
      const base64 = Buffer.from(bytes).toString('base64');
      const mime = file.type || 'application/octet-stream';

      const block = mime.startsWith('image/')
        ? { type: 'image', source: { type: 'base64', media_type: mime, data: base64 } }
        : { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: base64 } };

      messages = [{ role: 'user', content: [block, { type: 'text', text: PROMPT }] }];
    }

    const response = await client.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 1000,
      messages,
    });

    const raw = response.content.map(b => b.text || '').join('');
    const parsed = JSON.parse(raw.replace(/```json|```/g, '').trim());
    return Response.json(parsed);
  } catch (e) {
    return Response.json({ error: e.message }, { status: 500 });
  }
}
