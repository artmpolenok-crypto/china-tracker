import Anthropic from '@anthropic-ai/sdk';

const client = new Anthropic();

const PROMPT = `Parse this invoice/packing list from a Chinese supplier.
Return ONLY valid JSON, no markdown, no preamble:
{"items":[{"name":"...","qty":0,"unit_price_cny":0,"total_cny":0}],"total_cny":0,"invoice_number":null,"supplier":null}
All numbers must be numeric. If total_cny is missing, sum the items.`;

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
    return Response.json(parsed);
  } catch (e) {
    console.error('Parse error:', e);
    return Response.json({ error: e.message }, { status: 500 });
  }
}
