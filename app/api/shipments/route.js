import { kv } from '@vercel/kv';

const KEY = 'shipments';

async function getAll() {
  return (await kv.get(KEY)) || [];
}

export async function GET() {
  const shipments = await getAll();
  return Response.json(shipments);
}

export async function POST(request) {
  const body = await request.json();
  const shipments = await getAll();
  const updated = [...shipments, body];
  await kv.set(KEY, updated);
  return Response.json(body);
}

export async function PUT(request) {
  const body = await request.json();
  const shipments = await getAll();
  const updated = shipments.map(s => s.id === body.id ? body : s);
  await kv.set(KEY, updated);
  return Response.json(body);
}

export async function DELETE(request) {
  const { id } = await request.json();
  const shipments = await getAll();
  const updated = shipments.filter(s => s.id !== id);
  await kv.set(KEY, updated);
  return Response.json({ ok: true });
}
