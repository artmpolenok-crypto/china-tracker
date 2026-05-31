import { Redis } from '@upstash/redis';

const redis = new Redis({
  url: process.env.UPSTASH_REDIS_REST_URL,
  token: process.env.UPSTASH_REDIS_REST_TOKEN,
});

const KEY = 'warehouse';

async function getAll() {
  return (await redis.get(KEY)) || [];
}

export async function GET() {
  const items = await getAll();
  return Response.json(items);
}

export async function POST(request) {
  const body = await request.json();
  const items = await getAll();
  await redis.set(KEY, [...items, body]);
  return Response.json(body);
}

export async function PUT(request) {
  const body = await request.json();
  const items = await getAll();
  await redis.set(KEY, items.map(i => i.id === body.id ? body : i));
  return Response.json(body);
}

export async function DELETE(request) {
  const { id } = await request.json();
  const items = await getAll();
  await redis.set(KEY, items.filter(i => i.id !== id));
  return Response.json({ ok: true });
}
