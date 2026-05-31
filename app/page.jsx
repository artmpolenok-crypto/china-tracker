'use client';
import { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';

function genId() { return Date.now().toString(36) + Math.random().toString(36).slice(2); }
function fmt(n) { if (!n && n !== 0) return '—'; return Math.round(n).toLocaleString('ru-RU'); }
function fmtF(n, d = 2) { if (!n && n !== 0) return '—'; return (+n).toFixed(d); }

async function apiFetch(url, options = {}) {
  const r = await fetch(url, { headers: { 'Content-Type': 'application/json' }, ...options });
  return r.json();
}

async function parseInvoiceFile(file) {
  const mime = file.type;
  const isExcel = mime.includes('sheet') || mime.includes('excel') || file.name.match(/\.(xlsx|xls|csv)$/i);
  const fd = new FormData();
  if (isExcel) {
    const ab = await file.arrayBuffer();
    const wb = XLSX.read(ab, { type: 'array' });
    let text = '';
    wb.SheetNames.forEach(n => { text += `Sheet: ${n}\n` + XLSX.utils.sheet_to_csv(wb.Sheets[n]) + '\n\n'; });
    fd.append('text', text);
  } else {
    fd.append('file', file);
  }
  const r = await fetch('/api/parse', { method: 'POST', body: fd });
  const data = await r.json();
  if (data.error) throw new Error(data.error);
  return data;
}

async function compressPhoto(file) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      const max = 500;
      let w = img.width, h = img.height;
      if (w > max || h > max) {
        if (w > h) { h = Math.round(h * max / w); w = max; }
        else { w = Math.round(w * max / h); h = max; }
      }
      canvas.width = w; canvas.height = h;
      canvas.getContext('2d').drawImage(img, 0, 0, w, h);
      resolve(canvas.toDataURL('image/jpeg', 0.75));
    };
    img.src = URL.createObjectURL(file);
  });
}

const BADGES = {
  new:     { label: 'Новый',  cls: 'badge-new' },
  transit: { label: 'В пути', cls: 'badge-transit' },
  arrived: { label: 'Прибыл', cls: 'badge-arrived' },
  sold:    { label: 'Продан', cls: 'badge-sold' },
};

function Badge({ status }) {
  const b = BADGES[status] || BADGES.new;
  return <span className={`badge ${b.cls}`}>{b.label}</span>;
}

function Field({ label, children }) {
  return <div className="field"><label>{label}</label>{children}</div>;
}

function Row({ label, value }) {
  return (
    <tr>
      <td className="muted" style={{ width: '50%' }}>{label}</td>
      <td style={{ textAlign: 'right' }}>{value}</td>
    </tr>
  );
}

function MetricCard({ label, value, cls, sub }) {
  return (
    <div className="metric-card">
      <div className="metric-label">{label}</div>
      <div className={`metric-value ${cls || ''}`}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: 'rgba(255,255,255,0.6)', marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

function DropZone({ onFile, loading, label }) {
  const [drag, setDrag] = useState(false);
  const ref = useRef();
  function handleDrop(e) {
    e.preventDefault(); e.stopPropagation(); setDrag(false);
    const file = e.dataTransfer.files[0];
    if (file) onFile(file);
  }
  function handleDragOver(e) { e.preventDefault(); e.stopPropagation(); setDrag(true); }
  return (
    <div onDragOver={handleDragOver} onDragEnter={handleDragOver} onDragLeave={() => setDrag(false)}
      onDrop={handleDrop} onClick={() => !loading && ref.current.click()}
      style={{ border: `1.5px dashed ${drag ? 'rgba(255,255,255,0.6)' : 'rgba(255,255,255,0.3)'}`, borderRadius: 12, padding: '1.5rem 1rem', textAlign: 'center', cursor: loading ? 'default' : 'pointer', background: drag ? 'rgba(255,255,255,0.08)' : 'transparent' }}
    >
      <input ref={ref} type="file" accept=".pdf,.png,.jpg,.jpeg,.xlsx,.xls,.csv" style={{ display: 'none' }} onChange={e => e.target.files[0] && onFile(e.target.files[0])} />
      <div style={{ fontSize: 28, marginBottom: 6 }}>📄</div>
      {loading
        ? <div className="muted">Распознаём...</div>
        : <><div style={{ fontWeight: 500, marginBottom: 4 }}>{label || 'Загрузите накладную'}</div><div className="muted" style={{ fontSize: 12 }}>PDF · Фото · Excel</div></>
      }
    </div>
  );
}

// ─── SHIPMENTS ────────────────────────────────────────────────────────────────

function NewShipment({ onSave, onCancel }) {
  const [step, setStep] = useState(1);
  const [parsing, setParsing] = useState(false);
  const [parseError, setParseError] = useState('');
  const [fileName, setFileName] = useState('');
  const [saving, setSaving] = useState(false);
  const [form, setForm] = useState({ name: '', items: [], total_cny: '', invoice_number: '', supplier: '', cny_rate: '', paid_rub: '', ship_date: '', eta_date: '' });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  async function handleFile(file) {
    setFileName(file.name); setParsing(true); setParseError('');
    try {
      const d = await parseInvoiceFile(file);
      setForm(f => ({ ...f, items: d.items || [], total_cny: d.total_cny || '', invoice_number: d.invoice_number || '', supplier: d.supplier || '', name: d.supplier ? `Поставка — ${d.supplier}` : `Поставка ${new Date().toLocaleDateString('ru-RU')}` }));
      setStep(2);
    } catch { setParseError('Не удалось распознать файл. Введите вручную.'); }
    setParsing(false);
  }

  async function handleSave() {
    setSaving(true);
    const totalCNY = +form.total_cny || form.items.reduce((s, i) => s + (i.total_cny || 0), 0);
    const shipment = { id: genId(), createdAt: new Date().toISOString(), status: form.ship_date ? 'transit' : 'new', name: form.name || `Поставка ${new Date().toLocaleDateString('ru-RU')}`, items: form.items, total_cny: totalCNY, invoice_number: form.invoice_number, supplier: form.supplier, cny_rate: +form.cny_rate || 0, paid_rub: +form.paid_rub || 0, ship_date: form.ship_date, eta_date: form.eta_date, extra_paid_rub: 0, delivery_rub: 0, sale_price_rub: 0 };
    await apiFetch('/api/shipments', { method: 'POST', body: JSON.stringify(shipment) });
    onSave(shipment);
    setSaving(false);
  }

  const totalCNY = +form.total_cny || form.items.reduce((s, i) => s + (i.total_cny || 0), 0);
  const estRUB = totalCNY && form.cny_rate ? totalCNY * form.cny_rate : null;

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: '1.5rem' }}>
        <button onClick={onCancel} style={{ padding: '6px 10px' }}>← Назад</button>
        <div style={{ fontSize: 18, fontWeight: 500 }}>Новая поставка</div>
      </div>
      <div className="step-bar" style={{ marginBottom: '1.5rem' }}>
        {['Накладная','Оплата','Доставка','Итог'].map((s, i) => (
          <div key={i} style={{ flex: 1 }}>
            <div className="step-seg" style={{ background: i < step ? '#fff' : 'rgba(255,255,255,0.2)', marginBottom: 3 }} />
            {i === step - 1 && <div style={{ fontSize: 11, color: 'rgba(255,255,255,0.7)', textAlign: 'center' }}>{s}</div>}
          </div>
        ))}
      </div>

      {step === 1 && (
        <div>
          <DropZone onFile={handleFile} loading={parsing} />
          {parseError && <div className="error-box">{parseError}</div>}
          <div className="mt-1" style={{ textAlign: 'center' }}>
            <button style={{ border: 'none', background: 'none', color: 'rgba(255,255,255,0.6)', cursor: 'pointer', fontSize: 13 }} onClick={() => { set('name', `Поставка ${new Date().toLocaleDateString('ru-RU')}`); setStep(2); }}>Пропустить, ввести вручную →</button>
          </div>
        </div>
      )}

      {step === 2 && (
        <div>
          {fileName && <div className="info-box">📎 {fileName} — {form.items.length > 0 ? `${form.items.length} позиций` : 'вручную'}</div>}
          <Field label="Название поставки"><input value={form.name} onChange={e => set('name', e.target.value)} placeholder="Электроника январь" /></Field>
          <Field label="Итого по накладной, ¥"><input type="number" value={form.total_cny} onChange={e => set('total_cny', e.target.value)} placeholder="0" /></Field>
          <div className="grid-2">
            <Field label="Курс CNY/RUB"><input type="number" value={form.cny_rate} onChange={e => set('cny_rate', e.target.value)} placeholder="13.5" /></Field>
            <Field label="Оплачено, ₽"><input type="number" value={form.paid_rub} onChange={e => set('paid_rub', e.target.value)} placeholder="0" /></Field>
          </div>
          {estRUB && <div className="muted mb-1" style={{ fontSize: 13 }}>≈ {fmt(estRUB)} ₽</div>}
          <div className="row mt-1"><button onClick={() => setStep(1)}>← Назад</button><button className="primary flex-1" onClick={() => setStep(3)}>Далее →</button></div>
        </div>
      )}

      {step === 3 && (
        <div>
          <div className="grid-2">
            <Field label="Дата отправки"><input type="date" value={form.ship_date} onChange={e => set('ship_date', e.target.value)} /></Field>
            <Field label="Ожидаемое прибытие"><input type="date" value={form.eta_date} onChange={e => set('eta_date', e.target.value)} /></Field>
          </div>
          <div className="row mt-1"><button onClick={() => setStep(2)}>← Назад</button><button className="primary flex-1" onClick={() => setStep(4)}>Далее →</button></div>
        </div>
      )}

      {step === 4 && (
        <div>
          <div className="card mb-1">
            <div style={{ fontWeight: 500, fontSize: 15, marginBottom: 12 }}>{form.name}</div>
            <table className="items-table"><tbody>
              {totalCNY > 0 && <Row label="Закупка" value={`¥ ${fmt(totalCNY)}`} />}
              {form.cny_rate && <Row label="Курс" value={`${form.cny_rate} ₽/¥`} />}
              {form.paid_rub && <Row label="Оплачено" value={`${fmt(+form.paid_rub)} ₽`} />}
              {form.ship_date && <Row label="Отправлено" value={form.ship_date} />}
              {form.eta_date && <Row label="Ожидается" value={form.eta_date} />}
            </tbody></table>
            {form.items.length > 0 && (<><hr className="divider" /><div className="muted" style={{ fontSize: 12, marginBottom: 6 }}>Позиции ({form.items.length}):</div>{form.items.slice(0,4).map((item,i) => (<div key={i} style={{ fontSize: 12, display: 'flex', justifyContent: 'space-between', padding: '2px 0' }}><span>{item.name}</span><span className="muted">×{item.qty} · ¥{fmt(item.total_cny)}</span></div>))}{form.items.length > 4 && <div className="muted" style={{ fontSize: 12 }}>...и ещё {form.items.length - 4}</div>}</>)}
          </div>
          <div className="row"><button onClick={() => setStep(3)}>← Назад</button><button className="primary flex-1" onClick={handleSave} disabled={saving}>{saving ? 'Сохранение...' : 'Сохранить поставку'}</button></div>
        </div>
      )}
    </div>
  );
}

function ShipmentDetail({ shipment, onUpdate, onDelete, onBack }) {
  const [showArrival, setShowArrival] = useState(false);
  const [showSale, setShowSale] = useState(false);
  const [showItems, setShowItems] = useState(false);
  const [saving, setSaving] = useState(false);
  const [arrival, setArrival] = useState({ extra_paid_rub: '', delivery_rub: '', arrived_date: new Date().toISOString().slice(0, 10) });
  const [sale, setSale] = useState({ sale_price_rub: '' });

  const cost = (shipment.paid_rub || 0) + (shipment.extra_paid_rub || 0) + (shipment.delivery_rub || 0);
  const profit = shipment.status === 'sold' ? (shipment.sale_price_rub || 0) - cost : null;
  const margin = profit !== null && cost > 0 ? ((profit / cost) * 100).toFixed(1) : null;
  const previewProfit = sale.sale_price_rub ? (+sale.sale_price_rub) - cost : null;

  async function save(updated) { setSaving(true); await apiFetch('/api/shipments', { method: 'PUT', body: JSON.stringify(updated) }); onUpdate(updated); setSaving(false); }
  async function doArrival() { await save({ ...shipment, status: 'arrived', extra_paid_rub: +arrival.extra_paid_rub || 0, delivery_rub: +arrival.delivery_rub || 0, arrived_date: arrival.arrived_date }); setShowArrival(false); }
  async function doSale() { await save({ ...shipment, status: 'sold', sale_price_rub: +sale.sale_price_rub || 0, sold_date: new Date().toISOString().slice(0, 10) }); setShowSale(false); }
  async function doDelete() { if (!confirm('Удалить поставку?')) return; await apiFetch('/api/shipments', { method: 'DELETE', body: JSON.stringify({ id: shipment.id }) }); onDelete(); }

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: '1.5rem' }}>
        <button onClick={onBack} style={{ padding: '6px 10px' }}>← Назад</button>
        <div style={{ flex: 1 }}><div style={{ fontSize: 18, fontWeight: 500 }}>{shipment.name}</div><div className="muted" style={{ fontSize: 12, marginTop: 2 }}>Создан {new Date(shipment.createdAt).toLocaleDateString('ru-RU')}</div></div>
        <Badge status={shipment.status} />
      </div>
      <div className="grid-4 mb-1">
        <MetricCard label="Закупка (CNY)" value={shipment.total_cny ? `¥ ${fmt(shipment.total_cny)}` : '—'} />
        <MetricCard label="Оплачено" value={`${fmt(shipment.paid_rub)} ₽`} />
        <MetricCard label="Доставка" value={shipment.delivery_rub ? `${fmt(shipment.delivery_rub)} ₽` : '—'} />
        {profit !== null ? <MetricCard label="Прибыль" value={`${profit >= 0 ? '+' : ''}${fmt(profit)} ₽`} cls={profit >= 0 ? 'success' : 'danger'} sub={margin ? `${margin}% маржа` : null} /> : <MetricCard label="Себестоимость" value={`${fmt(cost)} ₽`} />}
      </div>
      <div className="card mb-1">
        <div className="muted" style={{ fontSize: 12, fontWeight: 500, marginBottom: 8 }}>ДЕТАЛИ</div>
        <table className="items-table"><tbody>
          {shipment.invoice_number && <Row label="№ накладной" value={shipment.invoice_number} />}
          {shipment.supplier && <Row label="Поставщик" value={shipment.supplier} />}
          {shipment.cny_rate > 0 && <Row label="Курс CNY" value={`${shipment.cny_rate} ₽/¥`} />}
          {shipment.ship_date && <Row label="Отправлено" value={shipment.ship_date} />}
          {shipment.eta_date && <Row label="Ожидается" value={shipment.eta_date} />}
          {shipment.arrived_date && <Row label="Прибыл" value={shipment.arrived_date} />}
          {shipment.extra_paid_rub > 0 && <Row label="Доп. оплата" value={`${fmt(shipment.extra_paid_rub)} ₽`} />}
          {shipment.sale_price_rub > 0 && <Row label="Продано за" value={`${fmt(shipment.sale_price_rub)} ₽`} />}
          {shipment.sold_date && <Row label="Дата продажи" value={shipment.sold_date} />}
        </tbody></table>
      </div>
      {shipment.items?.length > 0 && (
        <div className="card mb-1">
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: showItems ? 10 : 0 }}>
            <div className="muted" style={{ fontSize: 12, fontWeight: 500 }}>ПОЗИЦИИ ({shipment.items.length})</div>
            <button style={{ border: 'none', background: 'none', fontSize: 12, color: 'rgba(255,255,255,0.6)', cursor: 'pointer' }} onClick={() => setShowItems(v => !v)}>{showItems ? 'Скрыть' : 'Показать'}</button>
          </div>
          {showItems && shipment.items.map((item, i) => (
            <div key={i} style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, padding: '7px 0', borderBottom: i < shipment.items.length - 1 ? '0.5px solid rgba(255,255,255,0.1)' : 'none' }}>
              <div><div>{item.name}</div><div className="muted" style={{ fontSize: 11 }}>×{item.qty} · ¥{fmtF(item.unit_price_cny)}/шт</div></div>
              <div style={{ textAlign: 'right' }}><div>¥ {fmt(item.total_cny)}</div>{shipment.cny_rate > 0 && <div className="muted" style={{ fontSize: 11 }}>≈ {fmt(item.total_cny * shipment.cny_rate)} ₽</div>}</div>
            </div>
          ))}
        </div>
      )}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {shipment.status === 'new' && <button className="primary" onClick={() => save({ ...shipment, status: 'transit' })} disabled={saving} style={{ width: '100%', justifyContent: 'center' }}>🚛 В пути</button>}
        {shipment.status === 'transit' && !showArrival && <button className="primary" onClick={() => setShowArrival(true)} style={{ width: '100%', justifyContent: 'center' }}>📦 Груз прибыл</button>}
        {showArrival && (
          <div className="card">
            <div style={{ fontWeight: 500, marginBottom: 12 }}>Оформление прибытия</div>
            <Field label="Дата прибытия"><input type="date" value={arrival.arrived_date} onChange={e => setArrival(f => ({ ...f, arrived_date: e.target.value }))} /></Field>
            <Field label="Доп. оплата в России, ₽"><input type="number" value={arrival.extra_paid_rub} onChange={e => setArrival(f => ({ ...f, extra_paid_rub: e.target.value }))} placeholder="0" /></Field>
            <Field label="Доставка, ₽"><input type="number" value={arrival.delivery_rub} onChange={e => setArrival(f => ({ ...f, delivery_rub: e.target.value }))} placeholder="0" /></Field>
            <div className="row"><button onClick={() => setShowArrival(false)}>Отмена</button><button className="primary flex-1" onClick={doArrival} disabled={saving}>Сохранить</button></div>
          </div>
        )}
        {shipment.status === 'arrived' && !showSale && <button className="primary" onClick={() => setShowSale(true)} style={{ width: '100%', justifyContent: 'center' }}>💰 Продажная цена</button>}
        {showSale && (
          <div className="card">
            <div style={{ fontWeight: 500, marginBottom: 12 }}>Продажа</div>
            <Field label="Продажная цена, ₽"><input type="number" value={sale.sale_price_rub} onChange={e => setSale(f => ({ ...f, sale_price_rub: e.target.value }))} placeholder="0" /></Field>
            {previewProfit !== null && <div className={`profit-preview ${previewProfit >= 0 ? 'pos' : 'neg'}`}>Прибыль: {previewProfit >= 0 ? '+' : ''}{fmt(previewProfit)} ₽{cost > 0 ? ` (${((previewProfit / cost) * 100).toFixed(1)}%)` : ''}</div>}
            <div className="row"><button onClick={() => setShowSale(false)}>Отмена</button><button className="primary flex-1" onClick={doSale} disabled={saving}>Сохранить</button></div>
          </div>
        )}
        <button className="danger" onClick={doDelete} style={{ marginTop: 4 }}>🗑 Удалить поставку</button>
      </div>
    </div>
  );
}

// ─── WAREHOUSE ────────────────────────────────────────────────────────────────

function AddWarehouseItem({ onSave, onCancel }) {
  const [form, setForm] = useState({ name: '', qty: '', unit: 'шт', dimensions: '', sell_price: '', cost_rub: '', photo: null });
  const [saving, setSaving] = useState(false);
  const photoRef = useRef();
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  async function handlePhoto(file) { set('photo', await compressPhoto(file)); }

  async function handleSave() {
    if (!form.name || !form.qty) return;
    setSaving(true);
    const item = { id: genId(), createdAt: new Date().toISOString(), name: form.name, qty: +form.qty, unit: form.unit || 'шт', dimensions: form.dimensions, sell_price: +form.sell_price || 0, cost_rub: +form.cost_rub || 0, photo: form.photo, min_qty: 5, moves: [{ date: new Date().toISOString(), type: 'in', qty: +form.qty, note: 'Начальный остаток' }] };
    await apiFetch('/api/warehouse', { method: 'POST', body: JSON.stringify(item) });
    onSave(item);
    setSaving(false);
  }

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: '1.5rem' }}>
        <button onClick={onCancel} style={{ padding: '6px 10px' }}>← Назад</button>
        <div style={{ fontSize: 18, fontWeight: 500 }}>Новый товар</div>
      </div>
      <div style={{ marginBottom: '1rem' }}>
        {form.photo ? (
          <div style={{ position: 'relative' }}>
            <img src={form.photo} alt="" style={{ width: '100%', maxHeight: 200, objectFit: 'cover', borderRadius: 12 }} />
            <button onClick={() => set('photo', null)} style={{ position: 'absolute', top: 8, right: 8, padding: '4px 8px', fontSize: 12 }}>✕</button>
          </div>
        ) : (
          <div onClick={() => photoRef.current.click()} style={{ border: '1.5px dashed rgba(255,255,255,0.3)', borderRadius: 12, padding: '1.5rem', textAlign: 'center', cursor: 'pointer' }}>
            <div style={{ fontSize: 28, marginBottom: 6 }}>📷</div><div style={{ fontSize: 13 }}>Добавить фото товара</div>
          </div>
        )}
        <input ref={photoRef} type="file" accept="image/*" style={{ display: 'none' }} onChange={e => e.target.files[0] && handlePhoto(e.target.files[0])} />
      </div>
      <Field label="Название товара *"><input value={form.name} onChange={e => set('name', e.target.value)} placeholder="Кофейная чашка 210мл" /></Field>
      <div className="grid-2">
        <Field label="Количество *"><input type="number" value={form.qty} onChange={e => set('qty', e.target.value)} placeholder="0" /></Field>
        <Field label="Единица"><select value={form.unit} onChange={e => set('unit', e.target.value)}><option>шт</option><option>кг</option><option>м</option><option>л</option><option>упак</option><option>пара</option></select></Field>
      </div>
      <Field label="Размеры"><input value={form.dimensions} onChange={e => set('dimensions', e.target.value)} placeholder="11.4×8.7×5.9 см" /></Field>
      <div className="grid-2">
        <Field label="Себестоимость, ₽"><input type="number" value={form.cost_rub} onChange={e => set('cost_rub', e.target.value)} placeholder="0" /></Field>
        <Field label="Цена продажи, ₽"><input type="number" value={form.sell_price} onChange={e => set('sell_price', e.target.value)} placeholder="0" /></Field>
      </div>
      <div className="row mt-1">
        <button onClick={onCancel}>Отмена</button>
        <button className="primary flex-1" onClick={handleSave} disabled={saving || !form.name || !form.qty}>{saving ? 'Сохранение...' : 'Сохранить'}</button>
      </div>
    </div>
  );
}

function WarehouseItemDetail({ item, onUpdate, onDelete, onBack }) {
  const [showSale, setShowSale] = useState(false);
  const [showReceive, setShowReceive] = useState(false);
  const [saleQty, setSaleQty] = useState('');
  const [saleNote, setSaleNote] = useState('');
  const [recQty, setRecQty] = useState('');
  const [recNote, setRecNote] = useState('');
  const [parsing, setParsing] = useState(false);
  const [saving, setSaving] = useState(false);
  const photoRef = useRef();

  async function handleSale() {
    if (!saleQty || +saleQty > item.qty || +saleQty <= 0) return;
    setSaving(true);
    const updated = { ...item, qty: item.qty - +saleQty, moves: [...(item.moves || []), { date: new Date().toISOString(), type: 'out', qty: +saleQty, note: saleNote || 'Продажа' }] };
    await apiFetch('/api/warehouse', { method: 'PUT', body: JSON.stringify(updated) });
    onUpdate(updated); setShowSale(false); setSaleQty(''); setSaleNote(''); setSaving(false);
  }

  async function handleReceive() {
    if (!recQty) return;
    setSaving(true);
    const updated = { ...item, qty: item.qty + +recQty, moves: [...(item.moves || []), { date: new Date().toISOString(), type: 'in', qty: +recQty, note: recNote || 'Поступление' }] };
    await apiFetch('/api/warehouse', { method: 'PUT', body: JSON.stringify(updated) });
    onUpdate(updated); setShowReceive(false); setRecQty(''); setRecNote(''); setSaving(false);
  }

  async function handleSaleInvoice(file) {
    setParsing(true);
    try {
      const data = await parseInvoiceFile(file);
      const keyword = item.name.toLowerCase().split(' ')[0];
      const match = data.items?.find(i => i.name && i.name.toLowerCase().includes(keyword));
      if (match && match.qty) setSaleQty(String(match.qty));
      setSaleNote('Из накладной о продаже');
      setShowSale(true);
    } catch {}
    setParsing(false);
  }

  async function handlePhoto(file) {
    const compressed = await compressPhoto(file);
    const updated = { ...item, photo: compressed };
    await apiFetch('/api/warehouse', { method: 'PUT', body: JSON.stringify(updated) });
    onUpdate(updated);
  }

  async function doDelete() {
    if (!confirm('Удалить товар?')) return;
    await apiFetch('/api/warehouse', { method: 'DELETE', body: JSON.stringify({ id: item.id }) });
    onDelete();
  }

  const qtyColor = item.qty === 0 ? '#fca5a5' : item.qty <= (item.min_qty || 5) ? '#FDE68A' : '#86EFAC';

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: '1.5rem' }}>
        <button onClick={onBack} style={{ padding: '6px 10px' }}>← Назад</button>
        <div style={{ flex: 1 }}><div style={{ fontSize: 18, fontWeight: 500 }}>{item.name}</div>{item.dimensions && <div className="muted" style={{ fontSize: 12, marginTop: 2 }}>{item.dimensions}</div>}</div>
        <div style={{ fontSize: 22, fontWeight: 700, color: qtyColor }}>{item.qty} {item.unit}</div>
      </div>

      {item.photo ? (
        <div style={{ position: 'relative', marginBottom: '1rem' }}>
          <img src={item.photo} alt={item.name} style={{ width: '100%', maxHeight: 220, objectFit: 'cover', borderRadius: 12 }} />
          <button onClick={() => photoRef.current.click()} style={{ position: 'absolute', bottom: 8, right: 8, fontSize: 12 }}>📷 Изменить</button>
        </div>
      ) : (
        <div onClick={() => photoRef.current.click()} style={{ border: '1.5px dashed rgba(255,255,255,0.3)', borderRadius: 12, padding: '1rem', textAlign: 'center', cursor: 'pointer', marginBottom: '1rem' }}>
          📷 Добавить фото
        </div>
      )}
      <input ref={photoRef} type="file" accept="image/*" style={{ display: 'none' }} onChange={e => e.target.files[0] && handlePhoto(e.target.files[0])} />

      <div className="grid-4 mb-1">
        <MetricCard label="На складе" value={`${item.qty} ${item.unit}`} cls={item.qty === 0 ? 'danger' : ''} />
        <MetricCard label="Себестоимость/шт" value={item.cost_rub ? `${fmt(item.cost_rub)} ₽` : '—'} />
        <MetricCard label="Цена продажи" value={item.sell_price ? `${fmt(item.sell_price)} ₽` : '—'} />
        <MetricCard label="Стоимость остатка" value={item.cost_rub && item.qty ? `${fmt(item.cost_rub * item.qty)} ₽` : '—'} />
      </div>

      <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: '1rem' }}>
        {!showSale && !showReceive && (
          <>
            <button className="primary" onClick={() => setShowSale(true)} style={{ width: '100%', justifyContent: 'center' }}>📤 Списать (продажа)</button>
            <button onClick={() => setShowReceive(true)} style={{ width: '100%', justifyContent: 'center' }}>📥 Принять на склад</button>
          </>
        )}

        {showSale && (
          <div className="card">
            <div style={{ fontWeight: 500, marginBottom: 10 }}>Списание со склада</div>
            <DropZone onFile={handleSaleInvoice} loading={parsing} label="Загрузить накладную о продаже" />
            <div className="muted" style={{ fontSize: 12, textAlign: 'center', margin: '6px 0 10px' }}>или введите вручную</div>
            <Field label={`Количество (на складе: ${item.qty} ${item.unit})`}>
              <input type="number" value={saleQty} onChange={e => setSaleQty(e.target.value)} max={item.qty} placeholder="0" />
            </Field>
            <Field label="Комментарий"><input value={saleNote} onChange={e => setSaleNote(e.target.value)} placeholder="Продажа" /></Field>
            <div className="row"><button onClick={() => setShowSale(false)}>Отмена</button><button className="primary flex-1" onClick={handleSale} disabled={saving || !saleQty || +saleQty > item.qty || +saleQty <= 0}>Списать</button></div>
          </div>
        )}

        {showReceive && (
          <div className="card">
            <div style={{ fontWeight: 500, marginBottom: 12 }}>Приход на склад</div>
            <Field label="Количество"><input type="number" value={recQty} onChange={e => setRecQty(e.target.value)} placeholder="0" /></Field>
            <Field label="Комментарий"><input value={recNote} onChange={e => setRecNote(e.target.value)} placeholder="Поступление" /></Field>
            <div className="row"><button onClick={() => setShowReceive(false)}>Отмена</button><button className="primary flex-1" onClick={handleReceive} disabled={saving || !recQty}>Принять</button></div>
          </div>
        )}
      </div>

      {item.moves?.length > 0 && (
        <div className="card mb-1">
          <div className="muted" style={{ fontSize: 12, fontWeight: 500, marginBottom: 8 }}>ИСТОРИЯ ДВИЖЕНИЙ</div>
          {[...item.moves].reverse().slice(0, 15).map((m, i, arr) => (
            <div key={i} style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, padding: '6px 0', borderBottom: i < arr.length - 1 ? '0.5px solid rgba(255,255,255,0.1)' : 'none' }}>
              <div><span style={{ color: m.type === 'in' ? '#86EFAC' : '#fca5a5', marginRight: 8 }}>{m.type === 'in' ? '↑' : '↓'}</span>{m.note || (m.type === 'in' ? 'Поступление' : 'Списание')}</div>
              <div style={{ textAlign: 'right' }}>
                <div style={{ color: m.type === 'in' ? '#86EFAC' : '#fca5a5', fontWeight: 500 }}>{m.type === 'in' ? '+' : '-'}{m.qty} {item.unit}</div>
                <div className="muted" style={{ fontSize: 11 }}>{new Date(m.date).toLocaleDateString('ru-RU')}</div>
              </div>
            </div>
          ))}
        </div>
      )}

      <button className="danger" onClick={doDelete} style={{ width: '100%', justifyContent: 'center', marginTop: 4 }}>🗑 Удалить товар</button>
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────

export default function Home() {
  const [tab, setTab] = useState('shipments');
  const [view, setView] = useState('list');
  const [shipments, setShipments] = useState([]);
  const [warehouse, setWarehouse] = useState([]);
  const [selected, setSelected] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    Promise.all([
      fetch('/api/shipments').then(r => r.json()).catch(() => []),
      fetch('/api/warehouse').then(r => r.json()).catch(() => []),
    ]).then(([s, w]) => { setShipments(s || []); setWarehouse(w || []); setLoading(false); });
  }, []);

  function updateShipment(u) { setShipments(s => s.map(x => x.id === u.id ? u : x)); }
  function updateWarehouseItem(u) { setWarehouse(w => w.map(x => x.id === u.id ? u : x)); }

  if (loading) return <div className="container" style={{ color: 'rgba(255,255,255,0.6)', fontSize: 14, paddingTop: '3rem' }}>Загрузка...</div>;

  // Shipment new/detail
  if (view === 'new-shipment') return <div className="container"><NewShipment onSave={s => { setShipments(p => [...p, s]); setView('list'); }} onCancel={() => setView('list')} /></div>;
  if (view === 'detail-shipment') {
    const s = shipments.find(s => s.id === selected);
    if (!s) { setView('list'); return null; }
    return <div className="container"><ShipmentDetail shipment={s} onUpdate={updateShipment} onDelete={() => { setShipments(p => p.filter(x => x.id !== selected)); setView('list'); }} onBack={() => setView('list')} /></div>;
  }

  // Warehouse new/detail
  if (view === 'new-warehouse') return <div className="container"><AddWarehouseItem onSave={i => { setWarehouse(p => [...p, i]); setView('list'); }} onCancel={() => setView('list')} /></div>;
  if (view === 'detail-warehouse') {
    const item = warehouse.find(i => i.id === selected);
    if (!item) { setView('list'); return null; }
    return <div className="container"><WarehouseItemDetail item={item} onUpdate={updateWarehouseItem} onDelete={() => { setWarehouse(p => p.filter(x => x.id !== selected)); setView('list'); }} onBack={() => setView('list')} /></div>;
  }

  // ── List views ──
  const totalProfit = shipments.filter(s => s.status === 'sold').reduce((sum, s) => sum + ((s.sale_price_rub || 0) - (s.paid_rub || 0) - (s.extra_paid_rub || 0) - (s.delivery_rub || 0)), 0);

  return (
    <div className="container">
      {/* Tab nav */}
      <div style={{ display: 'flex', background: 'rgba(255,255,255,0.1)', borderRadius: 10, padding: 3, marginBottom: '1.5rem' }}>
        {[['shipments', '🚢 Поставки'], ['warehouse', '📦 Склад']].map(([key, label]) => (
          <button key={key} onClick={() => setTab(key)} style={{ flex: 1, justifyContent: 'center', border: 'none', borderRadius: 8, background: tab === key ? 'rgba(255,255,255,0.9)' : 'transparent', color: tab === key ? '#0F4C75' : 'rgba(255,255,255,0.8)', fontWeight: tab === key ? 500 : 400, padding: '8px 0' }}>{label}</button>
        ))}
      </div>

      {/* SHIPMENTS TAB */}
      {tab === 'shipments' && (
        <div>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '1.5rem' }}>
            <div><div style={{ fontSize: 20, fontWeight: 500 }}>Поставки</div><div className="muted" style={{ fontSize: 13, marginTop: 2 }}>Китай → Россия</div></div>
            <button className="primary" onClick={() => setView('new-shipment')}>+ Новая</button>
          </div>
          {shipments.length > 0 && (
            <div className="grid-4" style={{ marginBottom: '1.5rem' }}>
              <MetricCard label="Поставок" value={shipments.length} />
              <MetricCard label="В пути" value={shipments.filter(s => s.status === 'transit').length} />
              <MetricCard label="Прибыло" value={shipments.filter(s => s.status === 'arrived').length} />
              <MetricCard label="Прибыль" value={`${totalProfit >= 0 ? '+' : ''}${fmt(totalProfit)} ₽`} cls={shipments.some(s => s.status === 'sold') ? (totalProfit >= 0 ? 'success' : 'danger') : ''} />
            </div>
          )}
          {shipments.length === 0 ? (
            <div className="empty-state"><div style={{ fontSize: 40, marginBottom: '1rem' }}>📦</div><div style={{ fontSize: 15, fontWeight: 500, marginBottom: 6 }}>Поставок нет</div><div className="muted" style={{ fontSize: 13 }}>Нажмите «Новая» чтобы начать</div></div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {[...shipments].reverse().map(s => {
                const cost = (s.paid_rub || 0) + (s.extra_paid_rub || 0) + (s.delivery_rub || 0);
                const profit = s.status === 'sold' ? (s.sale_price_rub || 0) - cost : null;
                const daysLeft = s.eta_date && s.status === 'transit' ? Math.ceil((new Date(s.eta_date) - new Date()) / 86400000) : null;
                return (
                  <div key={s.id} className="card shipment-row" onClick={() => { setSelected(s.id); setView('detail-shipment'); }}>
                    <div style={{ fontSize: 24 }}>📦</div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 500, fontSize: 15, marginBottom: 3 }}>{s.name}</div>
                      <div className="muted" style={{ fontSize: 12 }}>
                        {s.total_cny ? `¥ ${fmt(s.total_cny)}` : ''}{s.items?.length ? ` · ${s.items.length} поз.` : ''}{daysLeft !== null ? ` · ${daysLeft > 0 ? `${daysLeft} дн.` : 'сегодня'}` : ''}{s.ship_date && !daysLeft ? ` · ${s.ship_date}` : ''}
                      </div>
                    </div>
                    <div style={{ textAlign: 'right', flexShrink: 0 }}>
                      <Badge status={s.status} />
                      {profit !== null && <div className={profit >= 0 ? 'success' : 'danger'} style={{ fontSize: 13, marginTop: 4, fontWeight: 500 }}>{profit >= 0 ? '+' : ''}{fmt(profit)} ₽</div>}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}

      {/* WAREHOUSE TAB */}
      {tab === 'warehouse' && (
        <div>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '1.5rem' }}>
            <div><div style={{ fontSize: 20, fontWeight: 500 }}>Склад</div><div className="muted" style={{ fontSize: 13, marginTop: 2 }}>{warehouse.length} позиций</div></div>
            <button className="primary" onClick={() => setView('new-warehouse')}>+ Товар</button>
          </div>

          {warehouse.length > 0 && (
            <div className="grid-4" style={{ marginBottom: '1.5rem' }}>
              <MetricCard label="Позиций" value={warehouse.length} />
              <MetricCard label="Единиц всего" value={fmt(warehouse.reduce((s, i) => s + (i.qty || 0), 0))} />
              <MetricCard label="Мало" value={warehouse.filter(i => (i.qty || 0) <= (i.min_qty || 5) && (i.qty || 0) > 0).length} cls={warehouse.filter(i => (i.qty||0) <= (i.min_qty||5) && (i.qty||0) > 0).length > 0 ? 'danger' : ''} />
              <MetricCard label="Нет в наличии" value={warehouse.filter(i => (i.qty || 0) === 0).length} cls={warehouse.filter(i => (i.qty||0) === 0).length > 0 ? 'danger' : ''} />
            </div>
          )}

          {warehouse.length === 0 ? (
            <div className="empty-state"><div style={{ fontSize: 40, marginBottom: '1rem' }}>🏪</div><div style={{ fontSize: 15, fontWeight: 500, marginBottom: 6 }}>Склад пуст</div><div className="muted" style={{ fontSize: 13 }}>Нажмите «Товар» чтобы добавить</div></div>
          ) : (
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(155px, 1fr))', gap: 12 }}>
              {warehouse.map(item => {
                const qtyColor = item.qty === 0 ? '#fca5a5' : item.qty <= (item.min_qty || 5) ? '#FDE68A' : '#86EFAC';
                return (
                  <div key={item.id} onClick={() => { setSelected(item.id); setView('detail-warehouse'); }}
                    style={{ background: 'rgba(255,255,255,0.1)', border: '1px solid rgba(255,255,255,0.15)', borderRadius: 12, overflow: 'hidden', cursor: 'pointer', transition: 'background 0.15s' }}>
                    {item.photo
                      ? <img src={item.photo} alt={item.name} style={{ width: '100%', height: 130, objectFit: 'cover' }} />
                      : <div style={{ width: '100%', height: 130, background: 'rgba(255,255,255,0.05)', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 36 }}>📦</div>
                    }
                    <div style={{ padding: '10px 12px' }}>
                      <div style={{ fontWeight: 500, fontSize: 13, marginBottom: 4, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{item.name}</div>
                      <div style={{ fontSize: 14, color: qtyColor, fontWeight: 500 }}>{item.qty} {item.unit}</div>
                      {item.dimensions && <div className="muted" style={{ fontSize: 11, marginTop: 2 }}>{item.dimensions}</div>}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}
    </div>
  );
}
