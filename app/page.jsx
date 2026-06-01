'use client';
import { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';

function genId() { return Date.now().toString(36) + Math.random().toString(36).slice(2); }
function fmt(n) { if (!n && n !== 0) return '—'; return Math.round(n).toLocaleString('ru-RU'); }
function fmtF(n, d = 2) { if (!n && n !== 0) return '—'; return (+n).toFixed(d); }

// ─── SVG ICONS ───────────────────────────────────────────────────────────────
const IconBox = ({size=22,color="currentColor"}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/></svg>;
const IconShip = ({size=22,color="currentColor"}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M2 21c.6.5 1.2 1 2.5 1C7 22 7 21 9.5 21s2.5 1 5 1 2.5-1 5-1c1.3 0 1.9.5 2.5 1"/><path d="M19.38 20A11.6 11.6 0 0 0 21 14l-9-4-9 4c0 3.2 2 6.5 4.62 8"/><path d="M19 13V7.58a2 2 0 0 0-1.28-1.87L12 3.58 6.28 5.71A2 2 0 0 0 5 7.58V13"/><line x1="12" y1="3.58" x2="12" y2="12"/></svg>;
const IconMoney = ({size=22,color="currentColor"}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>;
const IconTrash = ({size=18,color="currentColor"}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>;
const IconEdit = ({size=16,color="currentColor"}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>;
const IconBack = ({size=18,color="currentColor"}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="19" y1="12" x2="5" y2="12"/><polyline points="12 19 5 12 12 5"/></svg>;
const IconCamera = ({size=28,color="currentColor"}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/><circle cx="12" cy="13" r="4"/></svg>;
const IconRefresh = ({size=18,color="currentColor"}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg>;
const IconDoc = ({size=28,color="currentColor"}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>;
const IconTruck = ({size=20,color="currentColor"}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><rect x="1" y="3" width="15" height="13"/><polygon points="16 8 20 8 23 11 23 16 16 16 16 8"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/></svg>;
const IconArrowDown = ({size=18,color="currentColor"}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><polyline points="19 12 12 19 5 12"/></svg>;
const IconArrowUp = ({size=18,color="currentColor"}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="19" x2="12" y2="5"/><polyline points="5 12 12 5 19 12"/></svg>;
const IconChart = ({size=18,color="currentColor"}) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke={color} strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>;

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
  } else { fd.append('file', file); }
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
      if (w > max || h > max) { if (w > h) { h = Math.round(h * max / w); w = max; } else { w = Math.round(w * max / h); h = max; } }
      canvas.width = w; canvas.height = h;
      canvas.getContext('2d').drawImage(img, 0, 0, w, h);
      resolve(canvas.toDataURL('image/jpeg', 0.75));
    };
    img.src = URL.createObjectURL(file);
  });
}

const BADGES = { new: { label: 'Новый', cls: 'badge-new' }, transit: { label: 'В пути', cls: 'badge-transit' }, arrived: { label: 'Прибыл', cls: 'badge-arrived' }, sold: { label: 'Продан', cls: 'badge-sold' } };
function Badge({ status }) { const b = BADGES[status] || BADGES.new; return <span className={`badge ${b.cls}`}>{b.label}</span>; }
function Field({ label, children }) { return <div className="field"><label>{label}</label>{children}</div>; }
function Row({ label, value }) { return <tr><td className="muted" style={{ width: '50%' }}>{label}</td><td style={{ textAlign: 'right' }}>{value}</td></tr>; }
function MetricCard({ label, value, cls, sub }) { return <div className="metric-card"><div className="metric-label">{label}</div><div className={`metric-value ${cls || ''}`}>{value}</div>{sub && <div style={{ fontSize: 11, color: '#888', marginTop: 2 }}>{sub}</div>}</div>; }

function DropZone({ onFile, loading, label }) {
  const [drag, setDrag] = useState(false);
  const ref = useRef();
  return (
    <div onDragOver={e => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]); }}
      onClick={() => !loading && ref.current.click()}
      className={`drop-zone${drag ? ' drag' : ''}`}>
      <input ref={ref} type="file" accept=".pdf,.png,.jpg,.jpeg,.xlsx,.xls,.csv" style={{ display: 'none' }} onChange={e => e.target.files[0] && onFile(e.target.files[0])} />
      <div style={{ marginBottom: 8, color: "rgba(255,255,255,0.4)" }}><IconDoc size={32} /></div>
      {loading ? <div className="muted">Распознаём...</div> : <><div style={{ fontWeight: 500, marginBottom: 4 }}>{label || 'Загрузите накладную'}</div><div className="muted" style={{ fontSize: 12 }}>PDF · Фото · Excel</div></>}
    </div>
  );
}

// ─── SHIPMENTS ────────────────────────────────────────────────────────────────

function NewShipment({ onSave, onCancel }) {
  const [step, setStep] = useState(1);
  const [parsing, setParsing] = useState(false);
  const [parseError, setParseError] = useState('');
  const [saving, setSaving] = useState(false);
  const [form, setForm] = useState({ name: '', items: [], total_cny: '', invoice_number: '', supplier: '', cny_rate: '', paid_rub: '', ship_date: '', eta_date: '' });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  async function handleFile(file) {
    setParsing(true); setParseError('');
    try {
      const d = await parseInvoiceFile(file);
const expKw=['доставка','упаковка','поддон','транспорт','фрахт','тара','паллет','пломба','страховка','运费','木托','包装','配送'];
const isExp=(name)=>expKw.some(k=>(name||'').toLowerCase().includes(k.toLowerCase()));
const expCNY=(d.items||[]).filter(i=>isExp(i.name)).reduce((s,i)=>s+(i.total_cny||0),0);
      setForm(f => ({ ...f, items: d.items || [], total_cny: d.total_cny || '', invoice_number: d.invoice_number || '', supplier: d.supplier || '', _expense_cny: expCNY, name: d.supplier ? `Поставка — ${d.supplier}` : `Поставка ${new Date().toLocaleDateString('ru-RU')}` }));
      setStep(2);
    } catch { setParseError('Не удалось распознать. Введите вручную.'); }
    setParsing(false);
  }

  async function handleSave() {
    setSaving(true);
    const totalCNY = +form.total_cny || form.items.reduce((s, i) => s + (i.total_cny || 0), 0);
    const shipment = { id: genId(), createdAt: new Date().toISOString(), status: form.ship_date ? 'transit' : 'new', name: form.name || `Поставка ${new Date().toLocaleDateString('ru-RU')}`, items: form.items, total_cny: totalCNY, invoice_number: form.invoice_number, supplier: form.supplier, cny_rate: +form.cny_rate || 0, paid_rub: +form.paid_rub || 0, ship_date: form.ship_date, eta_date: form.eta_date, extra_paid_rub: 0, delivery_rub: 0, sale_price_rub: 0 };
    await apiFetch('/api/shipments', { method: 'POST', body: JSON.stringify(shipment) });
    onSave(shipment); setSaving(false);
  }

  const totalCNY = +form.total_cny || form.items.reduce((s, i) => s + (i.total_cny || 0), 0);

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: '1.5rem' }}>
        <button onClick={onCancel} style={{ padding: '6px 10px' }}><IconBack size={16} /> Назад</button>
        <div style={{ fontSize: 18, fontWeight: 600 }}>Новая поставка</div>
      </div>
      <div className="step-bar">
        {['Накладная','Оплата','Доставка','Итог'].map((s, i) => (
          <div key={i} style={{ flex: 1 }}>
            <div className="step-seg" style={{ background: i < step ? 'linear-gradient(135deg,#0077B6,#48CAE4)' : '#ddd', marginBottom: 3 }} />
            {i === step - 1 && <div style={{ fontSize: 11, color: '#0077B6', textAlign: 'center', fontWeight: 500 }}>{s}</div>}
          </div>
        ))}
      </div>
      {step === 1 && (
        <div>
          <DropZone onFile={handleFile} loading={parsing} />
          {parseError && <div className="error-box">{parseError}</div>}
          <div className="mt-1" style={{ textAlign: 'center' }}>
            <button style={{ border: 'none', background: 'none', color: '#0077B6', cursor: 'pointer', fontSize: 13 }} onClick={() => { set('name', `Поставка ${new Date().toLocaleDateString('ru-RU')}`); setStep(2); }}>Пропустить →</button>
          </div>
        </div>
      )}
      {step === 2 && (
        <div>
          <Field label="Название поставки"><input value={form.name} onChange={e => set('name', e.target.value)} placeholder="Электроника январь" /></Field>
          <Field label="Итого, ¥"><input type="number" value={form.total_cny} onChange={e => set('total_cny', e.target.value)} placeholder="0" /></Field>
          <div className="grid-2">
            <Field label="Курс CNY/RUB"><input type="number" value={form.cny_rate} onChange={e => set('cny_rate', e.target.value)} placeholder="13.5" /></Field>
            <Field label="Оплачено, ₽"><input type="number" value={form.paid_rub} onChange={e => set('paid_rub', e.target.value)} placeholder={form.total_cny && form.cny_rate ? String(Math.round(+form.total_cny * +form.cny_rate)) : '0'} onFocus={e => { if (!form.paid_rub && form.total_cny && form.cny_rate) set('paid_rub', Math.round(+form.total_cny * +form.cny_rate)); }} /></Field>
          </div>
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
            <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 12 }}>{form.name}</div>
            <table className="items-table"><tbody>
              {totalCNY > 0 && <Row label="Закупка" value={`¥ ${fmt(totalCNY)}`} />}
              {form.cny_rate && <Row label="Курс" value={`${form.cny_rate} ₽/¥`} />}
              {form.paid_rub && <Row label="Оплачено" value={`${fmt(+form.paid_rub)} ₽`} />}
              {form.ship_date && <Row label="Отправлено" value={form.ship_date} />}

            </tbody></table>
            {form.items.length > 0 && <><hr className="divider" /><div className="muted" style={{ fontSize: 12, marginBottom: 6 }}>Позиции ({form.items.length}):</div>{form.items.slice(0,4).map((item,i) => (<div key={i} style={{ fontSize: 12, display: 'flex', justifyContent: 'space-between', padding: '2px 0' }}><span>{item.name}</span><span className="muted">×{item.qty}</span></div>))}</>}
          </div>
          <div className="row"><button onClick={() => setStep(3)}>← Назад</button><button className="primary flex-1" onClick={handleSave} disabled={saving}>{saving ? 'Сохранение...' : 'Сохранить поставку'}</button></div>
        </div>
      )}
    </div>
  );
}

function ShipmentDetail({ shipment, onUpdate, onDelete, onBack, onWarehouseUpdate }) {
  const [showArrival, setShowArrival] = useState(false);
  const [showSale, setShowSale] = useState(false);
  const [showItems, setShowItems] = useState(false);
  const [showEdit, setShowEdit] = useState(false);
  const [saving, setSaving] = useState(false);
  const [arrival, setArrival] = useState({ extra_paid_rub: '', delivery_rub: '', arrived_date: new Date().toISOString().slice(0, 10) });
  const [sale, setSale] = useState({ sale_price_rub: '' });
  const [editShipment, setEditShipment] = useState({ paid_rub: shipment.paid_rub || '', cny_rate: shipment.cny_rate || '', total_cny: shipment.total_cny || '', name: shipment.name || '' });
  const setES = (k, v) => setEditShipment(f => ({ ...f, [k]: v }));

  async function saveShipmentEdit() {
    setSaving(true);
    const updated = { ...shipment, paid_rub: +editShipment.paid_rub || 0, cny_rate: +editShipment.cny_rate || 0, total_cny: +editShipment.total_cny || 0, name: editShipment.name };
    await apiFetch('/api/shipments', { method: 'PUT', body: JSON.stringify(updated) });
    onUpdate(updated); setShowEdit(false); setSaving(false);
  }

  const cost = (shipment.paid_rub || 0) + (shipment.extra_paid_rub || 0) + (shipment.delivery_rub || 0);
  const profit = shipment.status === 'sold' ? (shipment.sale_price_rub || 0) - (shipment.usn_tax || 0) - cost : null;
  const margin = profit !== null && cost > 0 ? ((profit / cost) * 100).toFixed(1) : null;
  const previewProfit = sale.sale_price_rub ? (+sale.sale_price_rub) - cost : null;

  async function save(updated) { setSaving(true); await apiFetch('/api/shipments', { method: 'PUT', body: JSON.stringify(updated) }); onUpdate(updated); setSaving(false); }

  async function doArrival() {
    setSaving(true);
    const updated = { ...shipment, status: 'arrived', extra_paid_rub: +arrival.extra_paid_rub || 0, delivery_rub: +arrival.delivery_rub || 0, arrived_date: arrival.arrived_date };
    await apiFetch('/api/shipments', { method: 'PUT', body: JSON.stringify(updated) });

    // Автоматически добавляем товары на склад (без расходных позиций)
    if (shipment.items?.length > 0) {
      const expenseKeywords = ['доставка', 'упаковка', 'поддон', 'транспорт', 'фрахт', 'тара', 'паллет', 'пломба', 'страховка', '运费', '木托', '包装'];
      const isExpense = (name) => expenseKeywords.some(k => name.toLowerCase().includes(k.toLowerCase()));
      const warehouseItems = shipment.items.filter(i => i.name && i.qty > 0 && !isExpense(i.name)).map(item => ({
        id: genId(),
        createdAt: new Date().toISOString(),
        name: item.name,
        qty: item.qty,
        unit: 'шт',
        dimensions: '',
        cost_rub: (() => {
const baseRate = shipment.cny_rate > 0 ? shipment.cny_rate : 0;
const baseCost = item.unit_price_cny ? item.unit_price_cny * baseRate : 0;
// Distribute expenses proportionally by item value
const expKw = ['доставка','упаковка','поддон','транспорт','фрахт','тара','паллет','пломба','страховка','运费','木托','包装','配送'];
const isExp = (n) => expKw.some(k => (n||'').toLowerCase().includes(k.toLowerCase()));
const goodsItems = shipment.items.filter(i => !isExp(i.name));
const expenseItems = shipment.items.filter(i => isExp(i.name));
const totalGoodsCNY = goodsItems.reduce((s, i) => s + (i.total_cny || 0), 0);
const totalExpenseCNY = expenseItems.reduce((s, i) => s + (i.total_cny || 0), 0);
const itemShare = totalGoodsCNY > 0 ? (item.total_cny || 0) / totalGoodsCNY : 0;
const expensePerUnit = item.qty > 0 ? (itemShare * totalExpenseCNY * baseRate) / item.qty : 0;
return Math.round(baseCost + expensePerUnit);
})(),
        sell_price: 0,
        photo: null,
        min_qty: 5,
        shipment_id: shipment.id,
        moves: [{ date: new Date().toISOString(), type: 'in', qty: item.qty, note: `Из поставки: ${shipment.name}` }]
      }));
      await Promise.all(warehouseItems.map(w => apiFetch('/api/warehouse', { method: 'POST', body: JSON.stringify(w) })));
      if (onWarehouseUpdate) onWarehouseUpdate(warehouseItems);
    }

    onUpdate(updated);
    setShowArrival(false);
    setSaving(false);
  }

  async function doSale() {
    const revenue = +sale.sale_price_rub || 0;
    const tax = Math.round(revenue * 0.06);
    await save({ ...shipment, status: 'sold', sale_price_rub: revenue, usn_tax: tax, sold_date: new Date().toISOString().slice(0, 10) });
    setShowSale(false);
  }
  async function doDelete() { if (!confirm('Удалить поставку?')) return; await apiFetch('/api/shipments', { method: 'DELETE', body: JSON.stringify({ id: shipment.id }) }); onDelete(); }

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: '1.5rem' }}>
        <button onClick={onBack} style={{ padding: '6px 10px' }}>← Назад</button>
        <div style={{ flex: 1 }}><div style={{ fontSize: 18, fontWeight: 600 }}>{shipment.name}</div><div className="muted" style={{ fontSize: 12, marginTop: 2 }}>Создан {new Date(shipment.createdAt).toLocaleDateString('ru-RU')}</div></div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <Badge status={shipment.status} />
          <button onClick={() => setShowEdit(v => !v)} style={{ fontSize: 12, padding: '4px 10px', color: '#0077B6', borderColor: '#cce0f0' }}>{showEdit ? '✕' : <IconEdit size={15} />}</button>
        </div>
      </div>
      {showEdit && (
        <div className="card mb-1" style={{ borderColor: '#0077B6' }}>
          <div style={{ fontWeight: 600, marginBottom: 12, color: '#0077B6' }}>Редактирование поставки</div>
          <Field label="Название"><input value={editShipment.name} onChange={e => setES('name', e.target.value)} /></Field>
          <div className="grid-2">
            <Field label="Сумма закупки, ¥"><input type="number" value={editShipment.total_cny} onChange={e => setES('total_cny', e.target.value)} placeholder="0" /></Field>
            <Field label="Курс CNY/RUB"><input type="number" value={editShipment.cny_rate} onChange={e => setES('cny_rate', e.target.value)} placeholder="13.5" /></Field>
          </div>
          <Field label="Оплачено в рублях">
            <input type="number" value={editShipment.paid_rub} onChange={e => setES('paid_rub', e.target.value)}
              placeholder={editShipment.total_cny && editShipment.cny_rate ? String(Math.round(+editShipment.total_cny * +editShipment.cny_rate)) : '0'}
              onFocus={e => { if (!editShipment.paid_rub && editShipment.total_cny && editShipment.cny_rate) setES('paid_rub', Math.round(+editShipment.total_cny * +editShipment.cny_rate)); }} />
          </Field>
          {editShipment.total_cny && editShipment.cny_rate && <div className="muted mb-1" style={{ fontSize: 13 }}>≈ {Math.round(+editShipment.total_cny * +editShipment.cny_rate).toLocaleString('ru-RU')} ₽</div>}
          <div className="row">
            <button onClick={() => setShowEdit(false)}>Отмена</button>
            <button className="primary flex-1" onClick={saveShipmentEdit} disabled={saving}>Сохранить</button>
          </div>
        </div>
      )}
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
          {shipment.usn_tax > 0 && <Row label="УСН 6%" value={`−${fmt(shipment.usn_tax)} ₽`} />}
          {shipment.sale_price_rub > 0 && profit !== null && <Row label="Каждому 👤" value={`${Math.round(profit/2) >= 0 ? '+' : ''}${fmt(Math.round(profit/2))} ₽`} />}
          {shipment.sold_date && <Row label="Дата продажи" value={shipment.sold_date} />}
        </tbody></table>
      </div>
      {shipment.items?.length > 0 && (
        <div className="card mb-1">
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: showItems ? 10 : 0 }}>
            <div className="muted" style={{ fontSize: 12, fontWeight: 500 }}>ПОЗИЦИИ ({shipment.items.length})</div>
            <button style={{ border: 'none', background: 'none', fontSize: 12, color: '#0077B6', cursor: 'pointer' }} onClick={() => setShowItems(v => !v)}>{showItems ? 'Скрыть' : 'Показать'}</button>
          </div>
          {showItems && shipment.items.map((item, i) => (
            <div key={i} style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, padding: '7px 0', borderBottom: i < shipment.items.length - 1 ? '1px solid #f0f0f0' : 'none' }}>
              <div><div>{item.name}</div><div className="muted" style={{ fontSize: 11 }}>×{item.qty} · ¥{fmtF(item.unit_price_cny)}/шт</div></div>
              <div style={{ textAlign: 'right' }}><div>¥ {fmt(item.total_cny)}</div>{shipment.cny_rate > 0 && <div className="muted" style={{ fontSize: 11 }}>≈ {fmt(item.total_cny * shipment.cny_rate)} ₽</div>}</div>
            </div>
          ))}
        </div>
      )}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {shipment.status === 'new' && <button className="primary" onClick={() => save({ ...shipment, status: 'transit' })} disabled={saving} style={{ width: '100%', justifyContent: 'center' }}><IconTruck size={18} /> Отправлено — в пути</button>}
        {shipment.status === 'transit' && !showArrival && (
          <button className="primary" onClick={() => setShowArrival(true)} style={{ width: '100%', justifyContent: 'center' }}><IconBox size={18} /> Груз прибыл → на склад</button>
        )}
        {showArrival && (
          <div className="card">
            <div style={{ fontWeight: 600, marginBottom: 6 }}>Оформление прибытия</div>
            <div className="info-box" style={{ fontSize: 12, marginBottom: 12 }}>
              ✅ {shipment.items?.length || 0} позиций автоматически добавятся на склад
            </div>
            <Field label="Дата прибытия"><input type="date" value={arrival.arrived_date} onChange={e => setArrival(f => ({ ...f, arrived_date: e.target.value }))} /></Field>
            <Field label="Доп. оплата в России, ₽"><input type="number" value={arrival.extra_paid_rub} onChange={e => setArrival(f => ({ ...f, extra_paid_rub: e.target.value }))} placeholder="0" /></Field>
            <Field label="Доставка, ₽"><input type="number" value={arrival.delivery_rub} onChange={e => setArrival(f => ({ ...f, delivery_rub: e.target.value }))} placeholder="0" /></Field>
            <div className="row"><button onClick={() => setShowArrival(false)}>Отмена</button><button className="primary flex-1" onClick={doArrival} disabled={saving}>{saving ? 'Добавляем на склад...' : 'Подтвердить прибытие'}</button></div>
          </div>
        )}
        {shipment.status === 'arrived' && !showSale && <button className="primary" onClick={() => setShowSale(true)} style={{ width: '100%', justifyContent: 'center' }}><IconMoney size={18} /> Указать продажную цену</button>}
        {showSale && (
          <div className="card">
            <div style={{ fontWeight: 600, marginBottom: 12 }}>Продажа</div>
            <Field label="Продажная цена, ₽"><input type="number" value={sale.sale_price_rub} onChange={e => setSale(f => ({ ...f, sale_price_rub: e.target.value }))} placeholder="0" /></Field>
           
            {sale.sale_price_rub > 0 && (() => {
              const revenue = +sale.sale_price_rub;
              const tax = sale.usn ? Math.round(revenue * 0.06) : 0;
              const netRevenue = revenue - tax;
              const profit = netRevenue - cost;
              const perPerson = Math.round(profit / 2);
              return (
                <div className="card" style={{ background: '#f0f8ff', marginBottom: 12, fontSize: 13 }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                    <tbody>
                      <tr><td className="muted">Выручка</td><td style={{ textAlign: 'right' }}>{fmt(revenue)} ₽</td></tr>
                      {<tr><td className="muted">УСН 6%</td><td style={{ textAlign: 'right', color: '#d93636' }}>−{fmt(tax)} ₽</td></tr>}
                      <tr><td className="muted">Себестоимость</td><td style={{ textAlign: 'right', color: '#d93636' }}>−{fmt(cost)} ₽</td></tr>
                      <tr style={{ borderTop: '1px solid #ddd' }}><td style={{ fontWeight: 600, paddingTop: 6 }}>Чистая прибыль</td><td style={{ textAlign: 'right', fontWeight: 700, color: profit >= 0 ? '#15803d' : '#d93636', paddingTop: 6 }}>{profit >= 0 ? '+' : ''}{fmt(profit)} ₽</td></tr>
                      <tr><td className="muted">Каждому 👤</td><td style={{ textAlign: 'right', color: '#0077B6', fontWeight: 600 }}>{perPerson >= 0 ? '+' : ''}{fmt(perPerson)} ₽</td></tr>
                    </tbody>
                  </table>
                </div>
              );
            })()}
            <div className="row"><button onClick={() => setShowSale(false)}>Отмена</button><button className="primary flex-1" onClick={doSale} disabled={saving}>Сохранить</button></div>
          </div>
        )}
        <button className="danger" onClick={doDelete} style={{ marginTop: 4 }}><IconTrash size={16} /> Удалить поставку</button>
      </div>
    </div>
  );
}

// ─── WAREHOUSE ────────────────────────────────────────────────────────────────

function WarehouseItemDetail({ item, onUpdate, onDelete, onBack }) {
  const [showSale, setShowSale] = useState(false);
  const [showReceive, setShowReceive] = useState(false);
  const [saleItems, setSaleItems] = useState([]);
  const [saleNote, setSaleNote] = useState('');
  const [saleQty, setSaleQty] = useState('');
  const [recQty, setRecQty] = useState('');
  const [recNote, setRecNote] = useState('');
  const [parsing, setParsing] = useState(false);
  const [saving, setSaving] = useState(false);
  const [editingSell, setEditingSell] = useState(false);
  const [sellPrice, setSellPrice] = useState(item.sell_price || '');
  const photoRef = useRef();

  async function handleSaleInvoice(file) {
    setParsing(true);
    try {
      const data = await parseInvoiceFile(file);
      setSaleItems(data.items || []);
      const keyword = item.name.toLowerCase().split(' ')[0];
      const match = data.items?.find(i => i.name && i.name.toLowerCase().includes(keyword));
      if (match?.qty) setSaleQty(String(match.qty));
      setSaleNote('Продажа по накладной');
    } catch {}
    setParsing(false);
  }

  const [saleAmount, setSaleAmount] = useState('');

  async function handleSale() {
    const qty = +saleQty;
    if (!qty || qty > item.qty || qty <= 0) return;
    setSaving(true);
    const costTotal = (item.cost_rub || 0) * qty;
    const saleTotal = +saleAmount || 0;
    const move = { date: new Date().toISOString(), type: 'out', qty, note: saleNote || 'Продажа', sale_amount: saleTotal, cost_amount: costTotal };
    const updated = { ...item, qty: item.qty - qty, moves: [...(item.moves || []), move] };
    await apiFetch('/api/warehouse', { method: 'PUT', body: JSON.stringify(updated) });
    onUpdate(updated); setShowSale(false); setSaleQty(''); setSaleNote(''); setSaleAmount(''); setSaleItems([]); setSaving(false);

// Авто-закрытие поставки если все товары проданы
if (updated.shipment_id) {
  try {
    const allItems = await apiFetch('/api/warehouse');
    const siblings = allItems.filter(i => i.shipment_id === updated.shipment_id);
    const allSold = siblings.length > 0 && siblings.every(i => (i.id === updated.id ? updated.qty : i.qty) === 0);
    if (allSold) {
      const totalRevenue = siblings.reduce((sum, i) => {
        const moves = (i.id === updated.id ? updated.moves : i.moves) || [];
        return sum + moves.filter(m => m.type === 'out').reduce((s, m) => s + (m.sale_amount || 0), 0);
      }, 0);
      const tax = Math.round(totalRevenue * 0.06);
      const shipments = await apiFetch('/api/shipments');
      const shipment = shipments.find(s => s.id === updated.shipment_id);
      if (shipment && shipment.status !== 'sold') {
        const updatedShipment = { ...shipment, status: 'sold', sale_price_rub: totalRevenue, usn_tax: tax, sold_date: new Date().toISOString().slice(0, 10) };
        await apiFetch('/api/shipments', { method: 'PUT', body: JSON.stringify(updatedShipment) });
      }
    }
  } catch(e) { console.error('auto-close error', e); }
}
  }

  async function handleReceive() {
    if (!recQty) return;
    setSaving(true);
    const updated = { ...item, qty: item.qty + +recQty, moves: [...(item.moves || []), { date: new Date().toISOString(), type: 'in', qty: +recQty, note: recNote || 'Поступление' }] };
    await apiFetch('/api/warehouse', { method: 'PUT', body: JSON.stringify(updated) });
    onUpdate(updated); setShowReceive(false); setRecQty(''); setRecNote(''); setSaving(false);
  }

  async function handlePhoto(file) {
    const compressed = await compressPhoto(file);
    const updated = { ...item, photo: compressed };
    await apiFetch('/api/warehouse', { method: 'PUT', body: JSON.stringify(updated) });
    onUpdate(updated);
  }

  async function saveSellPrice() {
    const updated = { ...item, sell_price: +sellPrice || 0 };
    await apiFetch('/api/warehouse', { method: 'PUT', body: JSON.stringify(updated) });
    onUpdate(updated); setEditingSell(false);
  }

  const [editing, setEditing] = useState(false);
  const [editForm, setEditForm] = useState({ name: item.name, dimensions: item.dimensions || '', unit: item.unit || 'шт', min_qty: item.min_qty || 5 });
  const setE = (k, v) => setEditForm(f => ({ ...f, [k]: v }));

  async function saveEdit() {
    const updated = { ...item, ...editForm };
    await apiFetch('/api/warehouse', { method: 'PUT', body: JSON.stringify(updated) });
    onUpdate(updated); setEditing(false);
  }

  async function doDelete() {
    if (!confirm('Удалить товар?')) return;
    await apiFetch('/api/warehouse', { method: 'DELETE', body: JSON.stringify({ id: item.id }) });
    onDelete();
  }

  const qtyColor = item.qty === 0 ? '#d93636' : item.qty <= (item.min_qty || 5) ? '#b45309' : '#15803d';

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: '1.5rem' }}>
        <button onClick={onBack} style={{ padding: '6px 10px' }}>← Назад</button>
        <div style={{ flex: 1 }}><div style={{ fontSize: 18, fontWeight: 600 }}>{item.name}</div>{item.dimensions && <div className="muted" style={{ fontSize: 12, marginTop: 2 }}>{item.dimensions}</div>}</div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <div style={{ fontSize: 22, fontWeight: 700, color: qtyColor }}>{item.qty} {item.unit}</div>
          <button onClick={() => setEditing(v => !v)} style={{ fontSize: 12, padding: '4px 10px', color: '#0077B6', borderColor: '#cce0f0' }}>{editing ? '✕' : '✏️ Изменить'}</button>
        </div>
      </div>
      {editing && (
        <div className="card mb-1" style={{ borderColor: '#0077B6' }}>
          <div style={{ fontWeight: 600, marginBottom: 12, color: '#0077B6' }}>Редактирование товара</div>
          <Field label="Название"><input value={editForm.name} onChange={e => setE('name', e.target.value)} /></Field>
          <div className="grid-2">
            <Field label="Единица измерения">
              <select value={editForm.unit} onChange={e => setE('unit', e.target.value)}>
                <option>шт</option><option>кг</option><option>м</option><option>л</option><option>упак</option><option>пара</option><option>м²</option><option>м³</option>
              </select>
            </Field>
            <Field label="Мин. остаток (уведомление)">
              <input type="number" value={editForm.min_qty} onChange={e => setE('min_qty', +e.target.value)} placeholder="5" />
            </Field>
          </div>
          <Field label="Размеры / объём / описание">
            <input value={editForm.dimensions} onChange={e => setE('dimensions', e.target.value)} placeholder="Например: 11×8×6 см, 0.5 л, арт. MH0879" />
          </Field>
          <div className="row">
            <button onClick={() => setEditing(false)}>Отмена</button>
            <button className="primary flex-1" onClick={saveEdit}>Сохранить</button>
          </div>
        </div>
      )}

      {item.photo ? (
        <div style={{ position: 'relative', marginBottom: '1rem' }}>
          <img src={item.photo} alt={item.name} style={{ width: '100%', maxHeight: 220, objectFit: 'cover', borderRadius: 12 }} />
          <button onClick={() => photoRef.current.click()} style={{ position: 'absolute', bottom: 8, right: 8, fontSize: 12 }}>📷 Изменить</button>
        </div>
      ) : (
        <div onClick={() => photoRef.current.click()} style={{ border: '2px dashed #90c8e8', borderRadius: 12, padding: '1.2rem', textAlign: 'center', cursor: 'pointer', marginBottom: '1rem', background: '#f0f8ff' }}>
          <span style={{ color: "rgba(255,255,255,0.4)" }}><IconCamera size={24} /></span><span style={{ display: "block", marginTop: 8, fontSize: 13, color: "rgba(255,255,255,0.4)" }}>Добавить фото</span>
        </div>
      )}
      <input ref={photoRef} type="file" accept="image/*" style={{ display: 'none' }} onChange={e => e.target.files[0] && handlePhoto(e.target.files[0])} />

      <div className="grid-4 mb-1">
        <MetricCard label="На складе" value={`${item.qty} ${item.unit}`} cls={item.qty === 0 ? 'danger' : ''} />
        <MetricCard label="Себестоим./шт" value={item.cost_rub ? `${fmt(item.cost_rub)} ₽` : '—'} />
        <div className="metric-card" style={{ cursor: 'pointer' }} onClick={() => setEditingSell(true)}>
          <div className="metric-label">Цена продажи</div>
          {editingSell ? (
            <div style={{ display: 'flex', gap: 4 }}>
              <input type="number" value={sellPrice} onChange={e => setSellPrice(e.target.value)} style={{ fontSize: 14, padding: '2px 6px' }} autoFocus />
              <button className="primary" onClick={saveSellPrice} style={{ padding: '2px 8px', fontSize: 12 }}>✓</button>
            </div>
          ) : (
            <div className="metric-value">{item.sell_price ? `${fmt(item.sell_price)} ₽` : <span style={{ fontSize: 14, color: '#0077B6' }}>+ Добавить</span>}</div>
          )}
        </div>
        <MetricCard label="Стоим. остатка" value={item.cost_rub && item.qty ? `${fmt(item.cost_rub * item.qty)} ₽` : '—'} />
      </div>

      <div style={{ display: 'flex', flexDirection: 'column', gap: 8, marginBottom: '1rem' }}>
        {!showSale && !showReceive && (
          <>
            <button className="primary" onClick={() => setShowSale(true)} style={{ width: '100%', justifyContent: 'center' }}><IconArrowDown size={18} /> Списать по накладной</button>
            <button onClick={() => setShowReceive(true)} style={{ width: '100%', justifyContent: 'center' }}><IconArrowUp size={18} /> Принять дополнительно</button>
          </>
        )}

        {showSale && (
          <div className="card">
            <div style={{ fontWeight: 600, marginBottom: 10 }}>Списание по накладной</div>
            <DropZone onFile={handleSaleInvoice} loading={parsing} label="Загрузить накладную о продаже" />
            {saleItems.length > 0 && (
              <div className="info-box" style={{ marginTop: 8, fontSize: 12 }}>
                📄 Найдено {saleItems.length} позиций — количество подставлено автоматически
              </div>
            )}
            <div className="muted" style={{ fontSize: 12, textAlign: 'center', margin: '8px 0' }}>или введите вручную</div>
            <Field label={`Количество (на складе: ${item.qty} ${item.unit})`}>
              <input type="number" value={saleQty} onChange={e => setSaleQty(e.target.value)} max={item.qty} placeholder="0" />
            </Field>
            <Field label="Сумма продажи, ₽">
              <input type="number" value={saleAmount} onChange={e => setSaleAmount(e.target.value)} placeholder="0"
                onFocus={e => { if (!saleAmount && saleQty && item.sell_price) setSaleAmount(String(Math.round(+saleQty * item.sell_price))); }} />
            </Field>
            {saleAmount && saleQty && item.cost_rub ? (
              <div className={`profit-preview ${+saleAmount - +saleQty * item.cost_rub >= 0 ? 'pos' : 'neg'}`}>
                Прибыль: {(+saleAmount - +saleQty * item.cost_rub) >= 0 ? '+' : ''}{Math.round(+saleAmount - +saleQty * item.cost_rub).toLocaleString('ru-RU')} ₽
              </div>
            ) : null}
            <Field label="Комментарий"><input value={saleNote} onChange={e => setSaleNote(e.target.value)} placeholder="Продажа" /></Field>
            <div className="row"><button onClick={() => { setShowSale(false); setSaleItems([]); }}>Отмена</button><button className="primary flex-1" onClick={handleSale} disabled={saving || !saleQty || +saleQty > item.qty || +saleQty <= 0}>Списать {saleQty} {item.unit}</button></div>
          </div>
        )}

        {showReceive && (
          <div className="card">
            <div style={{ fontWeight: 600, marginBottom: 12 }}>Дополнительный приход</div>
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
            <div key={i} style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, padding: '6px 0', borderBottom: i < arr.length - 1 ? '1px solid #f0f0f0' : 'none' }}>
              <div><span style={{ color: m.type === 'in' ? '#15803d' : '#d93636', marginRight: 8 }}>{m.type === 'in' ? '↑' : '↓'}</span>{m.note || (m.type === 'in' ? 'Поступление' : 'Списание')}</div>
              <div style={{ textAlign: 'right' }}>
                <div style={{ color: m.type === 'in' ? '#15803d' : '#d93636', fontWeight: 500 }}>{m.type === 'in' ? '+' : '-'}{m.qty} {item.unit}</div>
                {m.type === 'out' && m.sale_amount > 0 && <div style={{ fontSize: 11, color: m.sale_amount - (m.cost_amount||0) >= 0 ? '#15803d' : '#d93636', fontWeight: 500 }}>{m.sale_amount - (m.cost_amount||0) >= 0 ? '+' : ''}{Math.round(m.sale_amount - (m.cost_amount||0)).toLocaleString('ru-RU')} ₽</div>}
                <div className="muted" style={{ fontSize: 11 }}>{new Date(m.date).toLocaleDateString('ru-RU')}</div>
              </div>
            </div>
          ))}
        </div>
      )}
      <button className="danger" onClick={doDelete} style={{ width: '100%', justifyContent: 'center', marginTop: 4 }}><IconTrash size={16} /> Удалить товар</button>
    </div>
  );
}

function WarehouseStats({ warehouse }) {
  const totalQty = warehouse.reduce((s, i) => s + (i.qty || 0), 0);
  const totalSell = warehouse.reduce((s, i) => s + ((i.sell_price || 0) * (i.qty || 0)), 0);
  return (
    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(110px,1fr))', gap: 10, marginBottom: '1.5rem' }}>
      <MetricCard label="Позиций" value={warehouse.length} />
      <MetricCard label="Единиц" value={fmt(totalQty)} />
      <MetricCard label="Остаток (продажи)" value={fmt(totalSell) + ' ₽'} />
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
const [filter, setFilter] = useState('all');
const [showProfit, setShowProfit] = useState(false);
const [refreshing, setRefreshing] = useState(false);

async function refreshData() {
setRefreshing(true);
try {
const [s, w] = await Promise.all([
fetch('/api/shipments').then(r => r.json()).catch(() => []),
fetch('/api/warehouse').then(r => r.json()).catch(() => []),
]);
setShipments(s || []); setWarehouse(w || []);
} catch(e) {}
setRefreshing(false);
}

useEffect(() => {
let startY = 0;
let pulling = false;
let indicator = null;

const onTouchStart = (e) => {
if (window.scrollY === 0) { startY = e.touches[0].clientY; pulling = true; }
};

const onTouchMove = (e) => {
if (!pulling) return;
const dy = e.touches[0].clientY - startY;
if (dy > 10 && window.scrollY === 0) {
if (!indicator) {
indicator = document.createElement('div');
indicator.id = 'ptr-indicator';
indicator.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:9999;display:flex;align-items:center;justify-content:center;padding:12px;background:rgba(0,119,182,0.1);font-size:13px;color:#0077B6;transition:opacity 0.2s;';
indicator.textContent = '↓ Потяните чтобы обновить';
document.body.appendChild(indicator);
}
if (dy > 70) indicator.textContent = '↑ Отпустите для обновления';
else indicator.textContent = '↓ Потяните чтобы обновить';
}
};

const onTouchEnd = (e) => {
if (!pulling) return;
const dy = (e.changedTouches[0]?.clientY || startY) - startY;
if (indicator) { indicator.remove(); indicator = null; }
if (dy > 70 && window.scrollY === 0) refreshData();
pulling = false; startY = 0;
};

document.addEventListener('touchstart', onTouchStart, { passive: true });
document.addEventListener('touchmove', onTouchMove, { passive: true });
document.addEventListener('touchend', onTouchEnd, { passive: true });
return () => {
document.removeEventListener('touchstart', onTouchStart);
document.removeEventListener('touchmove', onTouchMove);
document.removeEventListener('touchend', onTouchEnd);
};
}, []);

  useEffect(() => {
    Promise.all([
      fetch('/api/shipments').then(r => r.json()).catch(() => []),
      fetch('/api/warehouse').then(r => r.json()).catch(() => []),
    ]).then(([s, w]) => { setShipments(s || []); setWarehouse(w || []); setLoading(false); });
  }, []);

  function updateShipment(u) { setShipments(s => s.map(x => x.id === u.id ? u : x)); }
  function updateWarehouseItem(u) { setWarehouse(w => w.map(x => x.id === u.id ? u : x)); }

  if (loading) return <div className="container" style={{ color: '#777', fontSize: 14, paddingTop: '3rem' }}>Загрузка...</div>;

  if (view === 'new-shipment') return <div className="container"><NewShipment onSave={s => { setShipments(p => [...p, s]); setView('list'); }} onCancel={() => setView('list')} /></div>;

  if (view === 'detail-shipment') {
    const s = shipments.find(s => s.id === selected);
    if (!s) { setView('list'); return null; }
    return <div className="container"><ShipmentDetail shipment={s} onUpdate={updateShipment} onDelete={() => { setShipments(p => p.filter(x => x.id !== selected)); setView('list'); }} onBack={() => setView('list')} onWarehouseUpdate={added => { setWarehouse(p => [...p, ...added]); }} /></div>;
  }

  if (view === 'detail-warehouse') {
    const item = warehouse.find(i => i.id === selected);
    if (!item) { setView('list'); return null; }
    return <div className="container"><WarehouseItemDetail item={item} onUpdate={updateWarehouseItem} onDelete={() => { setWarehouse(p => p.filter(x => x.id !== selected)); setView('list'); }} onBack={() => setView('list')} /></div>;
  }

  const totalProfit = shipments.filter(s => s.status === 'sold').reduce((sum, s) => sum + ((s.sale_price_rub || 0) - (s.usn_tax || 0) - (s.paid_rub || 0) - (s.extra_paid_rub || 0) - (s.delivery_rub || 0)), 0);

  return (
    <div className="container">
      {refreshing && <div style={{ position: 'fixed', top: 0, left: 0, right: 0, zIndex: 9999, background: '#0077B6', height: 3, animation: 'none' }}><div style={{ height: '100%', background: '#48CAE4', width: '60%', borderRadius: 2 }} /></div>}
<div className="tab-bar">
<button onClick={() => setTab('shipments')} className={tab === 'shipments' ? 'active' : ''}><IconShip size={16} /> Поставки</button>
<button onClick={() => setTab('warehouse')} className={tab === 'warehouse' ? 'active' : ''}><IconBox size={16} /> Склад</button>
</div>

      {showProfit && (
<div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)', zIndex: 1000, display: 'flex', alignItems: 'flex-end' }} onClick={() => setShowProfit(false)}>
<div style={{ background: '#fff', borderRadius: '16px 16px 0 0', padding: '1.5rem', width: '100%', maxHeight: '80vh', overflowY: 'auto' }} onClick={e => e.stopPropagation()}>
<div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1.2rem' }}>
<div style={{ fontSize: 18, fontWeight: 700 }}>Прибыль по партиям</div>
<button onClick={() => setShowProfit(false)} style={{ border: 'none', background: 'none', fontSize: 20, cursor: 'pointer' }}>✕</button>
</div>
{shipments.filter(s => s.status === 'sold').sort((a,b) => new Date(b.sold_date||b.createdAt) - new Date(a.sold_date||a.createdAt)).map(s => {
const cost = (s.paid_rub||0) + (s.extra_paid_rub||0) + (s.delivery_rub||0);
const tax = s.usn_tax || Math.round((s.sale_price_rub||0) * 0.06);
const profit = (s.sale_price_rub||0) - tax - cost;
const perPerson = Math.round(profit / 2);
return (
<div key={s.id} style={{ background: '#f8fafc', borderRadius: 12, padding: '1rem', marginBottom: 10, border: '1px solid #e0eef8' }}>
<div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 10 }}>
<div>
<div style={{ fontWeight: 600, fontSize: 15 }}>{s.name}</div>
<div style={{ fontSize: 12, color: '#888', marginTop: 2 }}>{s.sold_date || s.arrived_date || ''}</div>
</div>
<div style={{ fontWeight: 700, fontSize: 16, color: profit >= 0 ? '#15803d' : '#d93636' }}>{profit >= 0 ? '+' : ''}{fmt(profit)} ₽</div>
</div>
<table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}><tbody>
<tr><td style={{ color: '#888', paddingBottom: 4 }}>Выручка</td><td style={{ textAlign: 'right', paddingBottom: 4 }}>{fmt(s.sale_price_rub)} ₽</td></tr>
<tr><td style={{ color: '#888', paddingBottom: 4 }}>Себестоимость</td><td style={{ textAlign: 'right', color: '#d93636', paddingBottom: 4 }}>-{fmt(cost)} ₽</td></tr>
<tr><td style={{ color: '#888', paddingBottom: 4 }}>Налог УСН 6%</td><td style={{ textAlign: 'right', color: '#d93636', paddingBottom: 4 }}>-{fmt(tax)} ₽</td></tr>
<tr style={{ borderTop: '1px solid #e0eef8' }}>
<td style={{ paddingTop: 8, fontWeight: 600 }}>Каждому 👤</td>
<td style={{ textAlign: 'right', paddingTop: 8, fontWeight: 700, color: '#0077B6', fontSize: 15 }}>{perPerson >= 0 ? '+' : ''}{fmt(perPerson)} ₽</td>
</tr>
</tbody></table>
</div>
);
})}
<div style={{ background: 'linear-gradient(135deg,#0077B6,#48CAE4)', borderRadius: 12, padding: '1rem', color: '#fff', marginTop: 8 }}>
<div style={{ fontSize: 13, opacity: 0.85, marginBottom: 6 }}>Итого по всем партиям</div>
{(() => {
const sold = shipments.filter(s => s.status === 'sold');
const totRev = sold.reduce((s,x) => s + (x.sale_price_rub||0), 0);
const totCost = sold.reduce((s,x) => s + (x.paid_rub||0) + (x.extra_paid_rub||0) + (x.delivery_rub||0), 0);
const totTax = sold.reduce((s,x) => s + (x.usn_tax || Math.round((x.sale_price_rub||0)*0.06)), 0);
const totProfit = totRev - totCost - totTax;
return (
<div>
<div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 4 }}><span style={{ opacity: 0.85 }}>Выручка</span><span style={{ fontWeight: 600 }}>{fmt(totRev)} ₽</span></div>
<div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 4 }}><span style={{ opacity: 0.85 }}>Налог УСН</span><span style={{ fontWeight: 600 }}>-{fmt(totTax)} ₽</span></div>
<div style={{ display: 'flex', justifyContent: 'space-between', borderTop: '1px solid rgba(255,255,255,0.3)', paddingTop: 8, marginTop: 4 }}>
<span style={{ fontWeight: 700, fontSize: 15 }}>Каждому 👤</span>
<span style={{ fontWeight: 800, fontSize: 18 }}>{fmt(Math.round(totProfit/2))} ₽</span>
</div>
</div>
);
})()}
</div>
</div>
</div>
)}
{tab === 'shipments' && (
        <div>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '1.5rem' }}>
            <div><div style={{ fontSize: 22, fontWeight: 700 }}>Поставки</div><div className="muted" style={{ fontSize: 13, marginTop: 2 }}>Китай → Россия</div></div>
            <div style={{ display: 'flex', gap: 8 }}>
<button onClick={refreshData} style={{ padding: '8px 12px', fontSize: 18, lineHeight: 1 }} title="Обновить"><IconRefresh size={17} color={refreshing ? "#48CAE4" : "currentColor"} /></button>
<button className="primary" onClick={() => setView('new-shipment')}>+ Новая</button>
</div>
          </div>
          {shipments.length > 0 && (
            <div className="grid-4" style={{ marginBottom: '1.5rem' }}>
              <MetricCard label="Всего" value={shipments.length} />
              <MetricCard label="В пути" value={shipments.filter(s => s.status === 'transit').length} />
              <MetricCard label="Прибыло" value={shipments.filter(s => s.status === 'arrived').length} />
              <div className="metric-card" style={{ cursor: shipments.some(s => s.status === 'sold') ? 'pointer' : 'default' }} onClick={() => shipments.some(s => s.status === 'sold') && setShowProfit(true)}>
<div className="metric-label">Прибыль {shipments.some(s => s.status === 'sold') && <span style={{fontSize:10}}>↗️</span>}</div>
<div className={`metric-value ${shipments.some(s => s.status === 'sold') ? (totalProfit >= 0 ? 'success' : 'danger') : ''}`}>{totalProfit >= 0 ? '+' : ''}{fmt(totalProfit)} ₽</div>
</div>
            </div>
          )}
          {shipments.length > 0 && (
<div style={{ display: 'flex', gap: 6, marginBottom: '1rem', flexWrap: 'wrap' }}>
{[['all', 'Все'], ['new', 'Новые'], ['transit', 'В пути'], ['arrived', 'Прибыли'], ['sold', 'Проданы']].map(([key, label]) => (
<button key={key} onClick={() => setFilter(key)} style={{ padding: '5px 12px', fontSize: 13, borderRadius: 20, border: '1.5px solid', borderColor: filter === key ? '#0077B6' : '#d1d0c9', background: filter === key ? '#0077B6' : '#fff', color: filter === key ? '#fff' : '#555', fontWeight: filter === key ? 600 : 400 }}>{label}{key !== 'all' ? ' (' + shipments.filter(s => s.status === key).length + ')' : ''}</button>
))}
</div>
)}
{shipments.length === 0 ? (
            <div className="empty-state"><div style={{ fontSize: 48, marginBottom: '1rem' }}>📦</div><div style={{ fontSize: 16, fontWeight: 600, marginBottom: 6 }}>Поставок нет</div><div className="muted">Нажмите «Новая» чтобы начать</div></div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {[...shipments].filter(s => filter === 'all' || s.status === filter).reverse().map(s => {
                const cost = (s.paid_rub || 0) + (s.extra_paid_rub || 0) + (s.delivery_rub || 0);
                const profit = s.status === 'sold' ? (s.sale_price_rub || 0) - (s.usn_tax || 0) - cost : null;
                const daysLeft = s.eta_date && s.status === 'transit' ? Math.ceil((new Date(s.eta_date) - new Date()) / 86400000) : null;
            const isOverdue = s.status === 'transit' && daysLeft !== null && daysLeft <= 0;
                return (
                  <div key={s.id} className="card shipment-row" onClick={() => { setSelected(s.id); setView('detail-shipment'); }}
style={isOverdue ? { borderColor: '#f87171', background: '#fff5f5' } : {}}>
                    <div style={{ width: 44, height: 44, borderRadius: 12, background: "rgba(0,119,182,0.15)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}><IconBox size={22} color="#48CAE4" /></div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 3 }}>{s.name}</div>
                      <div className="muted" style={{ fontSize: 12 }}>{s.total_cny ? `¥ ${fmt(s.total_cny)}` : ''}{s.items?.length ? ` · ${s.items.length} поз.` : ''}{daysLeft !== null ? (isOverdue ? <span style={{ color: '#d93636', fontWeight: 600 }}> · ⚠️ Просрочено на {Math.abs(daysLeft)} дн.</span> : ` · ${daysLeft} дн. до прибытия`) : ''}{s.ship_date && !daysLeft ? ` · ${s.ship_date}` : ''}</div>
                    </div>
                    <div style={{ textAlign: 'right', flexShrink: 0 }}>
                      <Badge status={s.status} />
                      {profit !== null && <div className={profit >= 0 ? 'success' : 'danger'} style={{ fontSize: 13, marginTop: 4, fontWeight: 600 }}>{profit >= 0 ? '+' : ''}{fmt(profit)} ₽</div>}
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}

      {tab === 'warehouse' && (
        <div>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '1.5rem' }}>
            <div><div style={{ fontSize: 22, fontWeight: 700 }}>Склад</div><div className="muted" style={{ fontSize: 13, marginTop: 2 }}>{warehouse.length} позиций</div></div>
          </div>
          {warehouse.length > 0 && <WarehouseStats warehouse={warehouse} />}
          {warehouse.length === 0 ? (
            <div className="empty-state"><div style={{ fontSize: 48, marginBottom: '1rem' }}>🏪</div><div style={{ fontSize: 16, fontWeight: 600, marginBottom: 6 }}>Склад пуст</div><div className="muted">Когда груз прибудет — товары добавятся автоматически</div></div>
          ) : (
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(155px, 1fr))', gap: 12 }}>
              {warehouse.map(item => {
                const qtyColor = item.qty === 0 ? '#d93636' : item.qty <= (item.min_qty || 5) ? '#b45309' : '#15803d';
                return (
                  <div key={item.id} onClick={() => { setSelected(item.id); setView('detail-warehouse'); }}
                    className="warehouse-card">
                    {item.photo ? <img src={item.photo} alt={item.name} style={{ width: '100%', height: 130, objectFit: 'cover' }} /> : <div className="warehouse-card-img"><IconBox size={36} color="rgba(72,202,228,0.6)" /></div>}
                    <div style={{ padding: '10px 12px' }}>
                      <div style={{ fontWeight: 600, fontSize: 13, marginBottom: 4, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{item.name}</div>
                      <div style={{ fontSize: 14, color: qtyColor, fontWeight: 600 }}>{item.qty} {item.unit}</div>
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
