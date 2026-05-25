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
      {sub && <div style={{ fontSize: 11, color: '#5f5e5a', marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

function DropZone({ onFile, loading }) {
  const [drag, setDrag] = useState(false);
  const ref = useRef();
  return (
    <div
      className={`drop-zone${drag ? ' drag' : ''}`}
      onDragOver={e => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]); }}
      onClick={() => !loading && ref.current.click()}
    >
      <input ref={ref} type="file" accept=".pdf,.png,.jpg,.jpeg,.xlsx,.xls,.csv"
        style={{ display: 'none' }} onChange={e => e.target.files[0] && onFile(e.target.files[0])} />
      <div style={{ fontSize: 32, marginBottom: 8 }}>📄</div>
      {loading
        ? <div className="muted">Распознаём накладную...</div>
        : <>
          <div style={{ fontWeight: 500, marginBottom: 4 }}>Загрузите накладную</div>
          <div className="muted" style={{ fontSize: 12 }}>PDF · Фото (JPG, PNG) · Excel — перетащите или нажмите</div>
        </>
      }
    </div>
  );
}

function NewShipment({ onSave, onCancel }) {
  const [step, setStep] = useState(1);
  const [parsing, setParsing] = useState(false);
  const [parseError, setParseError] = useState('');
  const [fileName, setFileName] = useState('');
  const [saving, setSaving] = useState(false);
  const [form, setForm] = useState({
    name: '', items: [], total_cny: '', invoice_number: '', supplier: '',
    cny_rate: '', paid_rub: '', ship_date: '', eta_date: '',
  });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));

  async function handleFile(file) {
    setFileName(file.name); setParsing(true); setParseError('');
    try {
      const d = await parseInvoiceFile(file);
      setForm(f => ({
        ...f,
        items: d.items || [],
        total_cny: d.total_cny || '',
        invoice_number: d.invoice_number || '',
        supplier: d.supplier || '',
        name: d.supplier ? `Поставка — ${d.supplier}` : `Поставка ${new Date().toLocaleDateString('ru-RU')}`,
      }));
      setStep(2);
    } catch { setParseError('Не удалось распознать файл. Попробуйте другой формат или введите вручную.'); }
    setParsing(false);
  }

  async function handleSave() {
    setSaving(true);
    const totalCNY = +form.total_cny || form.items.reduce((s, i) => s + (i.total_cny || 0), 0);
    const shipment = {
      id: genId(), createdAt: new Date().toISOString(),
      status: form.ship_date ? 'transit' : 'new',
      name: form.name || `Поставка ${new Date().toLocaleDateString('ru-RU')}`,
      items: form.items, total_cny: totalCNY,
      invoice_number: form.invoice_number, supplier: form.supplier,
      cny_rate: +form.cny_rate || 0, paid_rub: +form.paid_rub || 0,
      ship_date: form.ship_date, eta_date: form.eta_date,
      extra_paid_rub: 0, delivery_rub: 0, sale_price_rub: 0,
    };
    await apiFetch('/api/shipments', { method: 'POST', body: JSON.stringify(shipment) });
    onSave(shipment);
    setSaving(false);
  }

  const totalCNY = +form.total_cny || form.items.reduce((s, i) => s + (i.total_cny || 0), 0);
  const estRUB = totalCNY && form.cny_rate ? totalCNY * form.cny_rate : null;
  const steps = ['Накладная', 'Оплата', 'Доставка', 'Итог'];

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: '1.5rem' }}>
        <button onClick={onCancel} style={{ padding: '6px 10px' }}>← Назад</button>
        <div style={{ fontSize: 18, fontWeight: 500 }}>Новая поставка</div>
      </div>

      <div className="step-bar">
        {steps.map((s, i) => (
          <div key={i} style={{ flex: 1 }}>
            <div className="step-seg" style={{ background: i < step ? '#1c1c1a' : '#d1d0c9' }} />
            {i === step - 1 && <div style={{ fontSize: 11, color: '#5f5e5a', textAlign: 'center', marginTop: 3 }}>{s}</div>}
          </div>
        ))}
      </div>

      {step === 1 && (
        <div>
          <DropZone onFile={handleFile} loading={parsing} />
          {parseError && <div className="error-box">{parseError}</div>}
          <div className="mt-1" style={{ textAlign: 'center' }}>
            <button style={{ border: 'none', background: 'none', color: '#5f5e5a', cursor: 'pointer', fontSize: 13 }}
              onClick={() => { set('name', `Поставка ${new Date().toLocaleDateString('ru-RU')}`); setStep(2); }}>
              Пропустить, ввести вручную →
            </button>
          </div>
        </div>
      )}

      {step === 2 && (
        <div>
          {fileName && <div className="info-box">📎 {fileName} — {form.items.length > 0 ? `${form.items.length} позиций` : 'данные введены вручную'}</div>}
          <Field label="Название поставки">
            <input value={form.name} onChange={e => set('name', e.target.value)} placeholder="Электроника январь" />
          </Field>
          <Field label="Итого по накладной, ¥ CNY">
            <input type="number" value={form.total_cny} onChange={e => set('total_cny', e.target.value)} placeholder="0" />
          </Field>
          <div className="grid-2">
            <Field label="Курс покупки CNY/RUB">
              <input type="number" value={form.cny_rate} onChange={e => set('cny_rate', e.target.value)} placeholder="13.5" />
            </Field>
            <Field label="Оплачено в рублях">
              <input type="number" value={form.paid_rub} onChange={e => set('paid_rub', e.target.value)} placeholder="0" />
            </Field>
          </div>
          {estRUB && <div className="muted mb-1" style={{ fontSize: 13 }}>Расчётная сумма: ≈ {fmt(estRUB)} ₽</div>}
          <div className="row">
            <button onClick={() => setStep(1)}>← Назад</button>
            <button className="primary flex-1" onClick={() => setStep(3)}>Далее →</button>
          </div>
        </div>
      )}

      {step === 3 && (
        <div>
          <div className="grid-2">
            <Field label="Дата отправки">
              <input type="date" value={form.ship_date} onChange={e => set('ship_date', e.target.value)} />
            </Field>
            <Field label="Ожидаемое прибытие">
              <input type="date" value={form.eta_date} onChange={e => set('eta_date', e.target.value)} />
            </Field>
          </div>
          <div className="muted mb-1" style={{ fontSize: 13 }}>Можно пропустить — поставка получит статус «Новый»</div>
          <div className="row">
            <button onClick={() => setStep(2)}>← Назад</button>
            <button className="primary flex-1" onClick={() => setStep(4)}>Далее →</button>
          </div>
        </div>
      )}

      {step === 4 && (
        <div>
          <div className="card mb-1">
            <div style={{ fontWeight: 500, fontSize: 15, marginBottom: 12 }}>{form.name}</div>
            <table className="items-table">
              <tbody>
                {totalCNY > 0 && <Row label="Сумма закупки" value={`¥ ${fmt(totalCNY)}`} />}
                {form.cny_rate && <Row label="Курс CNY" value={`${form.cny_rate} ₽/¥`} />}
                {form.paid_rub && <Row label="Оплачено" value={`${fmt(+form.paid_rub)} ₽`} />}
                {form.ship_date && <Row label="Дата отправки" value={form.ship_date} />}
                {form.eta_date && <Row label="Ожидается" value={form.eta_date} />}
              </tbody>
            </table>
            {form.items.length > 0 && (
              <>
                <hr className="divider" />
                <div className="muted" style={{ fontSize: 12, marginBottom: 6 }}>Позиции ({form.items.length}):</div>
                {form.items.slice(0, 5).map((item, i) => (
                  <div key={i} style={{ fontSize: 12, display: 'flex', justifyContent: 'space-between', padding: '2px 0' }}>
                    <span>{item.name}</span>
                    <span className="muted">×{item.qty} · ¥{fmt(item.total_cny)}</span>
                  </div>
                ))}
                {form.items.length > 5 && <div className="muted" style={{ fontSize: 12 }}>...и ещё {form.items.length - 5}</div>}
              </>
            )}
          </div>
          <div className="row">
            <button onClick={() => setStep(3)}>← Назад</button>
            <button className="primary flex-1" onClick={handleSave} disabled={saving}>
              {saving ? 'Сохранение...' : 'Сохранить поставку'}
            </button>
          </div>
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

  async function save(updated) {
    setSaving(true);
    await apiFetch('/api/shipments', { method: 'PUT', body: JSON.stringify(updated) });
    onUpdate(updated);
    setSaving(false);
  }

  async function doArrival() {
    await save({ ...shipment, status: 'arrived', extra_paid_rub: +arrival.extra_paid_rub || 0, delivery_rub: +arrival.delivery_rub || 0, arrived_date: arrival.arrived_date });
    setShowArrival(false);
  }
  async function doSale() {
    await save({ ...shipment, status: 'sold', sale_price_rub: +sale.sale_price_rub || 0, sold_date: new Date().toISOString().slice(0, 10) });
    setShowSale(false);
  }
  async function doDelete() {
    if (!confirm('Удалить поставку?')) return;
    await apiFetch('/api/shipments', { method: 'DELETE', body: JSON.stringify({ id: shipment.id }) });
    onDelete();
  }

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: '1.5rem' }}>
        <button onClick={onBack} style={{ padding: '6px 10px' }}>← Назад</button>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 18, fontWeight: 500 }}>{shipment.name}</div>
          <div className="muted" style={{ fontSize: 12, marginTop: 2 }}>Создан {new Date(shipment.createdAt).toLocaleDateString('ru-RU')}</div>
        </div>
        <Badge status={shipment.status} />
      </div>

      <div className="grid-4 mb-1">
        <MetricCard label="Закупка (CNY)" value={shipment.total_cny ? `¥ ${fmt(shipment.total_cny)}` : '—'} />
        <MetricCard label="Оплачено" value={`${fmt(shipment.paid_rub)} ₽`} />
        <MetricCard label="Доставка" value={shipment.delivery_rub ? `${fmt(shipment.delivery_rub)} ₽` : '—'} />
        {profit !== null
          ? <MetricCard label="Прибыль" value={`${profit >= 0 ? '+' : ''}${fmt(profit)} ₽`} cls={profit >= 0 ? 'success' : 'danger'} sub={margin ? `${margin}% маржа` : null} />
          : <MetricCard label="Себестоимость" value={`${fmt(cost)} ₽`} />
        }
      </div>

      <div className="card mb-1">
        <div className="muted" style={{ fontSize: 12, fontWeight: 500, marginBottom: 8 }}>ДЕТАЛИ</div>
        <table className="items-table">
          <tbody>
            {shipment.invoice_number && <Row label="№ накладной" value={shipment.invoice_number} />}
            {shipment.supplier && <Row label="Поставщик" value={shipment.supplier} />}
            {shipment.cny_rate > 0 && <Row label="Курс CNY" value={`${shipment.cny_rate} ₽/¥`} />}
            {shipment.ship_date && <Row label="Дата отправки" value={shipment.ship_date} />}
            {shipment.eta_date && <Row label="Ожидается" value={shipment.eta_date} />}
            {shipment.arrived_date && <Row label="Дата прибытия" value={shipment.arrived_date} />}
            {shipment.extra_paid_rub > 0 && <Row label="Доп. оплата (РФ)" value={`${fmt(shipment.extra_paid_rub)} ₽`} />}
            {shipment.sale_price_rub > 0 && <Row label="Продано за" value={`${fmt(shipment.sale_price_rub)} ₽`} />}
            {shipment.sold_date && <Row label="Дата продажи" value={shipment.sold_date} />}
          </tbody>
        </table>
      </div>

      {shipment.items?.length > 0 && (
        <div className="card mb-1">
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: showItems ? 10 : 0 }}>
            <div className="muted" style={{ fontSize: 12, fontWeight: 500 }}>ПОЗИЦИИ ({shipment.items.length})</div>
            <button style={{ border: 'none', background: 'none', fontSize: 12, color: '#5f5e5a', cursor: 'pointer' }} onClick={() => setShowItems(v => !v)}>
              {showItems ? 'Скрыть' : 'Показать'}
            </button>
          </div>
          {showItems && shipment.items.map((item, i) => (
            <div key={i} style={{ display: 'flex', justifyContent: 'space-between', fontSize: 13, padding: '7px 0', borderBottom: i < shipment.items.length - 1 ? '1px solid #e8e6df' : 'none' }}>
              <div>
                <div>{item.name}</div>
                <div className="muted" style={{ fontSize: 11 }}>×{item.qty} · ¥{fmtF(item.unit_price_cny)}/шт</div>
              </div>
              <div style={{ textAlign: 'right', flexShrink: 0 }}>
                <div>¥ {fmt(item.total_cny)}</div>
                {shipment.cny_rate > 0 && <div className="muted" style={{ fontSize: 11 }}>≈ {fmt(item.total_cny * shipment.cny_rate)} ₽</div>}
              </div>
            </div>
          ))}
        </div>
      )}

      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {shipment.status === 'new' && (
          <button className="primary" onClick={() => save({ ...shipment, status: 'transit' })} disabled={saving} style={{ width: '100%', justifyContent: 'center' }}>
            🚛 Отметить как «В пути»
          </button>
        )}

        {shipment.status === 'transit' && !showArrival && (
          <button className="primary" onClick={() => setShowArrival(true)} style={{ width: '100%', justifyContent: 'center' }}>
            📦 Груз прибыл
          </button>
        )}

        {showArrival && (
          <div className="card">
            <div style={{ fontWeight: 500, marginBottom: 12 }}>Оформление прибытия</div>
            <Field label="Дата прибытия">
              <input type="date" value={arrival.arrived_date} onChange={e => setArrival(f => ({ ...f, arrived_date: e.target.value }))} />
            </Field>
            <Field label="Дополнительно оплачено в России, ₽">
              <input type="number" value={arrival.extra_paid_rub} onChange={e => setArrival(f => ({ ...f, extra_paid_rub: e.target.value }))} placeholder="0" />
            </Field>
            <Field label="Стоимость доставки, ₽">
              <input type="number" value={arrival.delivery_rub} onChange={e => setArrival(f => ({ ...f, delivery_rub: e.target.value }))} placeholder="0" />
            </Field>
            <div className="row">
              <button onClick={() => setShowArrival(false)}>Отмена</button>
              <button className="primary flex-1" onClick={doArrival} disabled={saving}>Сохранить</button>
            </div>
          </div>
        )}

        {shipment.status === 'arrived' && !showSale && (
          <button className="primary" onClick={() => setShowSale(true)} style={{ width: '100%', justifyContent: 'center' }}>
            💰 Указать продажную цену
          </button>
        )}

        {showSale && (
          <div className="card">
            <div style={{ fontWeight: 500, marginBottom: 12 }}>Продажа</div>
            <Field label="Продажная цена, ₽">
              <input type="number" value={sale.sale_price_rub} onChange={e => setSale(f => ({ ...f, sale_price_rub: e.target.value }))} placeholder="0" />
            </Field>
            {previewProfit !== null && (
              <div className={`profit-preview ${previewProfit >= 0 ? 'pos' : 'neg'}`}>
                Прибыль: {previewProfit >= 0 ? '+' : ''}{fmt(previewProfit)} ₽{cost > 0 ? ` (${((previewProfit / cost) * 100).toFixed(1)}%)` : ''}
              </div>
            )}
            <div className="row">
              <button onClick={() => setShowSale(false)}>Отмена</button>
              <button className="primary flex-1" onClick={doSale} disabled={saving}>Сохранить</button>
            </div>
          </div>
        )}

        <button className="danger" onClick={doDelete} style={{ marginTop: 4 }}>
          🗑 Удалить поставку
        </button>
      </div>
    </div>
  );
}

export default function Home() {
  const [view, setView] = useState('list');
  const [shipments, setShipments] = useState([]);
  const [selected, setSelected] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetch('/api/shipments').then(r => r.json()).then(s => { setShipments(s); setLoading(false); });
  }, []);

  function updateLocal(updated) { setShipments(s => s.map(x => x.id === updated.id ? updated : x)); }

  if (loading) return <div className="container" style={{ color: '#5f5e5a', fontSize: 14, paddingTop: '3rem' }}>Загрузка...</div>;

  if (view === 'new') return (
    <div className="container">
      <NewShipment onSave={s => { setShipments(p => [...p, s]); setView('list'); }} onCancel={() => setView('list')} />
    </div>
  );

  if (view === 'detail') {
    const s = shipments.find(s => s.id === selected);
    if (!s) { setView('list'); return null; }
    return (
      <div className="container">
        <ShipmentDetail
          shipment={s}
          onUpdate={u => updateLocal(u)}
          onDelete={() => { setShipments(p => p.filter(x => x.id !== selected)); setView('list'); }}
          onBack={() => setView('list')}
        />
      </div>
    );
  }

  const totalProfit = shipments.filter(s => s.status === 'sold')
    .reduce((sum, s) => sum + ((s.sale_price_rub || 0) - (s.paid_rub || 0) - (s.extra_paid_rub || 0) - (s.delivery_rub || 0)), 0);

  return (
    <div className="container">
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '1.5rem' }}>
        <div>
          <div style={{ fontSize: 22, fontWeight: 500 }}>Поставки</div>
          <div className="muted" style={{ fontSize: 13, marginTop: 2 }}>Китай → Россия</div>
        </div>
        <button className="primary" onClick={() => setView('new')}>+ Новая поставка</button>
      </div>

      {shipments.length > 0 && (
        <div className="grid-4" style={{ marginBottom: '1.5rem' }}>
          <MetricCard label="Поставок" value={shipments.length} />
          <MetricCard label="В пути" value={shipments.filter(s => s.status === 'transit').length} />
          <MetricCard label="Прибыло" value={shipments.filter(s => s.status === 'arrived').length} />
          <MetricCard label="Прибыль" value={`${totalProfit >= 0 ? '+' : ''}${fmt(totalProfit)} ₽`}
            cls={shipments.some(s => s.status === 'sold') ? (totalProfit >= 0 ? 'success' : 'danger') : ''} />
        </div>
      )}

      {shipments.length === 0 ? (
        <div className="empty-state">
          <div style={{ fontSize: 40, marginBottom: '1rem' }}>📦</div>
          <div style={{ fontSize: 15, fontWeight: 500, marginBottom: 6 }}>Поставок пока нет</div>
          <div className="muted" style={{ fontSize: 13 }}>Нажмите «Новая поставка» чтобы начать</div>
        </div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          {[...shipments].reverse().map(s => {
            const cost = (s.paid_rub || 0) + (s.extra_paid_rub || 0) + (s.delivery_rub || 0);
            const profit = s.status === 'sold' ? (s.sale_price_rub || 0) - cost : null;
            const daysLeft = s.eta_date && s.status === 'transit'
              ? Math.ceil((new Date(s.eta_date) - new Date()) / 86400000) : null;
            return (
              <div key={s.id} className="card shipment-row" onClick={() => { setSelected(s.id); setView('detail'); }}>
                <div style={{ fontSize: 24 }}>📦</div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontWeight: 500, fontSize: 15, marginBottom: 3 }}>{s.name}</div>
                  <div className="muted" style={{ fontSize: 12 }}>
                    {s.total_cny ? `¥ ${fmt(s.total_cny)}` : ''}
                    {s.items?.length ? ` · ${s.items.length} поз.` : ''}
                    {daysLeft !== null ? ` · ${daysLeft > 0 ? `${daysLeft} дн. до прибытия` : 'прибытие сегодня/просрочено'}` : ''}
                    {s.ship_date && !daysLeft ? ` · Отправлено: ${s.ship_date}` : ''}
                  </div>
                </div>
                <div style={{ textAlign: 'right', flexShrink: 0 }}>
                  <Badge status={s.status} />
                  {profit !== null && (
                    <div className={profit >= 0 ? 'success' : 'danger'} style={{ fontSize: 13, marginTop: 4, fontWeight: 500 }}>
                      {profit >= 0 ? '+' : ''}{fmt(profit)} ₽
                    </div>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
