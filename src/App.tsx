import { useState } from 'react';
import { saveAs } from 'file-saver';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, WidthType, TextRun, BorderStyle, PageOrientation, ExternalHyperlink } from 'docx';

const cyrillicToLatinUz = (text: string): string => {
  if (!text) return '';

  const map: { [key: string]: string } = {
    'А': 'A', 'а': 'a',
    'Б': 'B', 'б': 'b',
    'В': 'V', 'в': 'v',
    'Г': 'G', 'г': 'g',
    'Д': 'D', 'д': 'd',
    'Е': 'E', 'е': 'e',
    'Ё': 'Yo', 'ё': 'yo',
    'Ж': 'J', 'ж': 'j',
    'З': 'Z', 'з': 'z',
    'И': 'I', 'и': 'i',
    'Й': 'Y', 'й': 'y',
    'К': 'K', 'к': 'k',
    'Л': 'L', 'л': 'l',
    'М': 'M', 'м': 'm',
    'Н': 'N', 'н': 'n',
    'О': 'O', 'о': 'o',
    'П': 'P', 'п': 'p',
    'Р': 'R', 'р': 'r',
    'С': 'S', 'с': 's',
    'Т': 'T', 'т': 't',
    'У': 'U', 'у': 'u',
    'Ф': 'F', 'ф': 'f',
    'Х': 'X', 'х': 'x',
    'Ц': 'Ts', 'ц': 'ts',
    'Ч': 'Ç', 'ч': 'ç',
    'Ш': 'Ş', 'ш': 'ş',
    'Щ': 'Ş', 'щ': 'ş',
    'Ъ': "'", 'ъ': "'",
    'Ь': '', 'ь': '',
    'Э': 'E', 'э': 'e',
    'Ю': 'Yu', 'ю': 'yu',
    'Я': 'Ya', 'я': 'ya',
    'Ў': 'Ö', 'ў': 'ö',
    'Қ': 'Q', 'қ': 'q',
    'Ғ': 'Ğ', 'ғ': 'ğ',
    'Ҳ': 'H', 'ҳ': 'h',
  };

  let transliterated = text.split('').map((char, i, arr) => {
    const prev = arr[i-1];
    const isVowel = prev && 'АаЕеЁёИиОоУуЭэЮюЯяЎўÖö'.includes(prev);
    const isStart = !prev || /[\s\p{P}]/u.test(prev);
    
    // 'Е' logic
    if ((char === 'Е' || char === 'е') && (isStart || isVowel)) {
      return char === 'Е' ? 'Ye' : 'ye';
    }
    
    // 'Ц' logic: at start of word becomes 'S', else 'Ts'
    if ((char === 'Ц' || char === 'ц') && isStart) {
      return char === 'Ц' ? 'S' : 's';
    }

    return map[char] !== undefined ? map[char] : char;
  }).join('');

  // Handle any pre-existing or residual old Latin digraphs / apostrophe-based letters:
  // «Sh» -> «Ş», «Ch» -> «Ç», «Oʻ» -> «Ö», «Gʻ» -> «Ğ»
  transliterated = transliterated
    .replace(/[OО][ʻ‘’’'`ʼ]/g, 'Ö')
    .replace(/[oо][ʻ‘’’'`ʼ]/g, 'ö')
    .replace(/[GГ][ʻ‘’’'`ʼ]/g, 'Ğ')
    .replace(/[gг][ʻ‘’’'`ʼ]/g, 'ğ')
    .replace(/SH/g, 'Ş')
    .replace(/Sh/g, 'Ş')
    .replace(/sh/g, 'ş')
    .replace(/CH/g, 'Ç')
    .replace(/Ch/g, 'Ç')
    .replace(/ch/g, 'ç');

  return transliterated;
};

type Busy = null | 'process' | 'headlines' | 'translate' | 'sync' | 'word';

const lines = (t: string): string[] =>
  t.split('\n').map(s => s.trim()).filter(s => s.length > 0 && !/^(источник|manba|манба):/i.test(s));

async function api<T = any>(path: string, body: object): Promise<T> {
  const res = await fetch(path, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  let data: any = {};
  try { data = await res.json(); } catch { throw new Error('Сервердан нотўғри жавоб келди ёки вақт тугади. Қайта уриниб кўринг.'); }
  if (!res.ok || data.error) throw new Error(data.error || 'Хатолик юз берди.');
  return data;
}

// Word: битта жадвал, икки устун (русча | ўзбекча), қалин сарлавҳа + матн
const buildDoc = (cols: string[][], sourceUrl = ''): Document => {
  const line = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
  const borders = { top: line, bottom: line, left: line, right: line };
  const cells = cols.map(ls => {
    const [head, ...body] = ls;
    const paras: Paragraph[] = [];
    if (head) paras.push(new Paragraph({ children: [new TextRun({ text: head, bold: true, size: 28, font: 'Calibri' })], spacing: { after: 280, line: 276 } }));
    body.forEach(t => paras.push(new Paragraph({ children: [new TextRun({ text: t, size: 28, font: 'Calibri' })], spacing: { after: 200, line: 276 } })));
    if (!paras.length) paras.push(new Paragraph({ children: [] }));
    return new TableCell({ width: { size: 50, type: WidthType.PERCENTAGE }, borders, margins: { top: 120, bottom: 120, left: 140, right: 140 }, children: paras });
  });
  return new Document({
    sections: [{
      properties: { page: { size: { orientation: PageOrientation.PORTRAIT }, margin: { top: 720, right: 720, bottom: 720, left: 720 } } },
      children: [
        new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [new TableRow({ children: cells })] }),
        ...(sourceUrl && /^https?:\/\//i.test(sourceUrl) ? [
          new Paragraph({
            children: [
              new ExternalHyperlink({ link: sourceUrl, children: [new TextRun({ text: sourceUrl, style: 'Hyperlink', size: 28, font: 'Calibri' })] })
            ],
            spacing: { before: 280, after: 200 }
          })
        ] : [])
      ]
    }]
  });
};

export default function App() {
  const [mode, setMode] = useState<'url' | 'text'>('url');
  const [url, setUrl] = useState('');
  const [sourceUrl, setSourceUrl] = useState('');
  const [text, setText] = useState('');
  const [step, setStep] = useState<1 | 2 | 3>(1);
  const [busy, setBusy] = useState<Busy>(null);
  const [error, setError] = useState('');
  const [note, setNote] = useState('');

  const [ru, setRu] = useState('');                 // 2-қадам: соддалаштирилган русча матн
  const [headlines, setHeadlines] = useState<string[]>([]);
  const [headline, setHeadline] = useState('');
  const [ruFull, setRuFull] = useState('');         // 3-қадам: русча (сарлавҳа + матн)
  const [uz, setUz] = useState('');                 // 3-қадам: ўзбекча кирилл

  const run = async (kind: Exclude<Busy, null>, fn: () => Promise<void>) => {
    setBusy(kind); setError(''); setNote('');
    try { await fn(); } catch (e: any) { setError(e?.message || String(e)); }
    setBusy(null);
  };

  const process = () => run('process', async () => {
    if (mode === 'url' && !url.trim().startsWith('http') && !url.includes('t.me/')) throw new Error('Илтимос, тўғри ҳавола (https://…) киритинг.');
    if (mode === 'text' && text.trim().length < 50) throw new Error('Илтимос, мақола матнини киритинг (камида 50 белги).');
    const d = await api('/api/process-article', { inputMode: mode, url: url.trim(), inputText: text, stylePrompt: '' });
    const hs: string[] = d.headlines || [];
    setRu(d.rewrittenText || ''); setHeadlines(hs); setHeadline(hs[0] || '');
    setSourceUrl(mode === 'url' ? (/^https?:\/\//i.test(url.trim()) ? url.trim() : `https://${url.trim()}`) : '');
    setStep(2);
  });

  const moreHeadlines = () => run('headlines', async () => {
    const d = await api('/api/regenerate-headlines', { rewrittenText: ru, stylePrompt: '', currentHeadlines: headlines });
    if (Array.isArray(d.headlines) && d.headlines.length) { setHeadlines(d.headlines); setHeadline(d.headlines[0]); }
  });

  const translate = () => run('translate', async () => {
    if (!headline.trim() || !ru.trim()) throw new Error('Сарлавҳа ва матн бўш бўлмасин.');
    const full = `${headline.trim()}\n\n${ru.trim()}`;
    const d = await api('/api/translate-uzbek', { fullText: full, stylePrompt: '' });
    setRuFull(full); setUz(d.uzbekCyrillic || ''); setStep(3);
  });

  const syncRu = () => run('sync', async () => {
    const [h, ...rest] = lines(ruFull);
    const d = await api('/api/sync-russian', { uzbekCyrillic: uz, headline: h || '', russianText: rest.join('\n\n') });
    setRuFull(`${d.headline || h || ''}\n\n${d.text || rest.join('\n\n')}`);
    setNote('Русча матн ўзбекча таҳрирингиз бўйича янгиланди.');
  });

  const saveWord = () => run('word', async () => {
    const left = lines(ruFull), right = lines(cyrillicToLatinUz(uz));
    if (!left.length || !right.length) throw new Error('Аввал таржимани тайёрланг.');
    const blob = await Packer.toBlob(buildDoc([left, right], sourceUrl));
    const name = (left[0].split(/\s+/).slice(0, 4).join(' ').replace(/[^\p{L}\p{N}\s_-]/gu, '').trim() || 'Maqola') + '.docx';
    saveAs(blob, name);
    setNote(`Word файли юклаб олишга юборилди: ${name}. Android телефонида Chrome орқали юкласангиз, одатда «Загрузки / Downloads» папкасига сақланади.`);
  });

  const restart = () => { setStep(1); setUrl(''); setSourceUrl(''); setText(''); setRu(''); setRuFull(''); setUz(''); setHeadlines([]); setHeadline(''); setError(''); setNote(''); };
  const wait = (k: Busy, label: string, idle: string) => (busy === k ? '⏳ ' + label : idle);

  return (
    <div className="wrap">
      <header>
        <h1>MigrantInfo Assist</h1>
        <p>Мақолаларни мигрантлар учун содда русчага қайта ёзинг, ўзбекчага таржима қилинг ва Word ҳужжати қилиб сақланг.</p>
      </header>

      {error && <div className="err">{error}</div>}
      {note && <div className="ok">{note}</div>}

      <section className="step">
        <div className="num" style={{ background: 'var(--coral)' }}>1</div>
        <h2>Мақолани киритинг</h2>
        <div className="tabs">
          <button type="button" className={'tab' + (mode === 'url' ? ' on' : '')} onClick={() => setMode('url')}>🔗 Ҳавола</button>
          <button type="button" className={'tab' + (mode === 'text' ? ' on' : '')} onClick={() => setMode('text')}>📝 Матн</button>
        </div>
        {mode === 'url'
          ? <input type="text" value={url} onChange={e => setUrl(e.target.value)} placeholder="https://… мақола ҳаволасини қўйинг" onKeyDown={e => e.key === 'Enter' && !busy && process()} />
          : <textarea value={text} onChange={e => setText(e.target.value)} placeholder="Мақолани шу ерга нусхалаб қўйинг…" />}
        <button type="button" className="btn b1" disabled={!!busy} onClick={process}>{wait('process', 'Ишламоқда…', '✨ Соддалаштириш')}</button>
      </section>

      {step >= 2 && (
        <section className="step">
          <div className="num" style={{ background: 'var(--sun)' }}>2</div>
          <h2>Сарлавҳани танланг, русча матнни текширинг</h2>
          <div className="chips">
            {headlines.map(h => <button type="button" key={h} className={'chip' + (h === headline ? ' on' : '')} onClick={() => setHeadline(h)}>{h}</button>)}
          </div>
          <input type="text" value={headline} onChange={e => setHeadline(e.target.value)} aria-label="Сарлавҳа" />
          <textarea className="mid" value={ru} onChange={e => setRu(e.target.value)} />
          <div className="row">
            <button type="button" className="btn b2" disabled={!!busy} onClick={translate}>{wait('translate', 'Таржима қилинмоқда…', '🌐 Ўзбекчага таржима қилиш')}</button>
            <button type="button" className="btn ghost" disabled={!!busy} onClick={moreHeadlines}>{wait('headlines', 'Янгиланмоқда…', '🔄 Янги сарлавҳалар')}</button>
          </div>
        </section>
      )}

      {step === 3 && (
        <section className="step">
          <div className="num" style={{ background: 'var(--blue)', color: '#fff' }}>3</div>
          <h2>Русча ва ўзбекча (кирилл) нусхаларни текширинг</h2>
          <div className="cols">
            <div><label>Русча</label><textarea value={ruFull} onChange={e => setRuFull(e.target.value)} /></div>
            <div><label>Ўзбекча (кирилл)</label><textarea value={uz} onChange={e => setUz(e.target.value)} /></div>
          </div>
          <p className="hint">Биринчи қатор — сарлавҳа. Word’га сақлаганда ўзбекча матн янги лотин алифбосига ўтади.</p>
          <div className="row">
            <button type="button" className="btn b4" disabled={!!busy} onClick={saveWord}>{wait('word', 'Тайёрланяпти…', '💾 Word (.docx) сақлаш')}</button>
            <button type="button" className="btn ghost" disabled={!!busy} onClick={syncRu}>{wait('sync', 'Янгиланмоқда…', '↻ Русчани ўзбекчага мослаш')}</button>
            <button type="button" className="btn ghost" disabled={!!busy} onClick={restart}>＋ Янги мақола</button>
          </div>
        </section>
      )}
    </div>
  );
}
