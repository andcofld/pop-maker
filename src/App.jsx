import { useMemo, useRef, useState } from 'react';
import {
  Download,
  FileSpreadsheet,
  ImagePlus,
  Printer,
  Sparkles,
} from 'lucide-react';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import * as XLSX from 'xlsx';

const SIZE_PRESETS = {
  '1/2': { label: 'A4 1/2', units: 4, colSpan: 2, rowSpan: 2 },
  '1/4': { label: 'A4 1/4', units: 2, colSpan: 1, rowSpan: 2 },
  '1/8': { label: 'A4 1/8', units: 1, colSpan: 1, rowSpan: 1 },
};

const SAMPLE_ROWS = [
  { productName: '소화가 잘되는 데일리밀크 2.3L', salePrice: '6,980', size: '1/2' },
  { productName: '매일우유 오리지널 900ml', salePrice: '2,980', size: '1/4' },
  { productName: '상하목장 유기농 우유', salePrice: '4,480', size: '1/4' },
  { productName: '바리스타룰스 로어슈거', salePrice: '1,680', size: '1/8' },
  { productName: '피크닉 사과', salePrice: '990', size: '1/8' },
];

function normalize(value) {
  return String(value || '')
    .replace(/\.[a-z0-9]+$/i, '')
    .replace(/\s+/g, '')
    .toLowerCase();
}

function formatPrice(value) {
  const digits = String(value || '').replace(/\D/g, '');
  return digits ? Number(digits).toLocaleString('ko-KR') : '';
}

function normalizeSize(value) {
  const raw = String(value || '').trim();
  if (raw.includes('1/2') || raw.includes('2분의1') || raw.includes('반')) return '1/2';
  if (raw.includes('1/8') || raw.includes('8분의1')) return '1/8';
  return '1/4';
}

function pick(row, keys) {
  const foundKey = Object.keys(row).find((key) => keys.includes(String(key).trim()));
  return foundKey ? row[foundKey] : '';
}

function buildPages(items) {
  const sorted = [...items].sort((a, b) => SIZE_PRESETS[b.size].units - SIZE_PRESETS[a.size].units);
  const pages = [];

  sorted.forEach((item) => {
    const units = SIZE_PRESETS[item.size].units;
    let page = pages.find((candidate) => candidate.used + units <= 8);

    if (!page) {
      page = { id: pages.length + 1, used: 0, items: [] };
      pages.push(page);
    }

    page.items.push(item);
    page.used += units;
  });

  return pages;
}

export default function App() {
  const [rows, setRows] = useState(SAMPLE_ROWS);
  const [templates, setTemplates] = useState({});
  const [saving, setSaving] = useState(false);
  const previewRef = useRef(null);

  const popItems = useMemo(() => {
    return rows.map((row, index) => {
      const productName = row.productName || `상품 ${index + 1}`;
      const template = templates[normalize(productName)];

      return {
        id: `${productName}-${index}`,
        productName,
        salePrice: formatPrice(row.salePrice),
        size: normalizeSize(row.size),
        templateUrl: template?.url || '',
      };
    });
  }, [rows, templates]);

  const pages = useMemo(() => buildPages(popItems), [popItems]);
  const matchedCount = popItems.filter((item) => item.templateUrl).length;

  const uploadTemplates = (event) => {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;

    const nextTemplates = {};
    files.forEach((file) => {
      nextTemplates[normalize(file.name)] = {
        name: file.name,
        url: URL.createObjectURL(file),
      };
    });

    setTemplates((current) => ({ ...current, ...nextTemplates }));
  };

  const uploadExcel = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    const nextRows = jsonRows
      .map((row) => ({
        productName: pick(row, ['제품명', '상품명', 'productName', 'name']),
        salePrice: pick(row, ['행사가', '판매가', '가격', 'salePrice', 'price']),
        size: pick(row, ['사이즈', 'POP사이즈', 'popSize', 'size']),
      }))
      .filter((row) => row.productName && row.salePrice);

    setRows(nextRows);
  };

  const savePdf = async () => {
    if (!previewRef.current) return;

    setSaving(true);
    try {
      const pdf = new jsPDF('p', 'mm', 'a4');
      const pageNodes = Array.from(previewRef.current.querySelectorAll('[data-page]'));

      for (const [index, pageNode] of pageNodes.entries()) {
        const canvas = await html2canvas(pageNode, {
          scale: 2,
          backgroundColor: '#ffffff',
          useCORS: true,
        });

        if (index > 0) pdf.addPage('a4', 'p');
        pdf.addImage(canvas.toDataURL('image/png'), 'PNG', 0, 0, 210, 297);
      }

      pdf.save('store-pop-pages.pdf');
    } finally {
      setSaving(false);
    }
  };

  return (
    <main className="min-h-screen bg-slate-100 text-slate-950">
      <div className="mx-auto grid min-h-screen max-w-[1600px] grid-cols-1 gap-5 p-4 lg:grid-cols-[420px_1fr] lg:p-6">
        <section className="no-print flex flex-col rounded-lg border border-slate-200 bg-white p-5 shadow-lg">
          <div className="border-b border-slate-200 pb-4">
            <p className="text-sm font-black text-blue-700">마트 POP Maker</p>
            <h1 className="mt-1 text-2xl font-black">엑셀 일괄 출력</h1>
          </div>

          <div className="mt-5 space-y-4">
            <UploadButton
              icon={<ImagePlus size={22} />}
              label="제품별 POP 템플릿 이미지"
              value={`${Object.keys(templates).length}개 등록됨`}
              accept="image/*"
              multiple
              onChange={uploadTemplates}
            />

            <UploadButton
              icon={<FileSpreadsheet size={22} />}
              label="제품명/행사가 엑셀"
              value={`${rows.length}개 상품 불러옴`}
              accept=".xlsx,.xls,.csv"
              onChange={uploadExcel}
            />

            <div className="rounded-lg border border-blue-100 bg-blue-50 p-4">
              <div className="flex items-center gap-2 font-black text-blue-800">
                <Sparkles size={18} />
                <span>자동 배치 결과</span>
              </div>
              <dl className="mt-3 grid grid-cols-3 gap-2 text-center">
                <Stat label="상품" value={popItems.length} />
                <Stat label="A4" value={pages.length} />
                <Stat label="매칭" value={`${matchedCount}/${popItems.length}`} />
              </dl>
            </div>

            <div className="rounded-lg border border-slate-200">
              <div className="grid grid-cols-[1fr_70px_62px] gap-2 border-b border-slate-200 bg-slate-50 px-3 py-2 text-xs font-black text-slate-500">
                <span>제품명</span>
                <span>행사가</span>
                <span>크기</span>
              </div>
              <div className="max-h-[330px] overflow-auto">
                {popItems.map((item) => (
                  <div
                    key={item.id}
                    className="grid grid-cols-[1fr_70px_62px] gap-2 border-b border-slate-100 px-3 py-2 text-sm last:border-b-0"
                  >
                    <span className="truncate font-bold">{item.productName}</span>
                    <span className="font-black text-red-600">{item.salePrice}</span>
                    <span className="font-bold text-slate-600">{item.size}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>

          <div className="mt-auto grid grid-cols-2 gap-3 pt-6">
            <button
              type="button"
              onClick={() => window.print()}
              className="flex h-12 items-center justify-center gap-2 rounded-lg bg-blue-700 font-black text-white"
            >
              <Printer size={19} />
              <span>인쇄하기</span>
            </button>

            <button
              type="button"
              onClick={savePdf}
              disabled={saving}
              className="flex h-12 items-center justify-center gap-2 rounded-lg bg-red-600 font-black text-white disabled:opacity-60"
            >
              <Download size={19} />
              <span>{saving ? '저장 중' : 'PDF 저장'}</span>
            </button>
          </div>
        </section>

        <section className="min-w-0">
          <div className="no-print mb-3 flex items-end justify-between px-1">
            <div>
              <p className="text-sm font-black text-blue-700">실시간 미리보기</p>
              <h2 className="text-xl font-black">A4 자동 배치</h2>
            </div>
          </div>

          <div
            ref={previewRef}
            className="flex flex-col items-center gap-5 overflow-auto rounded-lg border border-slate-300 bg-slate-200 p-5 print:block print:overflow-visible print:border-0 print:bg-white print:p-0"
          >
            {pages.map((page) => (
              <A4Page key={page.id} page={page} />
            ))}
          </div>
        </section>
      </div>
    </main>
  );
}

function UploadButton({ icon, label, value, accept, multiple = false, onChange }) {
  return (
    <label className="block">
      <span className="mb-2 block text-sm font-black text-slate-700">{label}</span>
      <div className="flex h-14 cursor-pointer items-center justify-between gap-3 rounded-lg border border-dashed border-slate-400 bg-slate-50 px-4 font-black text-slate-800">
        <span className="flex min-w-0 items-center gap-2">
          {icon}
          <span className="truncate">{value}</span>
        </span>
        <span className="rounded-md bg-white px-3 py-1 text-sm text-blue-700">선택</span>
      </div>
      <input className="hidden" type="file" accept={accept} multiple={multiple} onChange={onChange} />
    </label>
  );
}

function Stat({ label, value }) {
  return (
    <div className="rounded-md bg-white p-2">
      <dt className="text-xs font-black text-slate-500">{label}</dt>
      <dd className="text-xl font-black text-slate-950">{value}</dd>
    </div>
  );
}

function A4Page({ page }) {
  return (
    <div
      data-page
      className="grid aspect-[210/297] w-full max-w-[794px] auto-rows-fr grid-cols-2 gap-0 bg-white shadow-2xl print:h-[297mm] print:w-[210mm] print:max-w-none print:break-after-page print:shadow-none"
      style={{ gridTemplateRows: 'repeat(4, minmax(0, 1fr))' }}
    >
      {page.items.map((item) => (
        <PopTile key={item.id} item={item} />
      ))}
    </div>
  );
}

function PopTile({ item }) {
  const preset = SIZE_PRESETS[item.size];

  return (
    <article
      className="relative min-h-0 min-w-0 overflow-hidden border border-slate-300 bg-white"
      style={{
        gridColumn: `span ${preset.colSpan}`,
        gridRow: `span ${preset.rowSpan}`,
        containerType: 'inline-size',
      }}
    >
      {item.templateUrl ? (
        <img className="h-full w-full object-fill" src={item.templateUrl} alt="" />
      ) : (
        <div className="flex h-full w-full flex-col justify-between bg-white p-[5cqw]">
          <div>
            <div className="inline-flex rounded-full bg-blue-700 px-[3cqw] py-[1cqw] text-[5cqw] font-black text-white">
              Maeil
            </div>
            <h3 className="mt-[4cqw] break-keep text-[7cqw] font-black leading-tight text-blue-900 [overflow-wrap:anywhere]">
              {item.productName}
            </h3>
          </div>
          <div className="text-right">
            <span className="rounded bg-red-600 px-[2cqw] py-[1cqw] text-[4cqw] font-black text-white">
              행사가
            </span>
            <div className="text-[14cqw] font-black leading-none text-red-600">
              {item.salePrice}
              <span className="text-[5cqw]">원</span>
            </div>
          </div>
        </div>
      )}

      {item.templateUrl && (
        <div className="absolute bottom-[7%] right-[6%] whitespace-nowrap text-right font-black leading-none text-red-600">
          <span className="text-[18cqw]">{item.salePrice}</span>
          <span className="text-[7cqw]">원</span>
        </div>
      )}
    </article>
  );
}
