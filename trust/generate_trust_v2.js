'use strict';

/**
 * 自社株管理処分信託契約書 生成エンジン v2
 * 使い方: node generate_trust_v2.js clients/田中商事.json [output.docx]
 *
 * ファイル構成:
 *   generate_trust_v2.js   ← このファイル（エンジン。基本触らない）
 *   templates/
 *     articles.json        ← 条文テキスト（条文修正はここだけ）
 *     footnotes.json       ← 脚注テキスト（脚注修正はここだけ）
 *   clients/
 *     template.json        ← クライアント情報の雛形
 *     田中商事.json         ← クライアントごとに複製して使う
 *   output/                ← 生成されたdocxが入る
 */

const {
  Document, Packer, Paragraph, TextRun, AlignmentType, FootnoteReferenceRun,
} = require('docx');
const fs   = require('fs');
const path = require('path');

// ===== パス解決 =====
const SCRIPT_DIR   = path.dirname(path.resolve(process.argv[1]));
const ARTICLES_PATH  = path.join(SCRIPT_DIR, 'templates', 'articles.json');
const FOOTNOTES_PATH = path.join(SCRIPT_DIR, 'templates', 'footnotes.json');

// ===== CLI引数 =====
const args       = process.argv.slice(2);
const inputFile  = args[0];
const outputFile = args[1];

// ===== クライアントデータ読み込み =====
const P = { address: '　', name: '　', dob: '　' };
const DEFAULTS = {
  companyName: '［対象会社名］', shareCount: '［株数］',
  trustMoney: '［金額］', courtName: '［管轄地方裁判所名］',
  trustor:           {...P}, trustee:          {...P},
  successorTrustee:  {...P}, clearanceTrustee: {...P},
  beneficiaryType: 'A',
  beneficiary1:      {...P}, beneficiary2:     {...P}, beneficiary3: {...P},
  directionType: 'A', triggerDays: '２週間',
  reduceDuty: false, terminationType: 'death', residualType: 'A',
  residualPerson:    {...P}, residualPersonAlt:{...P},
};
const NESTED = [
  'trustor','trustee','successorTrustee','clearanceTrustee',
  'beneficiary1','beneficiary2','beneficiary3',
  'residualPerson','residualPersonAlt'
];

let data;
if (inputFile) {
  try {
    const loaded = JSON.parse(fs.readFileSync(inputFile, 'utf-8'));
    data = Object.assign({}, DEFAULTS, loaded);
    NESTED.forEach(k => {
      data[k] = Object.assign({}, P, loaded[k] || {});
    });
    console.log(`📂 入力: ${path.resolve(inputFile)}`);
  } catch(e) {
    console.error(`❌ JSON読み込み失敗: ${e.message}`); process.exit(1);
  }
} else {
  data = Object.assign({}, DEFAULTS);
  NESTED.forEach(k => { data[k] = {...P}; });
  console.log('ℹ️  入力JSONなし。デフォルト値で生成します。');
}

// ===== テンプレート読み込み =====
let ARTICLES, FOOTNOTES;
try {
  ARTICLES  = JSON.parse(fs.readFileSync(ARTICLES_PATH, 'utf-8'));
  FOOTNOTES = JSON.parse(fs.readFileSync(FOOTNOTES_PATH, 'utf-8'));
  console.log(`📋 テンプレート読み込み完了`);
} catch(e) {
  console.error(`❌ テンプレート読み込み失敗: ${e.message}`); process.exit(1);
}

// ===== スタイル定数 =====
const FONT         = "MS明朝";
const SIZE         = 20;   // 10pt
const SIZE_TITLE   = 28;   // 14pt
const SIZE_SUBTITLE = 24;  // 12pt
const SIZE_FN      = 16;   // 8pt（脚注）

// ===== ユーティリティ =====

/** {key} プレースホルダーをdataで展開 */
function expand(text) {
  return text.replace(/\{([^}]+)\}/g, (_, key) => {
    const parts = key.split('.');
    let val = data;
    for (const p of parts) val = val?.[p];
    return val ?? `{${key}}`;
  });
}

function run(text, opts = {}) {
  return new TextRun({ text: expand(text), font: FONT, size: SIZE, ...opts });
}

function para(text, opts = {}) {
  return new Paragraph({
    children: typeof text === 'string' ? [run(text)] : text,
    alignment: opts.align || AlignmentType.JUSTIFIED,
    spacing: { before: 60, after: 60, line: 360 },
    indent: opts.indent,
  });
}

function itemPara(text) {
  return new Paragraph({
    children: [run(text)],
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 40, after: 40, line: 360 },
    indent: { left: 440, hanging: 440 },
  });
}

function blank() {
  return new Paragraph({ children: [run("")], spacing: { before: 60, after: 60 } });
}

function articleHeader(num, title) {
  return new Paragraph({
    children: [new TextRun({ text: `第${num}条　${title}`, font: FONT, size: SIZE, bold: true })],
    alignment: AlignmentType.LEFT,
    spacing: { before: 300, after: 80 },
  });
}

// ===== 脚注番号割り当て =====
const fnContents = {};  // { id(number): text }
const fnKeyToId  = {};  // { key: id }
let fnSeq = 1;

function registerFn(key) {
  if (fnKeyToId[key] !== undefined) return fnKeyToId[key];
  const text = FOOTNOTES[key];
  if (!text) { console.warn(`⚠️  脚注キー未定義: ${key}`); return null; }
  const id = fnSeq++;
  fnContents[id] = text;
  fnKeyToId[key] = id;
  return id;
}

// 全条文の脚注キーを事前登録（IDを連番で確定させる）
function preRegisterFns() {
  for (const art of ARTICLES.articles) {
    if (art.fnKey) registerFn(art.fnKey);
    // para_fn2 の追加脚注
    if (art.paragraphs) {
      for (const p of art.paragraphs) {
        if (p.type === 'para_fn2' && p.fnKey) registerFn(p.fnKey);
      }
    }
  }
}

// ===== 段落レンダリング =====
function renderParagraph(p, articleFnKey) {
  const result = [];
  switch(p.type) {

    case 'para_fn': {
      const id = fnKeyToId[articleFnKey];
      result.push(para([run(p.text), ...(id != null ? [new FootnoteReferenceRun(id)] : [])]));
      break;
    }
    case 'para_fn2': {
      const id = fnKeyToId[p.fnKey];
      result.push(para([run(p.text), ...(id != null ? [new FootnoteReferenceRun(id)] : [])]));
      break;
    }
    case 'para_fn_reduceDuty': {
      const text = data.reduceDuty ? p.textReduceDuty : p.text;
      const id   = fnKeyToId[articleFnKey];
      result.push(para([run(text), ...(id != null ? [new FootnoteReferenceRun(id)] : [])]));
      break;
    }
    case 'para':
      result.push(para(p.text));
      break;
    case 'item':
      result.push(itemPara(p.text));
      break;
    case 'label':
      result.push(para([new TextRun({ text: expand(p.text), font: FONT, size: SIZE })],
        { align: AlignmentType.LEFT }));
      break;
    case 'item_beneficiary7': {
      const text = data.beneficiaryType === 'A' ? p.textA : p.textB;
      result.push(itemPara(text));
      break;
    }
  }
  return result;
}

// ===== 条文レンダリング =====
function renderArticle(art) {
  const result = [];
  result.push(articleHeader(art.num, art.title));

  // 分岐あり条文（第14条など）
  if (art.condition && art.variants) {
    const key = data[art.condition] || 'A';
    const paragraphs = art.variants[key] || art.variants['A'];
    for (const p of paragraphs) result.push(...renderParagraph(p, art.fnKey));
  } else if (art.paragraphs) {
    for (const p of art.paragraphs) result.push(...renderParagraph(p, art.fnKey));
  }

  result.push(blank());
  return result;
}

// ===== メイン組み立て =====
function buildDocument() {
  const children = [];

  // タイトル・前文
  children.push(
    new Paragraph({
      children: [new TextRun({ text: ARTICLES.title, font: FONT, size: SIZE_TITLE, bold: true })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 80 },
    }),
    new Paragraph({
      children: [new TextRun({ text: ARTICLES.subtitle, font: FONT, size: SIZE_SUBTITLE })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 40, after: 200 },
    }),
    blank(),
  );
  for (const line of ARTICLES.preamble) {
    children.push(para(line));
  }
  children.push(blank());

  // 条文
  for (const art of ARTICLES.articles) {
    children.push(...renderArticle(art));
  }

  // 末尾署名欄
  for (const line of ARTICLES.closing) {
    children.push(para(line));
  }
  children.push(
    blank(),
    para("【委託者】"),
    para(`　住　所：${data.trustor.address || '　'}`),
    para(`　氏　名：${data.trustor.name || '　'}　　　　　　　　　印`),
    blank(),
    para("【受託者】"),
    para(`　住　所：${data.trustee.address || '　'}`),
    para(`　氏　名：${data.trustee.name || '　'}　　　　　　　　　印`),
    blank(),
  );

  return children;
}

// ===== 脚注オブジェクト構築 =====
function buildFootnotes() {
  const footnotes = {};
  for (const [id, text] of Object.entries(fnContents)) {
    footnotes[Number(id)] = {
      children: [new Paragraph({
        children: [new TextRun({ text, font: FONT, size: SIZE_FN })],
        spacing: { before: 0, after: 0 },
      })]
    };
  }
  return footnotes;
}

// ===== 実行 =====
async function main() {
  preRegisterFns();

  const children  = buildDocument();
  const footnotes = buildFootnotes();

  const doc = new Document({
    footnotes,
    styles: { default: { document: { run: { font: FONT, size: SIZE } } } },
    sections: [{
      properties: {
        page: {
          size:   { width: 11906, height: 16838 },          // A4
          margin: { top: 1701, right: 1701, bottom: 1701, left: 1701 }, // 30mm
        },
      },
      children,
    }]
  });

  const buffer   = await Packer.toBuffer(doc);
  const baseName = (data.companyName || '信託契約書').replace(/[\[\]／\/:*?"<>|]/g, '');
  const dest     = outputFile || `自社株管理処分信託契約書_${baseName}.docx`;
  fs.writeFileSync(dest, buffer);
  console.log(`✅ 生成完了: ${path.resolve(dest)}`);
}

main().catch(console.error);
