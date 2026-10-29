// Performance regression test: proves that selecting a block does NOT rebuild
// the whole pane's DOM. Uses a synthetic 200-block document so the cost of a
// full rebuild would be obvious. Run: node perf_test.mjs
import { JSDOM } from 'jsdom';
import { readFileSync } from 'fs';

const html = readFileSync('../templates/index.html', 'utf8');

// a big document so any accidental full-rebuild is easy to measure
const bigBlocks = {};
for (let i = 1; i <= 200; i++) {
  bigBlocks[String(i)] = { id:String(i), type:'paragraph', text:'Paragraph '+i, order:i };
}
const bigUpload = { doc_id:'big', filename:'big.docx', blocks:bigBlocks, can_undo:false };

const dom = new JSDOM(html, {
  runScripts:'dangerously', url:'http://localhost/',
  beforeParse(w){
    w.fetch = () => Promise.resolve({ ok:true, status:200, json:()=>Promise.resolve(bigUpload) });
    w.FormData = class { append(){} }; w.File = class {};
    w.alert = ()=>{}; w.confirm = ()=>true;
  },
});
const { window } = dom, { document } = window;
const wait = ms => new Promise(r=>setTimeout(r,ms));

let passed=0, failed=0;
const check=(n,c)=>{ if(c){passed++;console.log('  [PASS] '+n);} else {failed++;console.log('  [FAIL] '+n);} };

setTimeout(async () => {
  console.log('\n=== performance test (200 blocks) ===\n');
  const input = document.getElementById('file-base');
  Object.defineProperty(input,'files',{ value:[{name:'big.docx'}], configurable:true });
  await window.eval("uploadDoc('base')");
  await wait(60);
  console.log('   rendered', window.eval('Object.keys(state.base.blocks).length'), 'blocks');

  // measure DOM churn during a pure selection toggle
  let created = 0;
  const orig = document.createElement.bind(document);
  document.createElement = t => { created++; return orig(t); };

  const t0 = Date.now();
  window.eval("toggleBlock('base','5')");
  await wait(5);
  const dt = Date.now() - t0;
  document.createElement = orig;

  console.log(`   one selection toggle: ${created} nodes created, ${dt}ms`);
  check('selection creates zero DOM nodes', created === 0);
  check('selection completes under 30ms', dt < 30);

  // scroll position survives a selection
  const render = document.getElementById('render-base');
  render.scrollTop = 400;
  window.eval("toggleBlock('base','9')");
  await wait(5);
  check('scroll position preserved on selection', render.scrollTop === 400);

  console.log(`\n=== ${passed} passed, ${failed} failed ===`);
  process.exit(failed ? 1 : 0);
}, 100);
