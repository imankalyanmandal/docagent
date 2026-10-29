// Loads the REAL templates/index.html in jsdom and drives the actual UI code,
// with fetch backed by responses captured from the real Flask backend.
import { JSDOM } from 'jsdom';
import { readFileSync } from 'fs';

const html = readFileSync('../templates/index.html', 'utf8');
const fx = JSON.parse(readFileSync('../fixtures.json', 'utf8'));

let passed=0, failed=0;
const check=(n,c)=>{ if(c){passed++;console.log('  [PASS] '+n);} else {failed++;console.log('  [FAIL] '+n);} };

// track every request the page makes so we can assert on the sequence
const requestLog = [];
function fakeFetch(url,opts={}){
  let p=null; if(typeof opts.body==='string'){try{p=JSON.parse(opts.body)}catch{}}
  requestLog.push({url, payload:p});
  const ok=(o,s=200)=>Promise.resolve({ok:s<400,status:s,json:()=>Promise.resolve(o)});

  if(url.endsWith('/upload')){
    // distinguish which fixture doc by filename the test set on the input
    return ok(fakeFetch._nextUpload || fx.upload);
  }
  if(url.endsWith('/edit_text')) return ok(fx.edit_text.body, fx.edit_text.status);
  if(url.endsWith('/upload_html')){
    if (p && p.append_to) return ok(fx.upload_html_append.body, fx.upload_html_append.status);
    return ok(fx.upload_html_new.body, fx.upload_html_new.status);
  }
  if(url.endsWith('/duplicate')) return ok(fx.duplicate.body, fx.duplicate.status);
  if(url.endsWith('/find_replace')) return ok(fx.find_replace.body, fx.find_replace.status);
  if(url.endsWith('/table')){
    const op = p.operation;
    if (op==='set_content' && fakeFetch._mergeMode) return ok(fx.merge_set_content.body, fx.merge_set_content.status);
    if (fx[op]) return ok(fx[op].body, fx[op].status);
    return ok({error:'no fixture for '+op}, 400);
  }
  if(url.endsWith('/delete') && fakeFetch._mergeMode) return ok(fx.merge_delete.body, fx.merge_delete.status);
  return ok({doc_id:p?.doc_id, blocks:fx.upload.blocks, can_undo:true});
}

const notes=[];
const dom=new JSDOM(html,{runScripts:'dangerously',url:'http://localhost/',
  beforeParse(w){ w.fetch=(u,o)=>fakeFetch(String(u),o); w.FormData=class{append(){}}; w.File=class{};
    w.alert=m=>notes.push('alert:'+m); w.confirm=()=>true; }});
const { window }=dom; const { document }=window;
const $=id=>document.getElementById(id);
const wait=ms=>new Promise(r=>setTimeout(r,ms));
const ev=e=>window.eval(e);
const tableId=()=>ev("Object.values(state.base.blocks).find(b=>b.type==='table')?.id");
const rowsOf=id=>ev(`state.base.blocks['${id}'].rows.length`);
const colsOf=id=>ev(`state.base.blocks['${id}'].rows[0].length`);

async function uploadInto(pane='base'){
  const input=$(`file-${pane}`);
  Object.defineProperty(input,'files',{value:[{name:'x.docx'}],configurable:true});
  await ev(`uploadDoc('${pane}')`);
  await wait(60);
}

(async()=>{
  console.log('\n=== Issue 4: server config (reloader off, threaded on) ===\n');
  const appSrc = readFileSync('../app.py','utf8');
  check('use_reloader=False set', /use_reloader\s*=\s*False/.test(appSrc));
  check('threaded=True set', /threaded\s*=\s*True/.test(appSrc));
  check('global exception handler present', /@app\.errorhandler\(Exception\)/.test(appSrc));

  console.log('\n=== Issue 2: text editing (paragraphs/headers) ===\n');
  fakeFetch._nextUpload = fx.upload;
  await uploadInto('base');
  check('upload populated blocks', ev('!!state.base.blocks'));
  check('startTextEdit defined', ev('typeof startTextEdit')==='function');

  // find a paragraph block's edit button (first button in its control cluster) and click it
  const paraId = ev("Object.values(state.base.blocks).find(b=>b.type==='paragraph').id");
  const wrapper = document.querySelector(`[data-block-id="${paraId}"]`);
  check('paragraph block rendered', !!wrapper);
  const editBtn = wrapper.querySelector('.move-controls button');   // the ✎ edit button
  check('paragraph has an Edit (pencil) button', !!editBtn && editBtn.textContent === '\u270E');
  editBtn.click();
  const textEl = wrapper.querySelector('.paragraph-text');
  check('text becomes contentEditable on click', textEl.contentEditable === 'true');
  textEl.textContent = 'Edited via test';
  textEl.dispatchEvent(new window.Event('blur'));
  await wait(60);
  const editReq = requestLog.find(r => r.url.endsWith('/edit_text'));
  check('edit dispatched a request to /edit_text', !!editReq);
  check('edit_text payload carries the new text', editReq && editReq.payload.text === 'Edited via test');

  console.log('\n=== Issue 1/3: table popup editor (values + sort), not inline ===\n');
  requestLog.length = 0;
  const tid = tableId();
  // table blocks must NOT get a plain contenteditable text node
  const tblWrapper = document.querySelector(`[data-block-id="${tid}"]`);
  check('table block has no header-text/paragraph-text (not inline-editable)',
        !tblWrapper.querySelector('.header-text') && !tblWrapper.querySelector('.paragraph-text'));
  const tblEditBtn = tblWrapper.querySelector('.move-controls button');   // first = ✎ edit
  check('table has an Edit (pencil) button', !!tblEditBtn && tblEditBtn.textContent === '\u270E');

  console.log('\n=== Discoverability: controls must be visible WITHOUT hovering ===\n');
  // jsdom computes styles straight from the stylesheet without simulating
  // :hover, so this catches the earlier bug where display:none gated behind
  // :hover made the controls invisible until you moused over the exact block.
  const moveWrapPara = wrapper.querySelector('.move-controls');
  const moveWrapTbl = tblWrapper.querySelector('.move-controls');
  check('paragraph control cluster visible without hover',
        !!moveWrapPara && window.getComputedStyle(moveWrapPara).display !== 'none');
  check('table control cluster visible without hover',
        !!moveWrapTbl && window.getComputedStyle(moveWrapTbl).display !== 'none');
  check('no stylesheet rule hides .move-controls by default',
        !/\.move-controls\s*\{[^}]*display\s*:\s*none/.test(html));

  ev(`openSingleTableEditor('${tid}')`);
  await wait(30);
  check('editor modal opens', !$('editor-modal').classList.contains('hidden'));
  const grid = $('em-grid');

  // #2: header row shows the table's ACTUAL first-row values, not "Col 1"
  const headerInputs = grid.querySelectorAll('tr:first-child .th-input');
  check('header shows real header cell inputs (not Col N labels)', headerInputs.length === colsOf(tid));
  check('header input carries actual first-row value',
        headerInputs[0].value === ev(`state.base.blocks['${tid}'].rows[0][0]`));
  check('no "Col 1" placeholder text used as header label',
        !/>Col \d+</.test(grid.querySelector('tr:first-child').innerHTML));

  // data-cell inputs = (rows-1) x cols, since row 0 is the header
  const dataInputs = [...grid.querySelectorAll('tr:not(:first-child) td input')];
  check('grid renders an <input> per data cell', dataInputs.length === (rowsOf(tid)-1)*colsOf(tid));

  // edit a data cell value
  const firstDataInput = grid.querySelector('tr:nth-child(2) td:nth-child(2) input');
  firstDataInput.value = 'EDITED-VALUE';
  firstDataInput.dispatchEvent(new window.Event('input'));
  check('editor.rows[1][0] reflects the typed edit', ev("editor.rows[1][0]") === 'EDITED-VALUE');

  // #4: sort control + delete-column control live inside the popup header
  check('per-column sort button exists in popup', !!grid.querySelector('.th-sort'));
  check('per-column delete-column button exists in popup', !!grid.querySelector('.th-delcol'));

  // #3: data rows have up/down reorder controls
  check('data rows have up/down move controls', grid.querySelectorAll('.row-move').length >= 2);

  // test row reordering actually swaps rows (and keeps header pinned)
  const beforeReorder = ev("JSON.stringify(editor.rows[1])");
  ev("moveEditorRow(2,-1)");   // move 2nd data row up
  await wait(10);
  check('moveEditorRow swapped data rows', ev("JSON.stringify(editor.rows[1])") !== beforeReorder);
  check('header row stayed pinned at index 0 after reorder',
        ev(`editor.rows[0][0]`) === headerInputs[0].value);

  // sort via the column sort button; header must remain row 0
  const sortBtn = $('em-grid').querySelector('.th-sort');
  sortBtn.click();
  await wait(10);
  check('sort set a sortCol', ev('editor.sortCol') !== null);
  check('header still pinned at row 0 after sort',
        ev(`editor.rows[0][0]`) === headerInputs[0].value);

  // delete a column inside the popup (#4)
  const colsBefore = ev('editor.rows[0].length');
  $('em-grid').querySelector('.th-delcol').click();
  await wait(10);
  check('editorDeleteColumn removed a column', ev('editor.rows[0].length') === colsBefore - 1);

  // save -> should call /table set_content
  requestLog.length = 0;
  await ev('saveEditor()');
  await wait(30);
  const saveReq = requestLog.find(r => r.url.endsWith('/table') && r.payload.operation==='set_content');
  check('Save calls /table set_content', !!saveReq);
  check('modal closes after save', $('editor-modal').classList.contains('hidden'));

  console.log('\n=== #5 export must not navigate away, #6 focus mode ===\n');
  check('doExport uses a download anchor, not location.href',
        /a\.download/.test(html) && !/location\.href\s*=\s*`\/export/.test(html));
  check('focus mode toggle defined', ev('typeof toggleFocusMode') === 'function');
  ev('toggleFocusMode()');
  check('focus-mode class added to body', document.body.classList.contains('focus-mode'));
  check('CSS hides other pane in focus mode', /body\.focus-mode\s+#pane-other\s*\{\s*display\s*:\s*none/.test(html));
  ev('toggleFocusMode()');
  check('focus mode toggles back off', !document.body.classList.contains('focus-mode'));

  console.log('\n=== Duplicate, Find & Replace, Outline (new features) ===\n');
  // ensure a base doc is loaded (upload fixture has blocks incl. a table)
  fakeFetch._nextUpload = fx.upload;
  await uploadInto('base');

  // Duplicate: enabled when blocks are checked, sends /duplicate
  const anyId = ev("Object.keys(state.base.blocks)[0]");
  ev(`state.base.checked = new Set(['${anyId}']); updateToolbar();`);
  check('Duplicate button enabled with a checked block', $('duplicate-btn').disabled === false);
  requestLog.length = 0;
  await ev('doDuplicate()');
  await wait(30);
  check('Duplicate sends /duplicate with block_ids', requestLog.some(r => r.url.endsWith('/duplicate') && r.payload.block_ids));

  // Find & replace modal + request
  check('Find & replace button enabled with a base doc', $('findrep-btn').disabled === false);
  ev('openFindReplace()');
  check('find/replace modal opens', !$('findrep-modal').classList.contains('hidden'));
  $('fr-find').value = ''; 
  await ev('doFindReplace()');
  check('empty find shows a hint (no request)', /enter text to find/i.test($('fr-status').textContent));
  requestLog.length = 0;
  $('fr-find').value = 'Ann'; $('fr-replace').value = 'Zoe'; $('fr-scope').value = 'all';
  await ev('doFindReplace()');
  await wait(30);
  const frReq = requestLog.find(r => r.url.endsWith('/find_replace'));
  check('find/replace sends the query', !!frReq && frReq.payload.find === 'Ann' && frReq.payload.replace === 'Zoe');
  check('find/replace reports a replacement count', /replacement/i.test($('fr-status').textContent));

  // Outline sidebar + word count
  check('toggleOutline defined', ev('typeof toggleOutline') === 'function');
  ev('toggleOutline()');
  check('outline drawer opens', !$('outline-drawer').classList.contains('hidden'));
  const outlineItems = document.querySelectorAll('.outline-item').length;
  check('outline lists the document headings', outlineItems >= 1);
  check('word count is shown', /\d+\s+words/.test($('outline-wordcount').textContent));
  // countWords sanity
  check('countWords counts table + text words',
        ev("countWords({a:{type:'paragraph',text:'one two'},b:{type:'table',rows:[['three','four']]}})") === 4);
  ev('toggleOutline()');
  check('outline drawer closes', $('outline-drawer').classList.contains('hidden'));

  console.log('\n=== HTML paste (Froala) ===\n');
  requestLog.length = 0;
  check('openHtmlPaste defined', ev('typeof openHtmlPaste') === 'function');
  ev('openHtmlPaste()');
  check('HTML modal opens', !$('html-modal').classList.contains('hidden'));

  // empty submit is guarded
  $('html-input').value = '';
  await ev("submitHtml('base-doc')");
  check('empty HTML shows a hint, no request', requestLog.length === 0 && /paste some html/i.test($('html-status').textContent));

  // load as a NEW base document
  $('html-input').value = '<h1>Pasted</h1><p>Text</p><table><tr><th>N</th><th>V</th></tr><tr><td>a</td><td>1</td></tr></table>';
  await ev("submitHtml('base-doc')");
  await wait(40);
  const htmlReq = requestLog.find(r => r.url.endsWith('/upload_html'));
  check('paste sends /upload_html', !!htmlReq);
  check('load-as-base did NOT send append_to', htmlReq && !htmlReq.payload.append_to);
  check('base pane now has the pasted blocks', ev('!!state.base.blocks'));
  check('HTML modal closed after submit', $('html-modal').classList.contains('hidden'));

  // append into the existing base doc
  requestLog.length = 0;
  ev('openHtmlPaste()');
  $('html-input').value = '<h2>Appended</h2><p>More</p>';
  await ev("submitHtml('append-base')");
  await wait(40);
  const appendReq = requestLog.find(r => r.url.endsWith('/upload_html'));
  check('append-to-base sends append_to with the base doc_id', !!appendReq && !!appendReq.payload.append_to);

  // load as OTHER document
  requestLog.length = 0;
  ev('openHtmlPaste()');
  $('html-input').value = '<h1>Ref</h1><p>ref text</p>';
  await ev("submitHtml('other-doc')");
  await wait(40);
  check('load-as-other populated the other pane', ev('!!state.other.blocks'));

  console.log('\n=== Merge scenario: multiple tables -> popup -> shuffle/delete -> save ===\n');
  fakeFetch._nextUpload = fx.upload_merge;
  fakeFetch._mergeMode = true;
  await uploadInto('base');
  const tableIds = Object.values(ev('state.base.blocks'))
    .filter(b => b.type==='table').sort((a,b)=>a.order-b.order).map(b=>b.id);
  check('merge fixture has 2 tables', tableIds.length === 2);

  ev(`toggleBlock('base','${tableIds[0]}')`);
  ev(`toggleBlock('base','${tableIds[1]}')`);
  await wait(20);
  check('merge-btn enables with 2 tables checked', $('merge-btn').disabled === false);

  ev('openMergeEditor()');
  await wait(20);
  check('identical headers: no header-choice prompt', $('merge-header-modal').classList.contains('hidden'));
  check('merge modal opens', !$('editor-modal').classList.contains('hidden'));
  const mergedRowCount = ev('editor.rows.length');
  console.log(`   combined rows in popup: ${mergedRowCount}`);
  check('duplicate header row auto-dropped', mergedRowCount === 3); // 1 header + 2 data rows total
  check('rows carry a source tag for traceability', ev('editor.tags.filter(t=>t).length') > 0);

  // --- mismatched-header merge must PROMPT, not silently demote a header ---
  console.log('\n=== merge with DIFFERENT headers -> must ask which to use ===\n');
  ev('closeEditor()');
  ev(`state.base.doc_id='dmix';
      state.base.blocks = {
        'a': {id:'a',type:'table',order:1,rows:[['Name','Dept'],['Ann','Eng']]},
        'b': {id:'b',type:'table',order:2,rows:[['Employee','Team'],['Bob','Ops']]}
      };
      state.base.checked = new Set(['a','b']);`);
  ev('openMergeEditor()');
  await wait(20);
  check('mismatched headers: header-choice prompt appears', !$('merge-header-modal').classList.contains('hidden'));
  check('prompt lists one option per table', document.querySelectorAll('.mh-option').length === 2);
  check('editor NOT opened until a header is chosen', $('editor-modal').classList.contains('hidden'));

  // choose the SECOND table's header, drop the others
  document.querySelector('input[name="mh-choice"][value="1"]').checked = true;
  $('mh-drop-others').checked = true;
  ev('confirmMergeHeader()');
  await wait(20);
  check('after choosing, editor opens', !$('editor-modal').classList.contains('hidden'));
  const mrows = JSON.parse(ev('JSON.stringify(editor.rows)'));
  check('chosen header is row 0', JSON.stringify(mrows[0]) === JSON.stringify(['Employee','Team']));
  check('the OTHER header did NOT leak into data rows',
        !mrows.slice(1).some(r => JSON.stringify(r) === JSON.stringify(['Name','Dept'])));
  check('data rows from both tables present',
        mrows.some(r=>r[0]==='Ann') && mrows.some(r=>r[0]==='Bob'));
  ev('closeEditor()');

  // reset state back to the merge fixture for the save-path test below
  fakeFetch._nextUpload = fx.upload_merge;
  fakeFetch._mergeMode = true;
  await uploadInto('base');
  const tableIds2 = Object.values(ev('state.base.blocks'))
    .filter(b => b.type==='table').sort((a,b)=>a.order-b.order).map(b=>b.id);
  ev(`toggleBlock('base','${tableIds2[0]}')`);
  ev(`toggleBlock('base','${tableIds2[1]}')`);
  await wait(20);
  ev('openMergeEditor()');
  await wait(20);

  // simulate the scenario: delete one row, then save
  const grid2 = $('em-grid');
  const beforeDelCount = ev('editor.rows.length');
  const delButtons = grid2.querySelectorAll('.row-del-btn');
  delButtons[delButtons.length-1].click();
  await wait(10);
  check('row deleted from popup before saving', ev('editor.rows.length') === beforeDelCount - 1);

  requestLog.length = 0;
  await ev('saveEditor()');
  await wait(40);
  const mergeSetContent = requestLog.find(r => r.url.endsWith('/table') && r.payload.operation==='set_content');
  const mergeDeleteReq  = requestLog.find(r => r.url.endsWith('/delete'));
  check('merge save writes combined rows via set_content', !!mergeSetContent && mergeSetContent.payload.block_id === tableIds[0]);
  check('merge save removes the OTHER table via /delete', !!mergeDeleteReq && mergeDeleteReq.payload.block_ids.includes(tableIds[1]));

  if(notes.length){ console.log('\n  alerts:',notes); }
  console.log(`\n=== ${passed} passed, ${failed} failed ===`);
  process.exit(failed?1:0);
})().catch(e=>{console.error('CRASH:',e);process.exit(2);});
