import { renderExcelSegmentsAsGrid } from './excelGrid.mjs';

const FIELD_ORDER = ['manager', 'stage', 'as_of', 'key_contacts', 'issue', 'confirmed', 'in_progress', 'unresolved', 'next_action'];
const FIELD_LABELS_CLIENT = {
  manager: '담당자', stage: '단계', as_of: '기준일', key_contacts: '주요 관계자',
  issue: '핵심 이슈', confirmed: '확정사항', in_progress: '협의 중 사항',
  unresolved: '미결사항', next_action: '다음 조치', checklist: '체크리스트',
};
let state = { scope: 'project', projectId: null, projectName: '', fundName: '' };

async function getJson(url) { return (await fetch(url)).json(); }
async function postJson(url, body) {
  const r = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  return r.json();
}
function esc(s) { return (s ?? '').toString().replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }

async function loadTree() {
  const data = await getJson('/api/tree');
  const tree = document.getElementById('tree');
  tree.innerHTML = '';
  for (const fund of data.funds) {
    const fundEl = document.createElement('div');
    fundEl.className = 'fund-row';
    fundEl.textContent = fund.name;
    tree.appendChild(fundEl);
    for (const p of fund.projects) {
      const el = document.createElement('div');
      el.className = 'project-row';
      el.textContent = p.name;
      el.dataset.projectId = p.id;
      el.addEventListener('click', () => selectProject(p.id, p.name, fund.name));
      tree.appendChild(el);
    }
  }
  if (data.funds[0]?.projects[0]) {
    const p = data.funds[0].projects[0];
    selectProject(p.id, p.name, data.funds[0].name);
  }
}

function setActiveProjectRow(id) {
  document.querySelectorAll('.project-row').forEach((el) => el.classList.toggle('active', Number(el.dataset.projectId) === id));
}

async function selectProject(id, name, fundName) {
  state = { scope: 'project', projectId: id, projectName: name, fundName };
  setActiveProjectRow(id);
  document.getElementById('crumb').textContent = `${fundName} > ${name}`;
  updateSearchScopeControl();
  runSearch();
  await renderStatusPanel();
  await renderDocumentsPanel();
  await renderActivitiesPanel();
  await renderSyncPanel();
  await renderReviewPanel();
  await renderAiPanel();
}

async function selectTeam() {
  state = { scope: 'team', projectId: null, projectName: '팀 공통 지식', fundName: '' };
  setActiveProjectRow(-1);
  document.getElementById('crumb').textContent = '팀 공통 지식';
  updateSearchScopeControl();
  runSearch();
  await renderStatusPanel();
  await renderDocumentsPanel();
  document.getElementById('panel-activities').innerHTML = '<p class="hint">업무 이력은 사업장 단위입니다.</p>';
  /* 마감 A1 — 갱신 단위는 사업장·펀드가 아니라 체험판에서는 준비된 예시 개정본을 반영하는 것뿐이다. */
  document.getElementById('panel-sync').innerHTML = '<p class="hint">「예시 문서 갱신 체험」은 왼쪽에서 사업장을 고른 뒤 「자료 갱신」 탭에서 실행합니다 — 실제 파일을 수집하지 않고 준비된 예시 개정본을 반영합니다.</p>';
  document.getElementById('side-review').innerHTML = '<p class="hint">변경 제안은 사업장 단위입니다.</p>';
  await renderAiPanel();
}

async function renderStatusPanel() {
  const panel = document.getElementById('panel-status');
  let items, docs = [];
  if (state.scope === 'team') {
    items = (await getJson('/api/team-status')).status;
  } else {
    const [data, docsData] = await Promise.all([
      getJson(`/api/projects/${state.projectId}`),
      getJson(`/api/documents?projectId=${state.projectId}`),
    ]);
    items = data.status;
    docs = docsData.documents.filter((d) => d.versions[0]);
  }
  items.sort((a, b) => FIELD_ORDER.indexOf(a.key) - FIELD_ORDER.indexOf(b.key));
  const docOptions = docs.map((d) => `<option value="${d.versions[0].id}">${esc(d.title)}</option>`).join('');

  /* 마감(B) — 담당자·단계처럼 한 줄이면 끝나는 항목이 각각 큰 카드를 차지하면, 정작
     읽어야 할 핵심 이슈·미결사항이 화면 아래로 밀린다. 짧은 항목은 요약 카드 한 장의
     여러 행으로 묶고, 긴 항목만 아래에 따로 둔다. 수정·이력·근거 연결은 두 형태 모두에
     그대로 남긴다(기능을 줄이지 않는다 — 배치만 바꾼다). */
  const SUMMARY_KEYS = ['manager', 'stage', 'as_of', 'key_contacts'];
  const evidenceHtml = (it) => (it.evidence.length
    ? it.evidence.map((e) => `<span class="evidence-chip" data-view="${e.doc_version_id}" data-segment="${e.segment_id ?? ''}">근거: ${esc(e.original_filename)}${e.page_no ? ' p.' + e.page_no : ''}${e.cell_range ? ' ' + esc(e.cell_range) : ''}</span>`).join(' ')
    : '<span class="hint">연결된 근거 없음</span>');
  const metaHtml = (it) => `v${it.versionNo}${it.effectiveDate ? ' · 기준일 ' + esc(it.effectiveDate) : ''} · ${esc(it.updatedBy)}, ${esc((it.updatedAt || '').slice(0, 16).replace('T', ' '))}`;
  const controlsHtml = (it) => `
      <div class="row-actions"><button class="edit-btn" data-edit="${it.id}">수정</button> <button class="edit-btn" data-history="${it.id}">이력(v${it.versionNo})</button></div>
      <div class="edit-form" data-form="${it.id}" style="display:none;margin-top:8px">
        <textarea rows="2" style="width:100%;padding:6px">${esc(it.value)}</textarea>
        ${docs.length ? `<select data-evidence="${it.id}" style="margin-top:6px"><option value="">근거 자료 선택 안 함</option>${docOptions}</select>` : ''}
        <button class="edit-btn" data-save="${it.id}" data-version="${it.versionNo}" style="margin-top:6px">저장</button>
      </div>`;

  const summary = items.filter((it) => SUMMARY_KEYS.includes(it.key));
  const rest = items.filter((it) => !SUMMARY_KEYS.includes(it.key));
  const summaryHtml = summary.length ? `
    <div class="status-card status-summary">
      <div class="k group-title">기본 정보</div>
      ${summary.map((it) => `
      <div class="sum-row" data-item-id="${it.id}" data-version="${it.versionNo}">
        <div class="k">${esc(it.label)}</div>
        <div class="sum-main">
          <div class="v" data-role="value">${esc(it.value) || '<span class="hint">미입력</span>'}</div>
          <div class="meta">${metaHtml(it)} · ${evidenceHtml(it)}</div>
          ${controlsHtml(it)}
        </div>
      </div>`).join('')}
    </div>` : '';
  const restHtml = rest.map((it) => `
    <div class="status-card" data-item-id="${it.id}" data-version="${it.versionNo}">
      <div class="k">${esc(it.label)}</div>
      <div class="v" data-role="value">${esc(it.value) || '<span class="hint">미입력</span>'}</div>
      <div class="meta">${metaHtml(it)}</div>
      ${evidenceHtml(it)}
      ${controlsHtml(it)}
    </div>`).join('');
  panel.innerHTML = (summaryHtml + (rest.length ? `<div class="k group-title">핵심 이슈 · 미결사항</div>${restHtml}` : ''))
    || '<p class="hint">등록된 현황 항목이 없습니다.</p>';

  panel.querySelectorAll('[data-edit]').forEach((btn) => btn.addEventListener('click', () => {
    const form = panel.querySelector(`[data-form="${btn.dataset.edit}"]`);
    form.style.display = form.style.display === 'none' ? 'block' : 'none';
  }));
  panel.querySelectorAll('[data-history]').forEach((btn) => btn.addEventListener('click', async () => {
    const data = await getJson(`/api/knowledge-items/${btn.dataset.history}/history`);
    showEvidencePane('<h3>버전 이력</h3>' + data.versions.map((v) => `
      <div class="status-card"><div class="meta">v${v.version_no} · ${esc(v.created_by)}, ${esc((v.created_at || '').slice(0, 16).replace('T', ' '))}${v.note ? ' · ' + esc(v.note) : ''}</div>
      <div class="v">${esc(v.value)}</div></div>`).join(''));
  }));
  panel.querySelectorAll('[data-save]').forEach((btn) => btn.addEventListener('click', async () => {
    const id = btn.dataset.save;
    const form = panel.querySelector(`[data-form="${id}"]`);
    const textarea = form.querySelector('textarea');
    const evidenceSelect = form.querySelector('[data-evidence]');
    const r = await postJson(`/api/knowledge-items/${id}/edit`, {
      baseVersionNo: Number(btn.dataset.version), value: textarea.value, actor: '담당자(시연)', note: '화면에서 직접 수정',
      evidenceDocVersionId: evidenceSelect?.value ? Number(evidenceSelect.value) : null,
    });
    if (!r.ok) {
      alert(r.reason === 'CONFLICT' ? `다른 곳에서 이미 수정되었습니다(현재 버전 ${r.currentNo}). 새로고침 후 다시 시도하세요.` : '저장 실패: ' + r.reason);
      return;
    }
    await renderStatusPanel();
  }));
  panel.querySelectorAll('[data-view]').forEach((chip) => chip.addEventListener('click', () => openViewer(chip.dataset.view, chip.dataset.segment || null)));
}

function statusPillHtml(v) {
  const label = { ready: '추출 완료', partial: '일부 추출', conversion_needed: '변환 필요', unsupported: '미지원', failed: '실패', pending: '대기', extracting: '추출 중' }[v] || v;
  return `<span class="status-pill ${v}">${label}</span>`;
}

/* B3-3 — 본문이 추출되지 않은 자료가 "정상 수집됨"처럼 보이면 안 된다. 상태 표시(색·짧은
   라벨)만으로는 "검색되지 않는다"는 사실이 안 드러나므로 한 줄로 명시한다. */
const EXTRACT_NOTE = {
  conversion_needed: '본문이 추출되지 않았습니다 — 검색·근거 인용 대상이 아니며 원본 보관·다운로드만 됩니다.',
  unsupported: '본문이 추출되지 않았습니다 — 검색·근거 인용 대상이 아니며 원본 보관·다운로드만 됩니다.',
  failed: '추출에 실패했습니다 — 본문이 검색·근거 인용 대상이 아닙니다.',
};

/* 마감 A3 — 본문이 하나도 추출되지 않은 자료에서 「원문 보기」를 누르면 "추출된 본문이
   없습니다"만 뜬다(= 반드시 실패하는 버튼). 이 상태에서는 버튼을 비활성으로 두고 왜
   못 보는지를 글자로 적는다. 실제로 보관된 원본의 다운로드는 그대로 남긴다. */
const NO_BODY_STATUS = new Set(['conversion_needed', 'unsupported', 'failed']);

async function renderDocumentsPanel() {
  const panel = document.getElementById('panel-documents');
  const q = state.scope === 'team' ? 'scope=team' : `projectId=${state.projectId}`;
  const data = await getJson(`/api/documents?${q}`);
  const memoFormHtml = state.scope === 'team' ? '' : `<form class="inline-form" id="memoForm">
    <label>텍스트 회의 메모 입력(파일 없이 바로 등록)</label>
    <input name="title" placeholder="제목(선택)">
    <textarea name="text" rows="3" placeholder="회의 내용을 붙여넣거나 입력하세요" required></textarea>
    <button type="submit">메모 등록</button>
  </form>`;
  const docsHtml = data.documents.map((d) => {
    const latest = d.versions[0];
    const noBody = latest && NO_BODY_STATUS.has(latest.extract_status);
    return `<div class="doc-row">
      <div>
        <div class="doc-title">${esc(d.title)}</div>
        <div class="doc-meta">${latest
          /* 마감 — "현재 버전"이 아니라 "최신 수집 버전"이다(체험판 예시 자료).
             그 버전이 실제로 수집된 시각을 분 단위로 같이 적는다. */
          ? `${statusPillHtml(latest.extract_status)} · 최신 수집 버전 v${latest.version_no} · 수집 시각 ${esc((latest.uploaded_at || '').slice(0, 16).replace('T', ' '))}${noBody ? '<div>본문 추출 미지원 — 원문 보기를 제공하지 않습니다(원본 다운로드는 가능).</div>' : ''}${EXTRACT_NOTE[latest.extract_status] ? `<div>${EXTRACT_NOTE[latest.extract_status]}</div>` : ''}${latest.extract_warnings ? `<div>${esc(latest.extract_warnings.split('\n')[0])}</div>` : ''}`
          /* A5 — 버전이 하나도 없는 행 = 사전검사에 거절돼 등록되지 못한 자료(옛 코드가
             남긴 잔재). 무엇인지 알려주고 사용자가 직접 정리할 수 있게 한다. */
          : '<span class="status-pill failed">수집 실패</span> · 등록된 버전 없음 — 체험판 예시 자료입니다'}</div>
      </div>
      <div class="doc-actions">
        ${latest
          ? `${noBody
              ? '<button disabled title="본문이 추출되지 않아 원문 보기를 열 수 없습니다. 원본을 내려받아 확인하세요.">본문 추출 미지원</button>'
              : `<button data-view="${latest.id}">원문 보기</button>`}${d.kind === 'text_memo' ? '' : `<button data-download="${latest.id}">다운로드</button>`}`
          : `<button data-discard="${d.id}">목록에서 정리</button>`}
      </div>
    </div>`;
  }).join('');
  /* 마감 — 원본 삭제는 감지하지 않는다(동기화 엔진을 넓히지 않았다). 목록이 무엇을
     보장하고 무엇을 보장하지 않는지 화면에 그대로 적는다. */
  const collectedNote = `<p class="hint">이 목록은 <b>마지막 수집 시각</b> 기준입니다${data.lastSyncFinishedAt
    ? ` — 마지막 예시 갱신 체험: ${esc(data.lastSyncFinishedAt.slice(0, 16).replace('T', ' '))}`
    : ' — 아직 예시 갱신 체험을 실행한 적이 없습니다'}. 체험판에는 실제 공유폴더가 없습니다 — 모두 미리 준비한 가상 자료입니다.</p>`;
  panel.innerHTML = memoFormHtml + collectedNote + (docsHtml || '<p class="hint">등록된 자료가 없습니다. 「자료 갱신」 탭에서 예시 갱신을 체험하거나 위에서 메모를 등록하세요.</p>');

  panel.querySelectorAll('[data-view]').forEach((btn) => btn.addEventListener('click', () => openViewer(btn.dataset.view)));
  panel.querySelectorAll('[data-download]').forEach((btn) => btn.addEventListener('click', () => {
    window.location.href = `/api/doc-versions/${btn.dataset.download}/download`;
  }));
  panel.querySelectorAll('[data-discard]').forEach((btn) => btn.addEventListener('click', async () => {
    if (!confirm('등록된 버전이 없는 자료입니다. 목록에서 정리할까요?')) return;
    const r = await postJson(`/api/documents/${btn.dataset.discard}/discard-empty`, {});
    if (!r.ok) { alert('정리 실패: ' + (r.reason === 'HAS_VERSIONS' ? '이 자료에는 등록된 버전이 있어 정리하지 않습니다.' : r.reason)); return; }
    await renderDocumentsPanel();
  }));
  const memoForm = document.getElementById('memoForm');
  if (memoForm) memoForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const f = new FormData(e.target);
    await postJson('/api/documents/text-memo', {
      projectId: state.projectId, title: f.get('title'), text: f.get('text'), actor: '담당자(시연)',
    });
    await renderDocumentsPanel();
  });
}

/* B3-1 — 근거 원문은 화면을 덮는 모달이 아니라 오른쪽 상시 패널에 연다. 검색 결과나
   근거 칩을 눌러도 중앙의 결과·현황이 가려지지 않아 "검색 → 결과 → 근거 원문"이 한
   화면에 같이 남는다. #viewerBody는 아무것도 안 열렸을 때 비어 있게 두고(안내 문구는
   형제 요소 #viewerEmpty가 맡는다), 내용이 들어올 때만 채운다. */
function showEvidencePane(html) {
  /* 마감(B) — 접어 둔 상태에서 근거를 열면 자동으로 다시 펼친다(누른 결과가 안 보이는
     일이 없게). 접기는 "선택 전 빈 패널"을 치우려는 것이지 기능을 숨기는 게 아니다. */
  setEvidenceCollapsed(false);
  document.getElementById('viewerEmpty').style.display = 'none';
  document.getElementById('viewerBody').innerHTML = html;
}
function clearEvidencePane() {
  document.getElementById('viewerBody').innerHTML = '';
  document.getElementById('viewerEmpty').style.display = '';
}

async function openViewer(docVersionId, segmentId) {
  const body = document.getElementById('viewerBody');
  showEvidencePane('<p class="hint">불러오는 중…</p>');
  const [meta, view] = await Promise.all([
    getJson(`/api/doc-versions/${docVersionId}`),
    getJson(`/api/doc-versions/${docVersionId}/view-html`),
  ]);
  /* A3 — 어느 버전을 보고 있는지 제목에 명시한다(검색에서 과거 버전 결과를 눌러 들어온
     경우, 최신 원문과 헷갈리지 않게). */
  let html = `<h3>${esc(meta.docVersion.original_filename)} <span class="hint">v${meta.docVersion.version_no} · 수정일 ${esc((meta.docVersion.uploaded_at || '').slice(0, 10))}</span></h3><p class="hint">상태: ${statusPillHtml(meta.docVersion.extract_status)} · 수집 시각 ${esc((meta.docVersion.uploaded_at || '').slice(0, 16).replace('T', ' '))}(체험판 — 미리 준비한 예시 본문입니다)</p>`;
  if (meta.docVersion.extract_warnings) html += `<p class="hint">${esc(meta.docVersion.extract_warnings).replace(/\n/g, '<br>')}</p>`;
  if (view.html) {
    html += view.html;
  } else if (meta.segments.some((s) => s.kind === 'excel_cell' || s.kind === 'excel_range')) {
    /* 값 하나만 뚝 떼어 보여주면 "무슨 항목의 금액인지" 문맥이 안 보인다 — 실제로
       추출된 시트·셀 주소만으로(단위·통화를 새로 만들어 넣지 않고) 표 형태로 재배치해
       같은 행·열의 제목이 값과 함께 보이게 한다. */
    html += '<hr>' + renderExcelSegmentsAsGrid(meta.segments);
  } else if (meta.segments.length) {
    /* 단계 2 — 근거 클릭이 "그 문서"가 아니라 "그 구간"으로 연결되게 각 블록에
       구간 id를 남겨둔다(아래에서 특정 구간으로 스크롤·강조). */
    html += '<hr>' + meta.segments.map((s) => `<p data-segment-id="${s.id}"><b>${esc(s.page_no ? '페이지 ' + s.page_no : s.block_ref) || ''}</b><br>${esc(s.text)}</p>`).join('');
  } else {
    html += '<p class="hint">추출된 본문이 없습니다. 원본 파일을 다운로드해 확인하세요.</p>';
  }
  body.innerHTML = html;
  if (segmentId) {
    /* view.html(Word 원문) 경로는 구간을 통째 문자열로 안 보여주고 mammoth 원본 html을
       그대로 쓰므로 data-segment-id가 없다 — 대신 extract/word.mjs가 심어 둔
       data-block-ref로 찾는다(그 구간의 block_ref는 meta.segments에서 역으로 찾는다). */
    let target = body.querySelector(`[data-segment-id="${segmentId}"]`);
    if (!target) {
      const seg = meta.segments.find((s) => String(s.id) === String(segmentId));
      if (seg?.blockRef || seg?.block_ref) target = body.querySelector(`[data-block-ref="${seg.blockRef || seg.block_ref}"]`);
    }
    if (target) { target.classList.add('segment-highlight'); target.scrollIntoView({ block: 'center' }); }
  }
}

document.getElementById('viewerClose').addEventListener('click', clearEvidencePane);

/* 마감(B) — 근거 패널 구조는 그대로 두고(오른쪽 상시 패널), 좁은 화면에서 아무것도
   안 열린 빈 패널이 자리를 과하게 차지할 때만 접을 수 있게 한다. */
function setEvidenceCollapsed(collapsed) {
  document.getElementById('layout').classList.toggle('evidence-collapsed', collapsed);
  document.getElementById('viewerCollapse').textContent = collapsed ? '펼치기' : '접기';
}
document.getElementById('viewerCollapse').addEventListener('click', () => {
  setEvidenceCollapsed(!document.getElementById('layout').classList.contains('evidence-collapsed'));
});

async function renderActivitiesPanel() {
  const panel = document.getElementById('panel-activities');
  const data = await getJson(`/api/activities?projectId=${state.projectId}`);
  const formHtml = `<form class="inline-form" id="activityForm">
    <label>발생일 <input type="date" name="occurredAt" required></label>
    <label>구분 <select name="category"><option>회의</option><option>협의</option><option>보고</option><option>의사결정</option><option>이행</option></select></label>
    <label>발생 사실/논의 <textarea name="fact" rows="2"></textarea></label>
    <label>결정/의견 <textarea name="decision" rows="2"></textarea></label>
    <label>이유 <input name="reason"></label>
    <label>다음 조치 <input name="nextAction"></label>
    <button type="submit">이력 등록</button>
  </form>`;
  const listHtml = data.activities.map((a) => `
    <div class="activity-row">
      <span class="cat">${esc(a.category)}</span> <b>${esc(a.occurred_at)}</b> · ${esc(a.author)}
      ${a.fact ? `<div>${esc(a.fact)}</div>` : ''}
      ${a.decision ? `<div>${esc(a.decision)}</div>` : ''}
      ${a.reason ? `<div class="hint">이유: ${esc(a.reason)}</div>` : ''}
      ${a.next_action ? `<div class="hint">다음 조치: ${esc(a.next_action)}</div>` : ''}
    </div>`).join('') || '<p class="hint">등록된 이력이 없습니다.</p>';
  panel.innerHTML = formHtml + listHtml;
  document.getElementById('activityForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const f = new FormData(e.target);
    await postJson('/api/activities', {
      projectId: state.projectId, occurredAt: f.get('occurredAt'), author: '담당자(시연)', category: f.get('category'),
      fact: f.get('fact'), decision: f.get('decision'), reason: f.get('reason'), nextAction: f.get('nextAction'),
    });
    await renderActivitiesPanel();
  });
}

/* R03: 갱신 결과 문구가 renderSyncPanel() 재호출로 사라지던 문제 — panel.innerHTML을
   통째로 새로 그릴 때마다 이 값을 같이 넣어서, 다시 그려도 마지막 결과가 남아있게 한다. */
let lastSyncMessage = '';

/* B3-3 — 지원 형식 판정은 서버(추출 dispatch 옆)에서 받아 그대로 적는다. 화면이
   목록을 따로 들고 있지 않으므로 추출기를 고쳤을 때 조용히 어긋나지 않는다. */
const FORMAT_LEVEL_LABEL = { supported: '지원', unsupported: '미지원', unverified: '미확인' };

/* 마감(B) — 기본 노출은 「형식 / 가능한 작업 / 제한」(과 한 단어짜리 구분 배지)만 둔다.
   파일 경로·함수명·테스트명 같은 판정 근거는 지우지 않고 「기술 확인 내역」으로 접는다 —
   예전엔 이 긴 설명이 표를 세로로 늘려 최근 갱신 결과를 화면 밖으로 밀어냈다. */
function formatsCardHtml(formats) {
  return `<div class="format-card">
    <p class="hint">「지원」은 본문 추출까지 되는 형식입니다. 「미지원」은 원본 보관·다운로드만 되고 본문이 검색되지 않습니다.
    「미확인」은 이 앱에서 실측하지 않아 어느 쪽이라고 적을 근거가 없는 경우입니다 — 추측해서 적지 않습니다.</p>
    <table><thead><tr><th>형식</th><th>구분</th><th>가능한 작업</th><th>제한</th></tr></thead><tbody>
    ${formats.map((f) => `<tr>
      <td class="ext">${esc(f.ext)}</td>
      <td><span class="fmt-badge ${f.level}">${FORMAT_LEVEL_LABEL[f.level] || esc(f.level)}</span></td>
      <td>${esc(f.can)}</td>
      <td>${esc(f.limit)}</td>
    </tr>`).join('')}
    </tbody></table>
    <details class="fmt-basis"><summary>기술 확인 내역(판정 근거 — 파일·함수·테스트)</summary>
      ${formats.map((f) => `<p class="hint"><b>${esc(f.ext)}</b> — ${esc(f.basis)}</p>`).join('')}
    </details></div>`;
}

/* 마감 A1 — 갱신 대상 연결. 예전엔 sourceConnectionId=1이 화면에 박혀 있어 어느 사업장을
   골라도 같은 목록이 보였다. 실제 연결 구조를 확인해 보니 이 앱의 연결 단위는 "사업장"이
   아니라 "공유폴더 하나"이고(sync.mjs가 연결 루트에서 펀드→사업장으로 전부 훑는다),
   데모 연결도 펀드A·펀드B와 그 아래 사업장 전부를 한 연결이 덮는다 — 그래서 ID가 1인
   것 자체는 오류가 아니다. 고칠 것은 "무엇이 갱신되는지 안 밝히는 것"이므로, 대상 이름과
   덮는 범위를 적고 전역 동작임을 분명히 한다. 연결이 둘 이상이면 고를 수 있게 한다. */
let syncConnectionId = null;

async function renderSyncPanel() {
  const panel = document.getElementById('panel-sync');
  const conns = (await getJson('/api/source-connections')).connections;
  if (!conns.length) {
    panel.innerHTML = '<p class="hint">연결된 공유폴더가 없습니다.</p>';
    return;
  }
  if (!conns.some((c) => c.id === syncConnectionId)) syncConnectionId = conns[0].id;
  const conn = conns.find((c) => c.id === syncConnectionId);
  const [mappings, runs, tree, failures, formats] = await Promise.all([
    getJson(`/api/folder-mappings?sourceConnectionId=${conn.id}`),
    getJson(`/api/sync/runs?sourceConnectionId=${conn.id}`),
    getJson('/api/tree'),
    getJson(`/api/sync/failures?sourceConnectionId=${conn.id}`),
    getJson('/api/supported-formats'),
  ]);
  const candidates = mappings.mappings.filter((m) => m.review_status === 'candidate');
  const confirmed = mappings.mappings.filter((m) => m.review_status === 'confirmed');
  /* A4 — 이미 잘못 분류돼 수집이 안 되고 있는 폴더(복구 대상). 서버가 폴더마다 실제로
     수집되는 분류를 알려주므로 화면은 그대로 보여주고 재분류 경로만 제공한다. */
  const unsupported = confirmed.filter((m) => m.unsupported);

  /* A4 — 지원하지 않는 조합은 애초에 고를 수 없게 한다. 펀드 폴더(깊이 1)는 펀드
     공통자료로만, 사업장 폴더(깊이 2)는 사업장으로만 연결된다 — 그 밖의 조합은 예전에
     "확인"은 되지만 파일이 영원히 수집되지 않았다. */
  const targetOptionsFor = (allowedKind) => {
    if (allowedKind === 'fund') return tree.funds.map((f) => `<option value="fund:${f.id}">${esc(f.name)}(펀드 공통자료로)</option>`).join('');
    if (allowedKind === 'project') return tree.funds.map((f) => `<optgroup label="${esc(f.name)}">
      ${f.projects.map((p) => `<option value="project:${p.id}">${esc(p.name)}</option>`).join('')}
    </optgroup>`).join('');
    return '';
  };
  const ALLOWED_KIND_HINT = {
    fund: '펀드 폴더 — 펀드 공통자료로만 연결할 수 있습니다(직속 파일만 수집).',
    project: '사업장 폴더 — 사업장으로만 연결할 수 있습니다(하위 폴더까지 수집).',
  };
  /* 4차 검토 — "이동"은 사업장(project) 폴더가 이름을 바꾼 경우를 뜻한다. 펀드 단위
     확인된 매핑(project_id 없음, 공통자료용)까지 여기 나열되면, 그걸 "이동 전 폴더"로
     고를 수 있어 실제로는 무관한 사업장 폴더가 펀드 폴더 아래로 편입되며 문서 경로가
     엉키고 다음 갱신에서 중복 문서가 생기는 게 실측됐다 — 사업장 단위 확인된 매핑만 보여준다. */
  const moveFromOptions = confirmed.filter((m) => m.project_id)
    .map((m) => `<option value="${esc(m.relative_path)}">${esc(m.relative_path)}</option>`).join('');

  /* 마감 A1 — 대상 카드. 선택한 사업장만 갱신하는 것처럼 오해하지 않도록, 이 갱신이
     "연결된 공유폴더 전체"를 대상으로 한다는 것과 그 대상 이름·범위를 먼저 적는다. */
  /* 체험판 — 실제 공유폴더에 연결되어 있지 않다. 준비된 예시 변경(v1→v2)을
     화면에 반영하는 것뿐이므로, 「수집했다」로 읽히지 않게 적는다. */
  const targetCardHtml = `
    <div class="sync-target">
      <div class="k">예시 문서 갱신 체험</div>
      <div class="v">${esc(conn.name)}</div>
      <div class="meta">${esc(conn.root_path)}</div>
      <p class="hint">체험 범위: <b>${esc(conn.fund_names.join(', ') || '(없음)')}</b> / <b>${esc(conn.project_names.join(', ') || '(없음)')}</b></p>
      <p class="hint"><b>실제 파일을 읽거나 AI로 분석하지 않습니다.</b> 미리 준비해 둔 문서 개정본(v1 → v2)을 화면에 반영하고, 그 v2를 근거로 하는 변경 제안을 만들어 「근거 확인 → 승인」 흐름을 체험하게 합니다.</p>
    </div>`;

  panel.innerHTML = `
    ${targetCardHtml}
    <button id="syncBtn" ${runs.running ? 'disabled' : ''}>${runs.running ? '반영 중…' : '예시 문서 갱신 체험 실행'}</button>
    <div id="syncResult">${lastSyncMessage}</div>
    <h4>최근 갱신 기록</h4>
    <div id="runs">${runs.runs.map((r) => `
      <div class="run">${esc(r.started_at?.slice(0, 16).replace('T', ' '))} — ${esc({ success: '성공', partial: '일부 성공', failed: '실패', running: '진행 중' }[r.status] || r.status)} (전체 ${r.items_total}, 정상 ${r.items_ok}, 재시도 ${r.items_failed})</div>`).join('') || '<p class="hint">아직 실행 기록이 없습니다.</p>'}</div>
    ${failures.failures.length ? `<h4>수집 실패 자료(최근 갱신)</h4>
    <div id="syncFailures">${failures.failures.map((f) => `
      <div class="candidate-row" style="flex-direction:column;align-items:stretch;gap:2px">
        <div><b>${esc(f.relative_path)}</b></div>
        <div class="hint">${esc(f.error || '사유 기록 없음')}</div>
      </div>`).join('')}<p class="hint">원본을 고친 뒤 "지금 자료 갱신"을 다시 누르면 재시도합니다. 원본을 지우면 다음 갱신에서 목록에서 사라집니다.</p></div>` : ''}
    <h4>미확인 폴더(체험판에는 없음)</h4>
    <div id="candidates">${candidates.map((c) => `
      <div class="candidate-row" style="flex-direction:column;align-items:stretch;gap:6px">
        <div><b>${esc(c.relative_path)}</b> (미확인)</div>
        <div class="hint">${esc(ALLOWED_KIND_HINT[c.allowed_kind] || '이 깊이의 폴더는 아직 수집을 지원하지 않습니다 — 펀드 폴더 또는 그 아래 사업장 폴더만 수집합니다.')}</div>
        ${c.allowed_kind ? `<div style="display:flex;gap:6px;flex-wrap:wrap">
          <select data-assign-target="${c.id}"><option value="">${c.allowed_kind === 'fund' ? '펀드' : '사업장'} 선택…</option>${targetOptionsFor(c.allowed_kind)}</select>
          <button data-confirm-new="${c.id}">이 폴더로 확인</button>
        </div>` : ''}
        ${moveFromOptions && c.allowed_kind === 'project' ? `<div style="display:flex;gap:6px;flex-wrap:wrap">
          <select data-move-from="${c.id}"><option value="">— 또는: 기존 폴더의 이동…</option>${moveFromOptions}</select>
          <button data-confirm-move="${c.id}">이동으로 확인</button>
        </div>` : ''}
      </div>`).join('') || '<p class="hint">없음</p>'}</div>
    ${unsupported.length ? `<h4>재분류가 필요한 폴더</h4>
    <div id="unsupportedMappings">${unsupported.map((m) => `
      <div class="candidate-row" style="flex-direction:column;align-items:stretch;gap:6px">
        <div><b>${esc(m.relative_path)}</b> — 지금 분류로는 자료가 수집되지 않습니다.</div>
        <div class="hint">${esc(ALLOWED_KIND_HINT[m.allowed_kind] || '이 깊이의 폴더는 아직 수집을 지원하지 않습니다.')}</div>
        ${m.allowed_kind ? `<div style="display:flex;gap:6px;flex-wrap:wrap">
          <select data-assign-target="${m.id}"><option value="">${m.allowed_kind === 'fund' ? '펀드' : '사업장'} 선택…</option>${targetOptionsFor(m.allowed_kind)}</select>
          <button data-confirm-new="${m.id}">재분류</button>
        </div>` : ''}
      </div>`).join('')}</div>` : ''}
    <h4>지원 파일 형식</h4>
    ${formatsCardHtml(formats.formats)}`;

  const connSelect = document.getElementById('syncConnSelect');
  if (connSelect) connSelect.addEventListener('change', async () => {
    syncConnectionId = Number(connSelect.value);
    lastSyncMessage = '';
    await renderSyncPanel();
  });

  const syncBtn = document.getElementById('syncBtn');
  if (syncBtn) syncBtn.addEventListener('click', async () => {
    lastSyncMessage = '<p class="hint">반영 중…</p>';
    await renderSyncPanel();
    const r = await postJson('/api/sync/run', { sourceConnectionId: conn.id, actor: '담당자(체험)' });
    lastSyncMessage = r.ok
      ? (r.addedVersions
        ? `<p class="hint">예시 개정본을 반영했습니다(${new Date().toLocaleTimeString('ko-KR')}) — 문서 ${r.addedVersions}건이 v2가 되었고, 그 v2를 근거로 한 변경 제안 ${r.addedProposals}건이 오른쪽 「변경 검토」에 생겼습니다. <b>실제 파일 수집·AI 분석은 하지 않았습니다.</b></p>`
        : `<p class="hint">이미 반영되어 있습니다 — 준비된 예시 개정본은 하나뿐입니다. 처음부터 다시 보려면 상단 「체험 초기화」를 누르세요.</p>`)
      : `<p class="hint">오류: ${esc(r.reason)}</p>`;
    await renderSyncPanel();
    await renderDocumentsPanel();
    /* 현황 탭의 "수정" 폼은 근거로 고를 문서 목록을 renderStatusPanel()에서 만든다 —
       갱신으로 새로 들어온 자료가 그 목록에 반영되려면 현황도 같이 다시 그려야 한다
       (안 그리면 갱신 직후 "근거 자료 선택" 드롭다운이 비어 있거나 아예 없다). */
    await renderStatusPanel();
    /* 체험판 — 예시 갱신은 그 v2 를 근거로 한 변경 제안도 함께 만든다.
       「변경 검토」를 다시 그리지 않으면 방금 생긴 제안이 안 보인다. */
    await renderReviewPanel();
  });

  panel.querySelectorAll('[data-confirm-new]').forEach((btn) => btn.addEventListener('click', async () => {
    const id = btn.dataset.confirmNew;
    const sel = panel.querySelector(`[data-assign-target="${id}"]`);
    if (!sel.value) { alert('사업장 또는 펀드를 먼저 선택하세요.'); return; }
    const [kind, targetId] = sel.value.split(':');
    const r = await postJson(`/api/folder-mappings/${id}/confirm`, kind === 'fund' ? { fundId: Number(targetId) } : { projectId: Number(targetId) });
    /* A4 — 서버가 거절하면 사유와 가능한 분류를 그대로 보여준다(조용히 사라지지 않는다). */
    lastSyncMessage = r.ok
      ? '<p class="hint">폴더 연결을 확인했습니다. "지금 자료 갱신"을 눌러 자료를 가져오세요.</p>'
      : `<p class="hint">확인 실패: ${esc(r.message || r.reason)}</p>`;
    await renderSyncPanel();
  }));
  panel.querySelectorAll('[data-confirm-move]').forEach((btn) => btn.addEventListener('click', async () => {
    const id = btn.dataset.confirmMove;
    const sel = panel.querySelector(`[data-move-from="${id}"]`);
    if (!sel.value) { alert('이동 전 폴더를 먼저 선택하세요.'); return; }
    const r = await postJson(`/api/folder-mappings/${id}/confirm-move`, { fromRelativePath: sel.value });
    lastSyncMessage = r.ok
      ? `<p class="hint">이동을 확인했습니다(문서 ${r.movedCount}건 이어짐). "지금 자료 갱신"을 눌러 확인하세요.</p>`
      : `<p class="hint">이동 확인 실패: ${esc(r.reason)}${r.conflicts ? ' — ' + esc(JSON.stringify(r.conflicts)) : ''}</p>`;
    await renderSyncPanel();
    await renderDocumentsPanel();
  }));
}

/* 단계 2 — 변경 검토 패널(§7). 선택 사업장의 대기·승인·반려·충돌 제안을 전부 보여준다.
   현재값/기준값/제안값을 나란히 보여주고, 근거는 문서 버전+구간으로 클릭 연결한다. */
const PROPOSAL_STATUS_LABEL = { pending: '대기', approved: '승인됨', rejected: '반려됨', conflict: '충돌' };
const PROPOSAL_SOURCE_LABEL = { synthetic: '테스트 제안 — AI 미연결', manual: '수동 제안' };

async function renderReviewPanel() {
  const panel = document.getElementById('side-review');
  const data = await getJson(`/api/proposals?projectId=${state.projectId}`);
  const proposals = data.proposals;
  const pendingCount = proposals.filter((p) => p.status === 'pending').length;

  panel.innerHTML = `
    <button id="proposeSyntheticBtn">시연용 제안 생성 (테스트 제안 — AI 미연결)</button>
    <p class="hint">대기 중인 제안 ${pendingCount}건 / 전체 ${proposals.length}건</p>
    <div id="proposalList">${proposals.map((p) => `
      <div class="proposal-card" data-proposal-id="${p.id}">
        <div><span class="proposal-status ${p.status}">${PROPOSAL_STATUS_LABEL[p.status]}</span>
          <b>${esc(FIELD_LABELS_CLIENT[p.key] || p.label)}</b> · ${esc(PROPOSAL_SOURCE_LABEL[p.source_type] || p.source_type)}</div>
        <div class="diff">
          <div class="before">현재(v${p.current_version_no ?? '-'})<br>${esc(p.current_value) || '(없음)'}</div>
          <div class="after">제안(기준 v${p.base_version_no})<br>${esc(p.proposed_value) || '(없음)'}</div>
        </div>
        <div class="hint">이유: ${esc(p.reason)}</div>
        ${p.evidence_doc_version_id
          ? `<span class="evidence-chip" data-view="${p.evidence_doc_version_id}" data-segment="${p.evidence_segment_id ?? ''}">근거: ${esc(p.evidence_filename)}${p.evidence_text ? ' — ' + esc(p.evidence_text.slice(0, 30)) + '…' : ''}</span>`
          : '<span class="hint">연결된 근거 없음</span>'}
        <div class="hint">제안: ${esc(p.proposed_by)}, ${esc((p.proposed_at || '').slice(0, 16).replace('T', ' '))}</div>
        ${p.status === 'pending' ? `<div style="margin-top:8px"><button class="edit-btn" data-approve="${p.id}">승인</button> <button class="edit-btn" data-reject="${p.id}">반려</button></div>` : ''}
        ${p.status === 'conflict' ? `<p class="hint">충돌 사유: ${esc(p.decision_note)}</p>` : ''}
        ${p.status === 'rejected' ? `<p class="hint">반려 사유: ${esc(p.decision_note)} (검토: ${esc(p.decided_by)})</p>` : ''}
        ${p.status === 'approved' ? `<p class="hint">승인: ${esc(p.decided_by)}, ${esc((p.decided_at || '').slice(0, 16).replace('T', ' '))}</p>` : ''}
      </div>`).join('') || '<p class="hint">제안이 없습니다.</p>'}</div>`;

  document.getElementById('proposeSyntheticBtn').addEventListener('click', async () => {
    const statusData = await getJson(`/api/projects/${state.projectId}`);
    const target = statusData.status.find((s) => s.key === 'unresolved') || statusData.status[0];
    if (!target) { alert('제안을 걸 현황 항목이 없습니다.'); return; }
    const docsData = await getJson(`/api/documents?projectId=${state.projectId}`);
    const meetingDoc = docsData.documents.find((d) => d.title.includes('회의록') && d.versions[0]);
    let evidenceDocVersionId = null, evidenceSegmentId = null, proposedValue = `${target.value || ''}(테스트 제안으로 조정 검토)`.trim();
    if (meetingDoc) {
      evidenceDocVersionId = meetingDoc.versions[0].id;
      const dv = await getJson(`/api/doc-versions/${evidenceDocVersionId}`);
      const seg = dv.segments.find((s) => /\d+월 \d+일/.test(s.text)) || dv.segments[0];
      evidenceSegmentId = seg?.id ?? null;
      if (seg) proposedValue = `회신 예정일 조정 — 협의 완료(근거: ${meetingDoc.title})`;
    }
    const r = await postJson('/api/proposals', {
      itemId: target.id, proposedValue, reason: '테스트 제안 — AI 미연결. 신규 회의록 근거로 조정을 제안합니다(합성 시나리오).',
      sourceType: 'synthetic', proposedBy: '담당자(시연)', evidenceDocVersionId, evidenceSegmentId,
    });
    if (!r.ok) { alert('제안 생성 실패: ' + r.reason); return; }
    await renderReviewPanel();
  });
  panel.querySelectorAll('[data-approve]').forEach((btn) => btn.addEventListener('click', async () => {
    const r = await postJson(`/api/proposals/${btn.dataset.approve}/approve`, { approver: '팀장(시연)' });
    if (!r.ok) alert(r.reason === 'CONFLICT' ? `충돌: 기준 버전 이후 현황이 이미 바뀌었습니다(현재 v${r.currentNo}).` : '승인 실패: ' + r.reason);
    await renderReviewPanel();
    await renderStatusPanel();
  }));
  panel.querySelectorAll('[data-reject]').forEach((btn) => btn.addEventListener('click', async () => {
    const reason = prompt('반려 사유를 입력하세요');
    if (reason === null) return;
    const r = await postJson(`/api/proposals/${btn.dataset.reject}/reject`, { rejecter: '팀장(시연)', reason });
    if (!r.ok) alert('반려 실패: ' + r.reason);
    await renderReviewPanel();
  }));
  panel.querySelectorAll('[data-view]').forEach((chip) => chip.addEventListener('click', () => openViewer(chip.dataset.view, chip.dataset.segment || null)));
}

/* 단계 3 — AI 질문 패널. 브리프 §12-15: AI 연결 실패는 예시 답으로 몰래 대체하지 않고
   있는 그대로 알린다. 추출/검색/모델응답 오류를 서로 다른 문구로 구분해서 보여준다. */
const AI_ERROR_LABEL = {
  AI_NOT_CONNECTED: 'AI가 연결되지 않았습니다 — 서버에 ANTHROPIC_API_KEY가 설정돼 있지 않습니다.',
  RETRIEVAL_FAILED: '근거 자료 검색 중 오류가 발생했습니다.',
  AI_REQUEST_FAILED: 'AI 응답 요청이 실패했습니다(네트워크·인증·요청 한도 등). 잠시 후 다시 시도하세요.',
  EMPTY_QUESTION: '질문을 입력하세요.',
  INVALID_SCOPE: '범위 오류입니다.',
};

async function renderAiPanel() {
  const panel = document.getElementById('side-ai');
  const status = await getJson('/api/ai/status');
  /* B3-2 — 연결되지 않은 질의응답이 주 화면 자리를 차지하지 않게 보조 탭으로 내렸고,
     대신 탭 이름에 연결 상태를 그대로 적는다(기능을 숨기지는 않는다). */
  document.getElementById('aiSideTab').textContent = status.connected ? 'AI 질문' : 'AI 질문 · 미연결';
  if (!status.connected) {
    /* 마감(B) — 위치는 지금처럼 보조로 두되, 키 설정 절차는 일반 사용자 핵심 흐름에서
       분리한다(평소엔 접어 두고, 필요한 사람만 펼친다). */
    panel.innerHTML = `
      <p class="example-mode">AI 연결 안 됨 — 이 화면에서 모델을 호출할 수 없습니다</p>
      <p class="hint">자료를 찾는 일은 AI 연결과 무관합니다 — 가운데 위 검색창을 쓰세요. 예시 답으로 대신 채우지 않습니다.</p>
      <details><summary class="hint">AI 키 설정 절차(관리자용)</summary>
        <p class="hint">서버 환경변수 ANTHROPIC_API_KEY가 설정되지 않았습니다. 키를 설정하고 서버를 재시작하면 이 화면에서 바로 질문할 수 있습니다.</p>
      </details>`;
    return;
  }
  panel.innerHTML = `
    <form id="aiAskForm">
      <textarea id="aiQuestion" rows="3" style="width:100%;padding:6px" placeholder="현재 범위(${esc(state.scope === 'team' ? '팀 공통' : state.projectName)}) 자료에 대해 질문하세요"></textarea>
      <button type="submit" id="aiAskBtn">질문하기</button>
    </form>
    <div id="aiAnswer"></div>`;
  document.getElementById('aiAskForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const q = document.getElementById('aiQuestion').value.trim();
    const answerBox = document.getElementById('aiAnswer');
    if (!q) { answerBox.innerHTML = `<p class="hint">${AI_ERROR_LABEL.EMPTY_QUESTION}</p>`; return; }
    const btn = document.getElementById('aiAskBtn');
    btn.disabled = true; btn.textContent = '생각 중…';
    answerBox.innerHTML = '<p class="hint">답변을 생성하는 중…</p>';
    const r = await postJson('/api/ai/ask', { question: q, scope: state.scope, id: state.scope === 'team' ? null : state.projectId });
    btn.disabled = false; btn.textContent = '질문하기';
    if (!r.ok) {
      answerBox.innerHTML = `<p class="hint">${esc(AI_ERROR_LABEL[r.reason] || ('오류: ' + r.reason))}</p>`;
      return;
    }
    const citeHtml = r.citations.length
      ? r.citations.map((c) => `<span class="evidence-chip" data-view="${c.docVersionId}" data-segment="${c.segmentId}">근거: ${esc(c.title)} — ${esc(c.snippet)}…</span>`).join(' ')
      : '<span class="hint">인용된 근거 없음</span>';
    answerBox.innerHTML = `<div class="v" style="white-space:pre-wrap">${esc(r.answer)}</div><div style="margin-top:6px">${citeHtml}</div>`;
    answerBox.querySelectorAll('[data-view]').forEach((chip) => chip.addEventListener('click', () => openViewer(chip.dataset.view, chip.dataset.segment || null)));
  });
}

document.querySelectorAll('.tab').forEach((tab) => tab.addEventListener('click', () => {
  document.querySelectorAll('.tab').forEach((t) => t.classList.remove('active'));
  document.querySelectorAll('.panel').forEach((p) => p.classList.remove('active'));
  tab.classList.add('active');
  document.getElementById(`panel-${tab.dataset.tab}`).classList.add('active');
}));
document.querySelectorAll('.sidetab').forEach((tab) => tab.addEventListener('click', () => {
  document.querySelectorAll('.sidetab').forEach((t) => t.classList.remove('active'));
  document.querySelectorAll('.sidepanel').forEach((p) => p.classList.remove('active'));
  tab.classList.add('active');
  document.getElementById(`side-${tab.dataset.sidetab}`).classList.add('active');
}));
document.getElementById('teamLink').addEventListener('click', selectTeam);

/* 마감 A2 — 검색이 언제나 전체 범위였다(결과마다 소속은 찍혔지만 범위를 좁힐 수 없었다).
   이제 범위를 화면에 명시하고, 사업장을 고른 상태면 그 사업장이 기본이다 — 전체 검색은
   사용자가 명시적으로 고를 때만 한다. 「팀 공통 지식」을 고른 상태면 좁힌 범위가
   팀 공통이 된다. 기존 버전 필터(A3: 이전 버전 포함)와는 독립이라 함께 쓸 수 있다. */
function scopedSearchLabel() {
  if (state.scope === 'team') return '팀 공통 지식';
  return `현재 사업장 — ${state.fundName ? state.fundName + ' · ' : ''}${state.projectName}`;
}
function updateSearchScopeControl() {
  const sel = document.getElementById('searchScope');
  sel.options[0].textContent = scopedSearchLabel();
  sel.value = 'scoped'; // 대상을 바꾸면 기본값(그 대상으로 좁힌 검색)으로 돌아온다
  updateSearchScopeNote();
}
function updateSearchScopeNote() {
  const sel = document.getElementById('searchScope');
  document.getElementById('searchScopeNote').innerHTML = sel.value === 'all'
    ? '검색 범위: <b>전체</b> — 모든 펀드·사업장과 팀 공통 지식에서 찾습니다.'
    : (state.scope === 'team'
      ? '검색 범위: <b>팀 공통 지식</b>만 — 사업장 자료는 찾지 않습니다.'
      : `검색 범위: <b>${esc(state.fundName)} · ${esc(state.projectName)}</b>의 자료·현황만 — 다른 사업장과 팀 공통 지식은 찾지 않습니다.`);
}

let searchTimer;
function runSearch() {
  clearTimeout(searchTimer);
  const q = document.getElementById('searchBox').value.trim();
  const includeOld = document.getElementById('includeOldVersions').checked;
  const scopeChoice = document.getElementById('searchScope').value;
  updateSearchScopeNote();
  const box = document.getElementById('searchResults');
  if (!q) { box.innerHTML = ''; return; }
  const scopeQuery = scopeChoice === 'all' ? ''
    : (state.scope === 'team' ? '&scope=team' : `&scope=project&projectId=${state.projectId}`);
  searchTimer = setTimeout(async () => {
    const r = await getJson(`/api/search?q=${encodeURIComponent(q)}${includeOld ? '&includeOldVersions=1' : ''}${scopeQuery}`);
    /* 브리프 §7 "결과와 인용마다 펀드·사업장을 표시한다" — 어디 소속 결과인지 항상 같이 보여준다. */
    const scopeLabel = (scope, fundName, projectNames) => {
      if (scope === 'team') return '팀 공통';
      if (scope === 'fund') return esc(fundName || '펀드');
      /* 동명 사업장이 다른 펀드에 있을 수 있어(브리프 §7), 사업장명만으로는 구분이 안 된다
         — 펀드명을 같이 보여줘야 동명 사업장 두 결과를 화면에서 구분할 수 있다. */
      const proj = esc(projectNames || '(연결된 사업장 없음)');
      return fundName ? `${esc(fundName)} · ${proj}` : proj;
    };
    /* A3 — 과거 버전 결과는 버전 번호와 그 버전의 문서 수정일을 같이 보여준다(어느 시점
       내용인지 결과만 보고 알 수 있어야 한다). 결과를 누르면 그 결과가 가리키는 바로 그
       문서 버전(doc_version_id)의 해당 구간이 열린다 — 최신 원문으로 바뀌지 않는다. */
    /* 마감 — "현재 버전"은 공유폴더의 현재 상태를 뜻하는 것처럼 읽힌다. 이 앱이 보장하는
       것은 "마지막으로 수집된 버전"뿐이므로 표현을 그대로 바꾼다. */
    const versionLabel = (s) => (s.version_no === s.current_version_no
      ? ' <span class="res-badge cur">최신 수집 버전</span>'
      : ` <span class="res-badge old">(이전 버전 v${s.version_no} · 수정일 ${esc((s.uploaded_at || '').slice(0, 10))})</span>`);
    /* B3-1 — 결과 줄마다 상태를 색이 아니라 글자로 적는다. 문서는 등록된 버전이 하나도
       없으면 수집 실패(사전검사에 거절된 자료), 있으면 현재 버전 번호와 추출 상태를
       같이 보여준다. 현황 항목은 사람이 아직 승인·반려하지 않은 변경 제안이 있으면
       「확인 대기」로 구분한다. */
    const docLabel = (d) => (!d.version_count
      ? ' <span class="res-badge fail">수집 실패 — 등록된 버전 없음</span>'
      : ` <span class="res-badge cur">최신 수집 버전 v${d.latest_version_no}</span> ${statusPillHtml(d.latest_extract_status)}`);
    const itemLabel = (i) => (i.pending_proposals
      ? ` <span class="res-badge wait">확인 대기 — 변경 제안 ${i.pending_proposals}건</span>`
      : '');
    const rows = [
      ...r.documents.map((d) => `<div class="result">📄 <b>[${scopeLabel(d.scope, d.fund_name, d.project_names)}]</b> ${esc(d.title)}${docLabel(d)}</div>`),
      ...r.segments.map((s) => `<div class="result" data-view="${s.doc_version_id}" data-segment="${s.id}" style="cursor:pointer">🔎 <b>[${scopeLabel(s.scope, s.fund_name, s.project_names)}]</b> ${esc(s.title)}${versionLabel(s)}: ${esc(s.text).slice(0, 60)}…</div>`),
      ...r.knowledgeItems.map((i) => `<div class="result">📌 <b>[${scopeLabel(i.scope, i.fund_name, i.project_name)}]</b> ${esc(FIELD_LABELS_CLIENT[i.key] || i.key)}: ${esc(i.value).slice(0, 60)}${itemLabel(i)}</div>`),
    ];
    box.innerHTML = rows.join('') || '<div class="result hint">결과 없음</div>';
    box.querySelectorAll('[data-view]').forEach((el) => el.addEventListener('click', () => openViewer(el.dataset.view, el.dataset.segment || null)));
  }, 200);
}
document.getElementById('searchBox').addEventListener('input', runSearch);
document.getElementById('includeOldVersions').addEventListener('change', runSearch);
document.getElementById('searchScope').addEventListener('change', runSearch);

loadTree();
