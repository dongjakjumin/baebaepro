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
      /* 개편 — 클릭만 받는 div 였다. 키보드로도 사업장을 고를 수 있게 버튼으로 둔다
         (클래스·data 속성·이벤트는 그대로다). */
      const el = document.createElement('button');
      el.type = 'button';
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
  document.getElementById('crumb').textContent = `${fundName} / ${name}`;
  document.getElementById('pageTitle').textContent = '사업장 지식관리';
  updateSearchScopeControl();
  runSearch();
  await renderStatusPanel();
  await renderDocumentsPanel();
  await renderActivitiesPanel();
  await renderSyncPanel();
  await renderReviewPanel();
}

async function selectTeam() {
  state = { scope: 'team', projectId: null, projectName: '팀 공통 지식', fundName: '' };
  setActiveProjectRow(-1);
  document.getElementById('crumb').textContent = '팀 공통 지식';
  document.getElementById('pageTitle').textContent = '팀 공통 지식';
  updateSearchScopeControl();
  runSearch();
  await renderStatusPanel();
  await renderDocumentsPanel();
  document.getElementById('panel-activities').innerHTML = '<p class="hint">업무 이력은 사업장 단위입니다. 왼쪽에서 사업장을 고르세요.</p>';
  /* 마감 A1 — 갱신 단위는 사업장·펀드가 아니라 체험판에서는 준비된 예시 개정본을 반영하는 것뿐이다. */
  document.getElementById('panel-sync').innerHTML = '<p class="hint">「갱신 체험」은 왼쪽에서 사업장을 고른 뒤 실행합니다 — 실제 파일을 수집하지 않고 준비된 예시 개정본을 반영합니다.</p>';
  document.getElementById('side-review').innerHTML = '<p class="hint">변경 제안은 사업장 단위입니다. 왼쪽에서 사업장을 고르세요.</p>';
}

/* ── 현황 탭 (개편) ─────────────────────────────────────────
   예전엔 항목마다 비슷한 길이의 카드가 세로로 줄지어, 먼저 읽어야 할 핵심 이슈·
   미결사항과 확인해야 할 변경이 화면 아래로 밀렸다. 이제 12열 벤토로 크기를 차등해
   ① 사업장 요약 ② 확인할 변경 ③ 주요 이슈·미결사항 ④ 최근 자료 ⑤ 최근 업무 이력
   ⑥ 갱신 체험 순으로 둔다. 카드에 넣는 값은 모두 저장된 실제 값이고, 대기·승인 수도
   실제 제안 목록에서 센다(장식용 숫자를 만들지 않는다).
   요약에 못 담은 항목(확정사항 등)은 「전체 현황」 펼침에서 그대로 읽고 고칠 수 있다. */
const ISSUE_KEYS = ['issue', 'unresolved'];
const KIND_LABEL = { docx: 'DOCX', xlsx: 'XLSX', pdf: 'PDF', hwp: 'HWP', memo: '메모', text_memo: '메모' };
const dt16 = (s) => esc((s || '').slice(0, 16).replace('T', ' '));

async function renderStatusPanel() {
  const panel = document.getElementById('panel-status');
  let items, docs = [], proposals = [], activities = [], project = null, fundName = '', lastSync = null;
  if (state.scope === 'team') {
    items = (await getJson('/api/team-status')).status;
  } else {
    const [data, docsData, propData, actData] = await Promise.all([
      getJson(`/api/projects/${state.projectId}`),
      getJson(`/api/documents?projectId=${state.projectId}`),
      getJson(`/api/proposals?projectId=${state.projectId}`),
      getJson(`/api/activities?projectId=${state.projectId}`),
    ]);
    items = data.status;
    project = data.project || null;
    fundName = (data.funds && data.funds[0] ? data.funds[0].name : '') || state.fundName;
    docs = docsData.documents.filter((d) => d.versions[0]);
    lastSync = docsData.lastSyncFinishedAt;
    proposals = propData.proposals || [];
    activities = actData.activities || [];
  }
  items.sort((a, b) => FIELD_ORDER.indexOf(a.key) - FIELD_ORDER.indexOf(b.key));
  const docOptions = docs.map((d) => `<option value="${d.versions[0].id}">${esc(d.title)}</option>`).join('');

  const evidenceHtml = (it) => (it.evidence.length
    ? it.evidence.map((e) => `<button type="button" class="evidence-chip" data-view="${e.doc_version_id}" data-segment="${e.segment_id ?? ''}">근거: ${esc(e.original_filename)}${e.page_no ? ' p.' + e.page_no : ''}${e.cell_range ? ' ' + esc(e.cell_range) : ''}</button>`).join(' ')
    : '<span class="hint">연결된 근거 없음</span>');
  const metaHtml = (it) => `v${it.versionNo}${it.effectiveDate ? ' · 기준일 ' + esc(it.effectiveDate) : ''} · ${esc(it.updatedBy)}, ${dt16(it.updatedAt)}`;
  const controlsHtml = (it) => `
      <div class="row-actions"><button class="edit-btn" data-edit="${it.id}">수정</button> <button class="edit-btn" data-history="${it.id}">이력(v${it.versionNo})</button></div>
      <div class="edit-form" data-form="${it.id}" style="display:none;margin-top:8px">
        <textarea rows="2">${esc(it.value)}</textarea>
        ${docs.length ? `<select data-evidence="${it.id}" style="margin-top:6px"><option value="">근거 자료 선택 안 함</option>${docOptions}</select>` : ''}
        <button class="edit-btn" data-save="${it.id}" data-version="${it.versionNo}" style="margin-top:8px">저장</button>
      </div>`;
  /* 항목 블록의 클래스·data 속성은 예전 그대로다 — 수정·이력·근거 연결이 배치만 바뀐다. */
  const itemHtml = (it) => `
    <div class="status-card" data-item-id="${it.id}" data-version="${it.versionNo}">
      <div class="k">${esc(it.label)}</div>
      <div class="v" data-role="value">${esc(it.value) || '<span class="hint">미입력</span>'}</div>
      <div class="meta">${metaHtml(it)}</div>
      ${evidenceHtml(it)}
      ${controlsHtml(it)}
    </div>`;

  if (state.scope === 'team') {
    panel.innerHTML = `
      <section class="km-card">
        <div class="km-card-head"><h2>팀 공통 지식</h2><span class="km-card-note">사업장과 무관하게 팀 전체에 적용되는 항목</span></div>
        <div class="km-items">${items.map(itemHtml).join('') || '<p class="hint">등록된 항목이 없습니다.</p>'}</div>
      </section>`;
    wireStatusControls(panel);
    return;
  }

  /* 최근 업무 이력 — 실제 저장된 것만 모은다. 현황 항목의 버전 이력(수정·승인)과
     등록된 업무 이력을 시각 역순으로 섞고, 줄마다 무엇인지 글자로 구분해 적는다. */
  const hists = await Promise.all(items.map(async (it) => ({
    it, versions: (await getJson(`/api/knowledge-items/${it.id}/history`)).versions || [],
  })));
  const events = [];
  hists.forEach(({ it, versions }) => versions.forEach((v) => {
    if (v.version_no <= 1) return;                 /* 초기값은 «변경»이 아니다 */
    events.push({ at: v.created_at, kind: /승인/.test(v.note || '') ? '승인' : '수정',
      main: `${it.label} v${v.version_no}`, by: v.created_by, text: v.value, note: v.note });
  }));
  activities.forEach((a) => events.push({ at: a.occurred_at, kind: esc(a.category), main: '업무 이력',
    by: a.author, text: a.decision || a.fact || '', note: '' }));
  events.sort((x, y) => String(y.at || '').localeCompare(String(x.at || '')));

  const pending = proposals.filter((p) => p.status === 'pending');
  const newest = proposals[0] || null;
  const mgr = items.find((it) => it.key === 'manager');
  const stg = items.find((it) => it.key === 'stage');
  const lastTouched = items.slice().sort((a, b) => String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')))[0];
  const issues = items.filter((it) => ISSUE_KEYS.includes(it.key));
  const others = items.filter((it) => !ISSUE_KEYS.includes(it.key));
  const recentDocs = docs.slice()
    .sort((a, b) => String(b.versions[0].uploaded_at || '').localeCompare(String(a.versions[0].uploaded_at || '')))
    .slice(0, 4);

  panel.innerHTML = `
    <section class="km-card km-navy km-c8">
      <div class="km-card-head"><h2>사업장 요약</h2><span class="km-card-note">저장된 값만 표시</span></div>
      <div class="km-sum-name">${esc(project ? project.name : state.projectName)}</div>
      <div class="km-sum-sub">${esc(project && project.type ? project.type : '유형 미입력')} · ${esc(fundName || '펀드 미연결')}</div>
      <div class="km-sum-grid">
        <div><div class="km-lbl">진행 단계</div><div class="km-big">${stg && stg.value ? esc(stg.value) : '미입력'}</div></div>
        <div><div class="km-lbl">담당자</div><div class="km-small">${mgr && mgr.value ? esc(mgr.value) : '미입력'}</div></div>
        <div><div class="km-lbl">최근 수정</div><div class="km-small">${lastTouched ? dt16(lastTouched.updatedAt) + ' · ' + esc(lastTouched.updatedBy) : '기록 없음'}</div></div>
      </div>
      <div class="km-card-foot">
        ${stg ? `<button type="button" class="km-btn-onnavy" data-goto-item="${stg.id}">진행 단계 수정</button>` : ''}
        ${mgr ? `<button type="button" class="km-btn-onnavy" data-goto-item="${mgr.id}">담당자 수정</button>` : ''}
        <button type="button" class="km-btn-onnavy" data-open-all="1">전체 현황 열기</button>
      </div>
    </section>

    <section class="km-card km-c4">
      <div class="km-card-head"><h2>확인할 변경</h2></div>
      <div class="km-kpi"><span class="km-kpi-n">${pending.length}</span><span class="km-kpi-u">건 확인 대기</span></div>
      ${newest ? `
      <p class="km-oneline"><span class="km-lbl">가장 최근 제안 (${esc(PROPOSAL_STATUS_LABEL[newest.status] || newest.status)})</span>${esc(FIELD_LABELS_CLIENT[newest.key] || newest.label || '')} — ${esc(String(newest.proposed_value || '').slice(0, 44))}</p>
      <div class="km-row-meta"><span class="km-tag">${newest.evidence_doc_version_id ? '근거 연결됨' : '근거 없음'}</span><span>${dt16(newest.proposed_at)}</span></div>`
      : `<p class="km-oneline">아직 제안이 없습니다.</p>
      <p class="hint">「갱신 체험」에서 예시 문서 갱신을 실행하면, 그 문서 v2를 근거로 한 변경 제안 1건이 생깁니다.</p>`}
      <div class="km-card-foot"><button type="button" class="${pending.length ? 'km-btn-primary' : 'km-btn-quiet'}" data-goto="review">변경 검토</button></div>
    </section>

    <section class="km-card km-c7">
      <div class="km-card-head"><h2>주요 이슈 · 미결사항</h2><span class="km-card-note">수정하면 버전이 올라갑니다</span></div>
      <div class="km-items">${issues.map(itemHtml).join('') || '<p class="hint">등록된 이슈·미결사항이 없습니다. 「전체 현황」에서 다른 항목을 확인하세요.</p>'}</div>
    </section>

    <section class="km-card km-c5">
      <div class="km-card-head"><h2>최근 자료</h2><span class="km-card-note">최신 수집 기준</span></div>
      <div class="km-rows">${recentDocs.map((d) => {
        const v = d.versions[0];
        return `<div class="km-row">
          <div class="km-row-main">${esc(d.title)}</div>
          <div class="km-row-meta"><span class="km-tag">${esc(KIND_LABEL[d.kind] || d.kind)}</span><span>v${v.version_no}</span>${statusPillHtml(v.extract_status)}<span>${dt16(v.uploaded_at)}</span></div>
        </div>`;
      }).join('') || '<p class="hint">수집된 자료가 없습니다.</p>'}</div>
      <div class="km-card-foot"><button type="button" class="km-btn-quiet" data-goto="documents">자료 전체 보기</button></div>
    </section>

    <section class="km-card km-c7">
      <div class="km-card-head"><h2>최근 업무 이력</h2><span class="km-card-note">저장된 수정·승인과 등록된 이력</span></div>
      <div class="km-rows">${events.slice(0, 3).map((e) => `
        <div class="km-row">
          <div class="km-row-main">${esc(e.text).slice(0, 90) || '(내용 없음)'}</div>
          <div class="km-row-meta"><span class="km-tag">${e.kind}</span><span>${e.main}</span><span>${esc(e.by || '')}</span><span>${dt16(e.at)}</span></div>
        </div>`).join('') || '<p class="hint">아직 저장된 수정·승인 이력이 없습니다.</p>'}</div>
      <div class="km-card-foot"><button type="button" class="km-btn-quiet" data-goto="activities">업무 이력 보기</button></div>
    </section>

    <section class="km-card km-c5">
      <div class="km-card-head"><h2>갱신 체험</h2></div>
      <p class="km-oneline">예시 문서 갱신 → 근거 확인 → 변경 승인</p>
      <p class="hint">실제 파일을 수집하거나 AI로 분석하지 않고, 미리 준비한 문서 개정본(v1 → v2)을 반영합니다.</p>
      <div class="km-row-meta" style="margin-top:10px">
        <span class="km-tag ${lastSync ? 'ok' : ''}">${lastSync ? '반영 완료' : '아직 실행 안 함'}</span>
        ${lastSync ? `<span>${dt16(lastSync)}</span>` : ''}
        <span class="km-tag ${pending.length ? 'wait' : ''}">확인 대기 ${pending.length}건</span>
      </div>
      <div class="km-card-foot"><button type="button" class="km-btn-quiet" data-goto="sync">갱신 체험 열기</button></div>
    </section>

    <details class="km-all">
      <summary>전체 현황 · 편집 (${others.length}개 항목)</summary>
      <div class="km-all-body"><div class="km-items">${others.map(itemHtml).join('') || '<p class="hint">등록된 항목이 없습니다.</p>'}</div></div>
    </details>`;

  wireStatusControls(panel);
}

/* 현황 카드 안의 수정·이력·근거·이동 버튼 연결. 패널을 다시 그릴 때마다 이 함수만
   한 번 부른다 — 예전처럼 innerHTML 뒤에 바로 붙이던 것과 같은 시점이라 이벤트가
   빠지거나 두 번 붙지 않는다. */
function wireStatusControls(panel) {
  panel.querySelectorAll('[data-goto]').forEach((btn) => btn.addEventListener('click', () => showPanel(btn.dataset.goto)));
  panel.querySelectorAll('[data-open-all]').forEach((btn) => btn.addEventListener('click', () => {
    const det = panel.querySelector('.km-all');
    if (det) { det.open = true; det.scrollIntoView({ block: 'nearest' }); }
  }));
  /* 요약 카드의 「수정」 — 편집 자리는 「전체 현황」 안의 그 항목 하나뿐이다(폼을 두 벌
     만들지 않는다). 펼쳐서 그 항목으로 데려가고 폼을 연다. */
  panel.querySelectorAll('[data-goto-item]').forEach((btn) => btn.addEventListener('click', () => {
    const det = panel.querySelector('.km-all');
    if (det) det.open = true;
    const row = panel.querySelector(`.km-all [data-item-id="${btn.dataset.gotoItem}"]`);
    if (!row) return;
    panel.querySelectorAll('.km-focus-item').forEach((el) => el.classList.remove('km-focus-item'));
    row.classList.add('km-focus-item');
    const form = row.querySelector('.edit-form');
    if (form) form.style.display = 'block';
    row.scrollIntoView({ block: 'center' });
    const ta = row.querySelector('textarea');
    if (ta) ta.focus();
  }));
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
    /* 요구 5-3 — 이 어댑터는 HTTP 상태를 늘 200으로 돌려준다. 실패는 반환 데이터의
       ok/reason 에만 담기므로 반드시 그것을 보고 판단한다. 안내는 창을 띄우지 않고
       고치던 자리에 글자로 남긴다. */
    if (!r.ok) {
      let msg = form.querySelector('.km-msg');
      if (!msg) { msg = document.createElement('p'); msg.className = 'km-msg err'; form.appendChild(msg); }
      msg.textContent = r.reason === 'CONFLICT'
        ? `다른 곳에서 이미 수정되었습니다(현재 버전 v${r.currentNo}). 새로고침 후 다시 시도하세요.`
        : `저장하지 못했습니다 — ${r.reason}`;
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

/* 자료 탭 — 자료마다 큰 카드를 두면 목록을 훑을 수가 없다. 문서명을 앞세운 표 하나로
   두고 형식·버전·상태·수집 시각을 같은 열에서 비교하게 한다(수집 시각 최신 순). */
let lastDocsMessage = '';

async function renderDocumentsPanel() {
  const panel = document.getElementById('panel-documents');
  const q = state.scope === 'team' ? 'scope=team' : `projectId=${state.projectId}`;
  const data = await getJson(`/api/documents?${q}`);
  const memoFormHtml = state.scope === 'team' ? '' : `<details class="km-fold"><summary>텍스트 회의 메모 등록 — 파일 없이 바로 등록</summary>
    <div class="km-fold-body"><form class="inline-form" id="memoForm">
      <label>제목(선택)<input name="title" placeholder="예: 9월 3주 임차 조건 협의"></label>
      <label>내용<textarea name="text" rows="3" placeholder="회의 내용을 붙여넣거나 입력하세요" required></textarea></label>
      <button type="submit">메모 등록</button>
    </form></div></details>`;
  const rowsHtml = data.documents.slice()
    .sort((a, b) => String((b.versions[0] || {}).uploaded_at || '').localeCompare(String((a.versions[0] || {}).uploaded_at || '')))
    .map((d) => {
      const latest = d.versions[0];
      const noBody = latest && NO_BODY_STATUS.has(latest.extract_status);
      /* A5 — 버전이 하나도 없는 행 = 사전검사에 거절돼 등록되지 못한 자료. */
      const notes = latest
        ? [noBody ? '본문 추출 미지원 — 원문 보기를 제공하지 않습니다.' : '',
          EXTRACT_NOTE[latest.extract_status] || '',
          latest.extract_warnings ? latest.extract_warnings.split('\n')[0] : ''].filter(Boolean)
        : ['등록된 버전 없음 — 체험판 예시 자료입니다'];
      /* 요구 5-1 — 이 버튼이 주는 것은 원본 DOCX/XLSX/PDF 가 아니라 예시 본문 텍스트다.
         본문 자체가 없는 형식(.hwp 등)은 내려줄 것이 없으므로 「원본 미제공」으로 적는다. */
      const acts = latest
        ? `${noBody
            ? '<button type="button" disabled title="본문이 추출되지 않아 원문 보기를 열 수 없습니다.">본문 추출 미지원</button>'
            : `<button type="button" data-view="${latest.id}">원문 보기</button>`}`
          + (noBody
            ? '<button type="button" disabled title="체험판에는 원본 파일이 없고, 이 형식은 본문도 추출되지 않습니다. 로컬 앱에서는 보관된 원본을 내려받습니다.">원본 미제공</button>'
            : (d.kind === 'text_memo' ? '' : `<button type="button" data-download="${latest.id}">예시 본문 내려받기 (.txt)</button>`))
        : `<button type="button" data-discard="${d.id}">목록에서 정리</button>`;
      return `<tr>
        <td><div class="km-doc-name">${esc(d.title)}</div>${notes.map((n) => `<div class="km-doc-note">${esc(n)}</div>`).join('')}</td>
        <td><span class="km-tag">${esc(KIND_LABEL[d.kind] || d.kind)}</span></td>
        <td>${latest ? 'v' + latest.version_no : '—'}</td>
        <td>${latest ? statusPillHtml(latest.extract_status) : '<span class="status-pill failed">수집 실패</span>'}</td>
        <td>${latest ? dt16(latest.uploaded_at) : '—'}</td>
        <td class="km-acts">${acts}</td>
      </tr>`;
    }).join('');
  /* 마감 — 원본 삭제는 감지하지 않는다(동기화 엔진을 넓히지 않았다). 목록이 무엇을
     보장하고 무엇을 보장하지 않는지 화면에 그대로 적는다. */
  const collectedNote = `<p class="hint">이 목록은 <b>마지막 수집 시각</b> 기준입니다${data.lastSyncFinishedAt
    ? ` — 마지막 예시 갱신 체험: ${esc(data.lastSyncFinishedAt.slice(0, 16).replace('T', ' '))}`
    : ' — 아직 예시 갱신 체험을 실행한 적이 없습니다'}. 체험판에는 실제 공유폴더가 없습니다 — 모두 미리 준비한 가상 자료입니다.</p>`;
  panel.innerHTML = lastDocsMessage + memoFormHtml + collectedNote + (rowsHtml
    ? `<div class="km-table-wrap"><table class="km-table">
        <thead><tr><th>문서명</th><th>형식</th><th>버전</th><th>상태</th><th>수집 시각</th><th></th></tr></thead>
        <tbody>${rowsHtml}</tbody></table></div>`
    : '<p class="hint">등록된 자료가 없습니다. 「갱신 체험」 탭에서 예시 갱신을 체험하거나 위에서 메모를 등록하세요.</p>');

  panel.querySelectorAll('[data-view]').forEach((btn) => btn.addEventListener('click', () => openViewer(btn.dataset.view)));
  panel.querySelectorAll('[data-download]').forEach((btn) => btn.addEventListener('click', () => {
    window.location.href = `/api/doc-versions/${btn.dataset.download}/download`;
  }));
  panel.querySelectorAll('[data-discard]').forEach((btn) => btn.addEventListener('click', async () => {
    if (!confirm('등록된 버전이 없는 자료입니다. 목록에서 정리할까요?')) return;
    const r = await postJson(`/api/documents/${btn.dataset.discard}/discard-empty`, {});
    lastDocsMessage = r.ok ? ''
      : `<p class="km-msg err">정리하지 못했습니다 — ${esc(r.reason === 'HAS_VERSIONS' ? '이 자료에는 등록된 버전이 있어 정리하지 않습니다.' : r.reason)}</p>`;
    await renderDocumentsPanel();
  }));
  const memoForm = document.getElementById('memoForm');
  if (memoForm) memoForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const f = new FormData(e.target);
    const r = await postJson('/api/documents/text-memo', {
      projectId: state.projectId, title: f.get('title'), text: f.get('text'), actor: '담당자(시연)',
    });
    /* 요구 5-3 — 이 호출은 실패해도 HTTP 200 으로 돌아온다. 예전에는 ok 를 보지 않아
       빈 내용이 조용히 버려졌다(등록된 것처럼 보였다). */
    lastDocsMessage = r.ok ? ''
      : `<p class="km-msg err">메모를 등록하지 못했습니다 — ${esc(r.reason === 'EMPTY' ? '내용이 비어 있습니다.' : r.reason)}</p>`;
    await renderDocumentsPanel();
  });
}

/* B3-1 — 근거 원문은 화면을 덮는 모달이 아니라 오른쪽 상시 패널에 연다. 검색 결과나
   근거 칩을 눌러도 중앙의 결과·현황이 가려지지 않아 "검색 → 결과 → 근거 원문"이 한
   화면에 같이 남는다. #viewerBody는 아무것도 안 열렸을 때 비어 있게 두고(안내 문구는
   형제 요소 #viewerEmpty가 맡는다), 내용이 들어올 때만 채운다. */
/* 개편 — 예전엔 아무것도 안 열렸어도 오른쪽 패널이 약 440px 을 늘 차지했고, 「접기」와
   「닫기」가 따로 있어 차이가 불명확했다. 이제 기본은 닫힘이고 결과·근거 칩을 누를 때만
   열린다(닫기 하나로 통합). 좁은 화면에서는 같은 요소가 드로어로 뜬다(CSS). */
let evidenceReturnFocus = null;
function evidenceIsOpen() { return document.getElementById('layout').classList.contains('evidence-open'); }
function showEvidencePane(html) {
  if (!evidenceIsOpen()) {
    const a = document.activeElement;
    evidenceReturnFocus = (a && a !== document.body) ? a : null;
  }
  document.getElementById('side').hidden = false;
  document.getElementById('layout').classList.add('evidence-open');
  document.getElementById('viewerBody').innerHTML = html;
}
function clearEvidencePane() {
  const wasOpen = evidenceIsOpen();
  document.getElementById('layout').classList.remove('evidence-open');
  document.getElementById('side').hidden = true;
  document.getElementById('viewerBody').innerHTML = '';
  /* 열기 전에 있던 자리로 포커스를 되돌린다(키보드로 왔다면 그 자리에서 계속). */
  if (wasOpen && evidenceReturnFocus && document.contains(evidenceReturnFocus)) {
    try { evidenceReturnFocus.focus(); } catch (e) { /* 지워진 요소 */ }
  }
  evidenceReturnFocus = null;
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
  /* 제목 · 선택 버전 · 자료 기준 시각을 먼저 적고, 이것이 원본 뷰어가 아니라는 사실을
     제목 바로 아래에 둔다(실제 DOCX/PDF 원본을 여는 것처럼 보이면 안 된다). */
  let html = `<h3>${esc(meta.docVersion.original_filename)}</h3>
    <p class="hint">선택한 버전 <b>v${meta.docVersion.version_no}</b> · 수정일 ${esc((meta.docVersion.uploaded_at || '').slice(0, 10))} · 자료 기준 시각 ${dt16(meta.docVersion.uploaded_at)} · 상태 ${statusPillHtml(meta.docVersion.extract_status)}</p>
    <div class="km-prepared">미리 준비한 예시 본문</div>`;
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
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && evidenceIsOpen()) { e.preventDefault(); clearEvidencePane(); }
});

/* 업무 이력 탭 — 예전엔 발생일·작성자·내용이 작은 글씨로 한 줄에 몰려 있었다.
   이제 시각을 앞에 크게 두고 항목마다 이름을 붙여 줄을 나눈다. 등록 폼은 접어 두어
   읽는 일이 먼저 오게 하고(기능은 그대로), 등록 실패는 글자로 알린다. */
let lastActivityMessage = '';

async function renderActivitiesPanel() {
  const panel = document.getElementById('panel-activities');
  const data = await getJson(`/api/activities?projectId=${state.projectId}`);
  const formHtml = `<details class="km-fold"><summary>업무 이력 등록</summary><div class="km-fold-body">
    <form class="inline-form" id="activityForm">
    <label>발생일 <input type="date" name="occurredAt" required></label>
    <label>구분 <select name="category"><option>회의</option><option>협의</option><option>보고</option><option>의사결정</option><option>이행</option></select></label>
    <label>발생 사실/논의 <textarea name="fact" rows="2"></textarea></label>
    <label>결정/의견 <textarea name="decision" rows="2"></textarea></label>
    <label>이유 <input name="reason"></label>
    <label>다음 조치 <input name="nextAction"></label>
    <button type="submit">이력 등록</button>
  </form></div></details>`;
  const field = (label, value) => (value ? `<dt>${label}</dt><dd>${esc(value)}</dd>` : '');
  const listHtml = data.activities.map((a) => `
    <div class="activity-row">
      <div class="act-head">
        <span class="act-when">${esc(a.occurred_at)}</span>
        <span class="cat">${esc(a.category)}</span>
        <span>${esc(a.author)}</span>
      </div>
      <dl class="act-grid">
        ${field('발생 사실', a.fact)}
        ${field('결정 · 의견', a.decision)}
        ${field('이유', a.reason)}
        ${field('다음 조치', a.next_action)}
      </dl>
    </div>`).join('') || '<p class="hint">등록된 이력이 없습니다.</p>';
  panel.innerHTML = lastActivityMessage + formHtml + listHtml;
  document.getElementById('activityForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const f = new FormData(e.target);
    const r = await postJson('/api/activities', {
      projectId: state.projectId, occurredAt: f.get('occurredAt'), author: '담당자(시연)', category: f.get('category'),
      fact: f.get('fact'), decision: f.get('decision'), reason: f.get('reason'), nextAction: f.get('nextAction'),
    });
    /* 요구 5-3 — 이 호출도 HTTP 200 으로만 돌아온다. ok 를 보고 판단한다. */
    lastActivityMessage = r.ok ? '' : `<p class="km-msg err">이력을 등록하지 못했습니다 — ${esc(r.reason || '사유 없음')}</p>`;
    await renderActivitiesPanel();
    await renderStatusPanel();
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
     화면에 반영하는 것뿐이므로, 「수집했다」로 읽히지 않게 적는다.
     개편 — 예전엔 연결 이름만 적혀 있어서, B·C 사업장을 보는 중에 실행하면 A물류센터의
     문서가 바뀌는 것이 화면에 드러나지 않았다. 버튼 앞에서 대상 사업장·문서와 적용될
     변경을 먼저 적고, 대상이 아닌 사업장을 보고 있으면 그 사실과 이동 동선을 준다. */
  const t = conn.demo_target || null;
  const targetFund = t ? (tree.funds.find((f) => f.projects.some((p) => p.id === t.project_id)) || null) : null;
  const offTarget = !!(t && state.scope === 'project' && state.projectId !== t.project_id);
  const targetCardHtml = `
    <div class="sync-target">
      <div class="k">갱신 체험 대상</div>
      <div class="v">${esc(t ? t.project_name : conn.name)}</div>
      <div class="meta">${esc(t ? t.document_title : conn.root_path)}</div>
      ${t ? `<p class="hint">실행하면 이 문서가 <b>v${t.from_version} → v${t.to_version}</b> 로 바뀌고, 그 v${t.to_version} 본문을 근거로 한 <b>변경 제안 1건</b>이 「확인할 변경」에 생깁니다. 승인은 사람이 「변경 검토」에서 따로 누릅니다.</p>` : ''}
      <p class="hint"><b>실제 파일을 읽거나 AI로 분석하지 않습니다.</b> 미리 준비해 둔 개정본을 반영할 뿐입니다 — ${esc(conn.root_path)}</p>
      ${offTarget ? `<div class="km-notice">
        <span>지금 보고 있는 사업장은 <b>${esc(state.projectName)}</b>입니다. 이 체험은 <b>${esc(t.project_name)}</b>의 예시 문서를 갱신합니다 — 실행해도 지금 보는 사업장의 자료·현황은 바뀌지 않습니다.</span>
        <button type="button" class="km-btn-quiet" id="syncGotoTarget">${esc(t.project_name)} 보기</button>
      </div>` : ''}
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
    <details class="km-fold"><summary>폴더 연결 · 지원 파일 형식 (참고)</summary><div class="km-fold-body">
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
    ${formatsCardHtml(formats.formats)}
    </div></details>`;

  const gotoTarget = document.getElementById('syncGotoTarget');
  if (gotoTarget) gotoTarget.addEventListener('click', () => {
    selectProject(t.project_id, t.project_name, targetFund ? targetFund.name : '');
  });

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
        ? `<p class="km-msg info">예시 개정본을 반영했습니다(${new Date().toLocaleTimeString('ko-KR')}) — 문서 ${r.addedVersions}건이 v2가 되었고, 그 v2를 근거로 한 변경 제안 ${r.addedProposals}건이 「현황 → 확인할 변경」에 생겼습니다. <b>실제 파일 수집·AI 분석은 하지 않았습니다.</b></p>`
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

/* 승인·반려 결과와 충돌은 창을 띄우지 않고 이 자리에 글자로 남긴다(패널을 다시 그려도
   남아 있어야 하므로 값을 밖에 둔다 — 갱신 결과 문구와 같은 방식). */
let lastReviewMessage = '';

async function renderReviewPanel() {
  const panel = document.getElementById('side-review');
  const data = await getJson(`/api/proposals?projectId=${state.projectId}`);
  const proposals = data.proposals;
  const pendingCount = proposals.filter((p) => p.status === 'pending').length;

  panel.innerHTML = `
    ${lastReviewMessage}
    <p class="hint">대기 중인 제안 ${pendingCount}건 / 전체 ${proposals.length}건 · 승인은 근거를 열어 확인한 뒤 사람이 누릅니다.</p>
    <div id="proposalList">${proposals.map((p) => `
      <div class="proposal-card" data-proposal-id="${p.id}">
        <div><span class="proposal-status ${p.status}">${PROPOSAL_STATUS_LABEL[p.status]}</span>
          <b>${esc(FIELD_LABELS_CLIENT[p.key] || p.label)}</b> · ${esc(PROPOSAL_SOURCE_LABEL[p.source_type] || p.source_type)}</div>
        <div class="diff">
          <div class="before"><span class="km-lbl">현재 값 (v${p.current_version_no ?? '-'})</span>${esc(p.current_value) || '(없음)'}</div>
          <div class="after"><span class="km-lbl">제안 값 (기준 v${p.base_version_no})</span>${esc(p.proposed_value) || '(없음)'}</div>
        </div>
        <div class="hint">이유: ${esc(p.reason)}</div>
        <div style="margin-top:6px">${p.evidence_doc_version_id
          ? `<button type="button" class="evidence-chip" data-view="${p.evidence_doc_version_id}" data-segment="${p.evidence_segment_id ?? ''}">근거: ${esc(p.evidence_filename)}${p.evidence_text ? ' — ' + esc(p.evidence_text.slice(0, 30)) + '…' : ''}</button>`
          : '<span class="hint">연결된 근거 없음</span>'}</div>
        <div class="hint">제안: ${esc(p.proposed_by)}, ${dt16(p.proposed_at)}</div>
        ${p.status === 'pending' ? `<div class="km-decide"><button type="button" data-approve="${p.id}">승인</button> <button type="button" data-reject="${p.id}">반려</button></div>` : ''}
        ${p.status === 'conflict' ? `<p class="km-msg err">충돌 — ${esc(p.decision_note)}</p>` : ''}
        ${p.status === 'rejected' ? `<p class="hint">반려 사유: ${esc(p.decision_note)} (검토: ${esc(p.decided_by)})</p>` : ''}
        ${p.status === 'approved' ? `<p class="hint">승인: ${esc(p.decided_by)}, ${dt16(p.decided_at)}</p>` : ''}
      </div>`).join('') || '<p class="hint">제안이 없습니다. 「갱신 체험」에서 예시 문서 갱신을 실행하면 근거가 붙은 제안 1건이 생깁니다.</p>'}
    </div>
    <div class="km-card-foot"><button type="button" id="proposeSyntheticBtn">시연용 제안 생성 (테스트 제안 — AI 미연결)</button></div>`;

  document.getElementById('proposeSyntheticBtn').addEventListener('click', async () => {
    const statusData = await getJson(`/api/projects/${state.projectId}`);
    const target = statusData.status.find((s) => s.key === 'unresolved') || statusData.status[0];
    if (!target) {
      lastReviewMessage = '<p class="km-msg err">제안을 걸 현황 항목이 없습니다.</p>';
      await renderReviewPanel();
      return;
    }
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
    lastReviewMessage = r.ok ? '' : `<p class="km-msg err">제안을 만들지 못했습니다 — ${esc(r.reason)}</p>`;
    await renderReviewPanel();
    await renderStatusPanel();
  });
  panel.querySelectorAll('[data-approve]').forEach((btn) => btn.addEventListener('click', async () => {
    const r = await postJson(`/api/proposals/${btn.dataset.approve}/approve`, { approver: '팀장(시연)' });
    /* 요구 5-3 — 이 어댑터는 승인 충돌도 HTTP 200 + { ok:false, reason:'CONFLICT' } 로
       돌려준다. 상태 코드가 아니라 ok/reason 을 보고 판단하고, 결과를 글자로 남긴다. */
    lastReviewMessage = r.ok
      ? `<p class="km-msg info">승인했습니다 — 현황이 v${r.versionNo} 로 올라갔고, 승인한 근거가 그 항목에 연결됐습니다.</p>`
      : `<p class="km-msg err">${r.reason === 'CONFLICT'
        ? `승인하지 못했습니다 — 기준 버전 이후 현황이 이미 v${r.currentNo} 로 바뀌었습니다(충돌). 아래 제안이 충돌 상태로 남습니다.`
        : `승인하지 못했습니다 — ${esc(r.reason)}`}</p>`;
    await renderReviewPanel();
    await renderStatusPanel();
  }));
  panel.querySelectorAll('[data-reject]').forEach((btn) => btn.addEventListener('click', async () => {
    const reason = prompt('반려 사유를 입력하세요');
    if (reason === null) return;
    const r = await postJson(`/api/proposals/${btn.dataset.reject}/reject`, { rejecter: '팀장(시연)', reason });
    lastReviewMessage = r.ok ? '' : `<p class="km-msg err">반려하지 못했습니다 — ${esc(r.reason)}</p>`;
    await renderReviewPanel();
    await renderStatusPanel();
  }));
  panel.querySelectorAll('[data-view]').forEach((chip) => chip.addEventListener('click', () => openViewer(chip.dataset.view, chip.dataset.segment || null)));
}

/* 탭 전환 — 「변경 검토」는 탭 줄에 없고 현황의 「확인할 변경」 카드에서 들어온다.
   그래서 탭 목록과 패널 목록을 따로 맞추지 않고, 이름 하나로 둘을 같이 정한다
   (탭 없는 패널을 열면 탭 줄에서는 아무것도 선택되지 않는다).
   AI 질문 보조 탭은 이 체험판 주 화면에서 뺐다 — 체험판은 AI를 쓰지 않고, 안내줄에
   그렇게 적어 두었다. 로컬 업무용 앱의 AI 기능은 그대로 남아 있다. */
function showPanel(name) {
  document.querySelectorAll('.tab').forEach((t) => {
    const on = t.dataset.tab === name;
    t.classList.toggle('active', on);
    t.setAttribute('aria-selected', on ? 'true' : 'false');
  });
  document.querySelectorAll('.panel').forEach((p) => p.classList.toggle('active', p.id === `panel-${name}`));
  document.getElementById('center').scrollTop = 0;
}
document.querySelectorAll('.tab').forEach((tab) => tab.addEventListener('click', () => showPanel(tab.dataset.tab)));
document.getElementById('reviewBack').addEventListener('click', () => showPanel('status'));
document.getElementById('teamLink').addEventListener('click', selectTeam);

/* 마감 A2 — 검색이 언제나 전체 범위였다(결과마다 소속은 찍혔지만 범위를 좁힐 수 없었다).
   이제 범위를 화면에 명시하고, 사업장을 고른 상태면 그 사업장이 기본이다 — 전체 검색은
   사용자가 명시적으로 고를 때만 한다. 「팀 공통 지식」을 고른 상태면 좁힌 범위가
   팀 공통이 된다. 기존 버전 필터(A3: 이전 버전 포함)와는 독립이라 함께 쓸 수 있다. */
/* 고른 대상 이름은 바로 아래 도움말 줄이 전부 적는다 — 선택 상자까지 길게 늘어나면
   같은 줄의 「이전 버전 포함」·「검색 닫기」가 줄바꿈으로 밀린다. 짧게 둔다. */
function scopedSearchLabel() {
  return state.scope === 'team' ? '팀 공통 지식' : '현재 사업장';
}
function updateSearchScopeControl() {
  const sel = document.getElementById('searchScope');
  sel.options[0].textContent = scopedSearchLabel();
  sel.value = 'scoped'; // 대상을 바꾸면 기본값(그 대상으로 좁힌 검색)으로 돌아온다
  updateSearchScopeNote();
}
/* 개편 — 예전 도움말은 «그 사업장의 자료·현황만» 이라고 적었는데, 실제 검색 코드는
   docInProject() 에서 소속 펀드의 공통 문서도 그 사업장 결과에 포함한다(확인함).
   화면 글자를 코드에 맞춘다 — 「A 자료만 검색」은 사실이 아니다.
   또한 이 검색은 문자열 일치다. 「AI 검색」·「의미 기반 검색」이라고 적지 않는다. */
const MATCH_NOTE = '제목·본문에 <b>글자가 그대로 일치</b>하는 것만 찾습니다(AI 아님).';
function updateSearchScopeNote() {
  const sel = document.getElementById('searchScope');
  document.getElementById('searchScopeNote').innerHTML = (sel.value === 'all'
    ? '검색 범위: <b>전체</b> — 모든 펀드·사업장과 팀 공통 지식. '
    : (state.scope === 'team'
      ? '검색 범위: <b>팀 공통 지식</b> — 사업장 자료는 찾지 않습니다. '
      : `검색 범위: <b>현재 사업장 + 소속 펀드 공통자료</b> (${esc(state.projectName)} · ${esc(state.fundName)}). `)) + MATCH_NOTE;
}

/* 일치 구절만 강조한다(강조색을 줄 전체에 바르지 않는다). 잘라 보여줄 때도 일치한
   자리를 중심으로 잘라 무엇이 걸렸는지 결과 줄에서 바로 읽히게 한다. */
function matchSnippet(text, needle) {
  const s = String(text ?? '');
  if (!needle) return esc(s.slice(0, 80));
  const i = s.toLowerCase().indexOf(needle.toLowerCase());
  if (i < 0) return esc(s.slice(0, 80)) + (s.length > 80 ? '…' : '');
  const from = Math.max(0, i - 24);
  const to = Math.min(s.length, i + needle.length + 56);
  return (from > 0 ? '…' : '') + esc(s.slice(from, i))
    + `<mark>${esc(s.slice(i, i + needle.length))}</mark>`
    + esc(s.slice(i + needle.length, to)) + (to < s.length ? '…' : '');
}
function matchTitle(text, needle) {
  const s = String(text ?? '');
  const i = needle ? s.toLowerCase().indexOf(needle.toLowerCase()) : -1;
  if (i < 0) return esc(s);
  return esc(s.slice(0, i)) + `<mark>${esc(s.slice(i, i + needle.length))}</mark>` + esc(s.slice(i + needle.length));
}
const ICO_DOC = '<svg class="km-ico" viewBox="0 0 16 16" aria-hidden="true"><path d="M4 2h5l3 3v9H4z" fill="none" stroke="currentColor" stroke-width="1.3"/><path d="M9 2v3h3" fill="none" stroke="currentColor" stroke-width="1.3"/></svg>';
const ICO_SEG = '<svg class="km-ico" viewBox="0 0 16 16" aria-hidden="true"><path d="M2.5 4h11M2.5 7.5h11M2.5 11h7" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>';
const ICO_ITEM = '<svg class="km-ico" viewBox="0 0 16 16" aria-hidden="true"><rect x="2.5" y="2.5" width="11" height="11" rx="2" fill="none" stroke="currentColor" stroke-width="1.3"/><path d="M5.5 8.2l1.8 1.8 3.2-3.6" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/></svg>';

let searchTimer;
function runSearch() {
  clearTimeout(searchTimer);
  const q = document.getElementById('searchBox').value.trim();
  const includeOld = document.getElementById('includeOldVersions').checked;
  const scopeChoice = document.getElementById('searchScope').value;
  updateSearchScopeNote();
  const box = document.getElementById('searchResults');
  /* 검색어가 있으면 현황 그리드를 접고 결과가 주 작업 영역을 쓴다(작은 박스에 끼우지
     않는다). 검색을 닫으면 원래 탭이 그대로 다시 보인다 — 탭 선택은 건드리지 않는다. */
  document.getElementById('layout').classList.toggle('km-searching', !!q);
  document.getElementById('searchClose').hidden = !q;
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
    /* 문서 제목 일치 · 본문 구간 일치 · 현황 항목 일치는 서로 다른 종류다. 한데 합쳐
       「문서 N건」이라고 부르지 않고 묶음별로 무엇이 몇 건인지 따로 센다. */
    const group = (icon, title, rows) => (rows.length
      ? `<div class="km-res-group"><div class="km-res-head">${icon}${title} ${rows.length}건</div>${rows.join('')}</div>`
      : '');
    const html = group(ICO_DOC, '문서 제목 일치', r.documents.map((d) => `
        <div class="result"><span class="res-where">[${scopeLabel(d.scope, d.fund_name, d.project_names)}]</span>
          <span class="res-title">${matchTitle(d.title, q)}</span>${docLabel(d)}</div>`))
      + group(ICO_SEG, '본문 구간 일치', r.segments.map((s) => `
        <button type="button" class="result" data-view="${s.doc_version_id}" data-segment="${s.id}">
          <span class="res-where">[${scopeLabel(s.scope, s.fund_name, s.project_names)}]</span>
          <span class="res-title">${esc(s.title)}</span>${versionLabel(s)}
          <span class="res-snip">${matchSnippet(s.text, q)}</span></button>`))
      + group(ICO_ITEM, '현황 항목 일치', r.knowledgeItems.map((i) => `
        <div class="result"><span class="res-where">[${scopeLabel(i.scope, i.fund_name, i.project_name)}]</span>
          <span class="res-title">${esc(FIELD_LABELS_CLIENT[i.key] || i.key)}</span>${itemLabel(i)}
          <span class="res-snip">${matchSnippet(i.value, q)}</span></div>`));
    /* 결과가 없을 때 범위를 몰래 넓히지 않는다 — 넓힐지는 사람이 고른다. */
    box.innerHTML = html || `<div class="km-empty">
      <p>현재 범위에서 결과가 없습니다.</p>
      ${scopeChoice === 'all' ? '<p class="hint">전체 범위에서도 일치하는 글자가 없습니다.</p>'
        : '<button type="button" class="km-btn-quiet" id="searchWiden">전체 범위에서 검색</button>'}</div>`;
    box.querySelectorAll('[data-view]').forEach((el) => el.addEventListener('click', () => openViewer(el.dataset.view, el.dataset.segment || null)));
    const widen = document.getElementById('searchWiden');
    if (widen) widen.addEventListener('click', () => {
      const sel = document.getElementById('searchScope');
      sel.value = 'all';
      runSearch();
    });
  }, 200);
}
document.getElementById('searchBox').addEventListener('input', runSearch);
document.getElementById('includeOldVersions').addEventListener('change', runSearch);
document.getElementById('searchScope').addEventListener('change', runSearch);
/* 「검색 닫기」 — 검색 상태(검색어·범위·버전 선택·열려 있던 근거)를 정리하고
   원래 보고 있던 탭·사업장으로 돌아간다(탭 선택은 그대로 남아 있다). */
document.getElementById('searchClose').addEventListener('click', () => {
  document.getElementById('searchBox').value = '';
  document.getElementById('searchScope').value = 'scoped';
  document.getElementById('includeOldVersions').checked = false;
  clearEvidencePane();
  runSearch();
  document.getElementById('searchBox').focus();
});

loadTree();
