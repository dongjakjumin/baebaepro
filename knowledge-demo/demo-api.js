/* ══════════════════════════════════════════════════════════════
   지식관리 공개 체험판 — API 어댑터

   로컬 앱(Node + SQLite)의 화면(app.js·style.css)을 그대로 쓰되, 서버 대신
   이 파일이 `/api/*` 요청에 답한다. GitHub Pages 같은 정적 호스팅에서
   구성과 업무 흐름을 눌러 볼 수 있게 하는 것이 목적이다.

   지켜야 할 것 ─────────────────────────────────────────────
   · 자료는 전부 가상이다. 실제 DB·공유폴더·개인 PC 경로·API 키를 담지 않는다.
   · AI는 호출하지 않는다. 「자료 갱신」은 실제 파일 수집이 아니라 준비된
     예시 변경을 화면에 반영하는 «체험»이며, 화면에도 그렇게 적는다.
   · 저장은 데모 전용 키 하나(DEMO_KEY)만 쓴다. 재무모델·캘린더의 저장값을
     건드리지 않는다.

   app.js 는 fetch 로만 서버와 이야기한다. 그래서 fetch 하나만 가로채면
   화면 코드를 고치지 않고도 서버 없이 돌릴 수 있다.
   ══════════════════════════════════════════════════════════════ */
(function () {
  'use strict';

  var DEMO_KEY = 'jb-km-demo-v1';   /* 이 데모 말고는 아무것도 쓰지 않는 키 */
  var NOW = '2026-09-19T09:00:00.000Z';

  /* ── 가상 자료 ────────────────────────────────────────────
     실제 사내 자료가 아니다. 펀드·사업장·문서·사람 이름은 전부 지어낸 것이다. */
  function seed() {
    return {
      funds: [
        { id: 1, name: '가상1호 펀드', created_at: NOW },
        { id: 2, name: '가상2호 펀드', created_at: NOW }
      ],
      projects: [
        { id: 1, fund_id: 1, name: '가상 A물류센터', type: '물류센터', created_at: NOW },
        { id: 2, fund_id: 1, name: '가상 B산업단지', type: '산업단지', created_at: NOW },
        { id: 3, fund_id: 2, name: '가상 C오피스', type: '오피스', created_at: NOW }
      ],

      /* 문서 — versions[0] 이 최신이다(로컬 앱과 같은 순서).
         blocks 는 원문 보기용 문단, 각 문단이 곧 검색 구간(segment)이다. */
      documents: [
        {
          id: 1, scope: 'fund', fund_id: 1, project_id: null, kind: 'docx',
          title: '가상1호 운용가이드.docx', created_at: NOW,
          versions: [{
            id: 101, version_no: 1, original_filename: '가상1호 운용가이드.docx',
            extract_status: 'ready', extract_warnings: null, uploaded_at: NOW, size: 1180,
            blocks: [
              '가상1호 펀드 운용가이드',
              '이 문서는 가상1호 펀드 산하 사업장에 공통 적용되는 운용 원칙을 담는다.',
              '개별 사업장의 협의 결과를 곧바로 팀 전체 정책으로 삼지 않는다.',
              '임차인 조건 변경은 투자심의 절차를 거친다.'
            ]
          }]
        },
        {
          /* v1 → v2 가 있는 문서. 「최신/과거 버전」과 「예시 문서 갱신 체험」의 주인공. */
          id: 2, scope: 'project', fund_id: null, project_id: 1, kind: 'docx',
          title: '가상 A물류센터 인수인계 노트.docx', created_at: NOW,
          versions: [{
            id: 102, version_no: 1, original_filename: '가상 A물류센터 인수인계 노트.docx',
            extract_status: 'ready', extract_warnings: null, uploaded_at: '2026-09-01T02:00:00.000Z', size: 1420,
            blocks: [
              '가상 A물류센터 인수인계 노트 (최초본)',
              '임차인 조건 협의는 아직 진행 중이며 임대차계약은 체결되지 않았다.',
              '회신 예정일은 10월 8일이며 그때까지는 회신 대기 상태로 둔다.',
              '담보 설정과 책임준공 확약은 사람이 원문으로 확인해야 한다.'
            ]
          }]
        },
        {
          id: 3, scope: 'project', fund_id: null, project_id: 2, kind: 'xlsx',
          title: '가상 B산업단지 인허가 진행현황.xlsx', created_at: NOW,
          versions: [{
            id: 103, version_no: 1, original_filename: '가상 B산업단지 인허가 진행현황.xlsx',
            extract_status: 'partial',
            extract_warnings: '숨긴 시트 1개는 추출 대상에서 제외했습니다 · 계산값이 저장되지 않은 수식은 값 없이 표시합니다',
            uploaded_at: NOW, size: 9300,
            blocks: [
              '가상 B산업단지 인허가 진행현황',
              '건축허가 신청 접수 완료, 교통영향평가 협의 진행 중이다.',
              '임차인 조건 검토 의견은 별도 시트에 정리한다.',
              '착공 예정일은 관계기관 협의 결과에 따라 바뀔 수 있다.'
            ]
          }]
        },
        {
          id: 4, scope: 'project', fund_id: null, project_id: 3, kind: 'pdf',
          title: '가상 C오피스 매각검토 보고서.pdf', created_at: NOW,
          versions: [{
            id: 104, version_no: 1, original_filename: '가상 C오피스 매각검토 보고서.pdf',
            extract_status: 'ready',
            extract_warnings: '표 구조는 보존되지 않습니다 — 원문에서 확인하세요',
            uploaded_at: NOW, size: 21400,
            blocks: [
              '가상 C오피스 매각검토 보고서',
              '매각 자문사 선정을 위한 제안요청을 준비 중이다.',
              '공실률 개선 추이를 3개월 더 관찰한 뒤 매각 시점을 정한다.'
            ]
          }]
        },
        {
          /* 본문이 추출되지 않는 형식 — 「원문 보기」가 막히는 자리를 보여 준다. */
          id: 5, scope: 'project', fund_id: null, project_id: 1, kind: 'hwp',
          title: '가상 현장점검 결과.hwp', created_at: NOW,
          versions: [{
            id: 105, version_no: 1, original_filename: '가상 현장점검 결과.hwp',
            extract_status: 'conversion_needed', extract_warnings: null, uploaded_at: NOW, size: 4100,
            blocks: []
          }]
        }
      ],

      /* 현황 — 사업장별 카드. history 는 「이력」 버튼이 읽는다. */
      items: [
        mkItem(11, 'project', 1, 'manager', '담당자', '김가상 과장(가상)'),
        mkItem(12, 'project', 1, 'stage', '단계', '임차 조건 협의'),
        mkItem(13, 'project', 1, 'issue', '핵심 이슈', '임차인 조건 협의 중 — 계약 조건 미확정'),
        mkItem(14, 'project', 1, 'unresolved', '미결사항', '회신 예정일 10월 8일 — 회신 대기'),
        mkItem(15, 'project', 1, 'confirmed', '확정사항', '임대차계약 미체결'),
        mkItem(21, 'project', 2, 'manager', '담당자', '이가상 대리(가상)'),
        mkItem(22, 'project', 2, 'stage', '단계', '인허가 협의'),
        mkItem(23, 'project', 2, 'issue', '핵심 이슈', '교통영향평가 협의 지연 가능성'),
        mkItem(24, 'project', 2, 'unresolved', '미결사항', '착공 예정일 미확정'),
        mkItem(31, 'project', 3, 'manager', '담당자', '박가상 차장(가상)'),
        mkItem(32, 'project', 3, 'stage', '단계', '매각 검토'),
        mkItem(33, 'project', 3, 'issue', '핵심 이슈', '공실률 개선 추이 관찰 중'),
        mkItem(41, 'team', null, 'checklist', '팀 점검표',
          '신규 사업장 등록 시 펀드 연결과 원본 보관 여부를 반드시 확인한다.')
      ],

      activities: [
        { id: 1, project_id: 1, occurred_at: '2026-09-05', author: '김가상 과장(가상)', category: '협의',
          fact: '임차인 측과 임대 조건 2차 협의', decision: '회신 기한을 10월 8일로 합의',
          reason: '임차인 내부 검토 일정 반영', next_action: '10월 8일 회신 확인' }
      ],

      proposals: [],
      runs: [],
      nextIds: { doc: 6, ver: 106, item: 90, act: 10, prop: 1, run: 1, hist: 500 },

      /* 「예시 문서 갱신 체험」을 이미 실행했는지 */
      demoSyncDone: false
    };
  }

  function mkItem(id, scope, projectId, key, label, value) {
    return {
      id: id, scope: scope, project_id: projectId, fund_id: null,
      key: key, label: label, value: value,
      versionNo: 1, effectiveDate: '2026-09-01',
      updatedBy: '예시 데이터', updatedAt: NOW, note: '초기값',
      evidence: [],
      history: [{ version_no: 1, value: value, created_by: '예시 데이터', created_at: NOW, note: '초기값' }]
    };
  }

  /* 「예시 문서 갱신 체험」이 반영하는 준비된 변경 — v2 본문.
     실제 파일을 읽거나 AI로 분석해 만든 것이 아니라 미리 적어 둔 예시다. */
  var PREPARED_V2 = {
    documentId: 2,
    version: {
      id: 106, version_no: 2, original_filename: '가상 A물류센터 인수인계 노트.docx',
      extract_status: 'ready', extract_warnings: null, uploaded_at: '2026-09-18T01:00:00.000Z', size: 1510,
      blocks: [
        '가상 A물류센터 인수인계 노트 (개정본)',
        '임차인 조건 협의가 타결되어 임대차계약 체결을 준비한다.',
        '회신은 10월 8일에 확인되었고 미결사항을 해소할 수 있다.',
        '담보 설정과 책임준공 확약은 여전히 사람이 원문으로 확인해야 한다.'
      ]
    }
  };

  /* ── 저장 ────────────────────────────────────────────────
     데모 전용 키 하나만 쓴다. 다른 도구의 저장값은 읽지도 쓰지도 않는다. */
  var db = null;
  function load() {
    if (db) return db;
    try {
      var raw = localStorage.getItem(DEMO_KEY);
      if (raw) { db = JSON.parse(raw); return db; }
    } catch (e) { /* 사생활 보호 모드 등 — 메모리로만 돈다 */ }
    db = seed();
    return db;
  }
  function save() {
    try { localStorage.setItem(DEMO_KEY, JSON.stringify(db)); } catch (e) {}
  }
  function reset() {
    db = seed();
    try { localStorage.removeItem(DEMO_KEY); } catch (e) {}
  }
  window.__kmDemoReset = function () { reset(); location.reload(); };

  /* ── 조회 도우미 ───────────────────────────────────────── */
  var D = load;
  function fundOf(projectId) {
    var p = D().projects.filter(function (x) { return x.id === +projectId; })[0];
    if (!p) return null;
    return D().funds.filter(function (f) { return f.id === p.fund_id; })[0] || null;
  }
  function projectName(id) {
    var p = D().projects.filter(function (x) { return x.id === +id; })[0];
    return p ? p.name : '';
  }
  /* 문서가 그 사업장 것인가 — 펀드 공통자료는 산하 사업장 전부에서 보인다 */
  function docInProject(doc, projectId) {
    if (doc.project_id === +projectId) return true;
    if (doc.scope === 'fund') {
      var f = fundOf(projectId);
      return !!f && doc.fund_id === f.id;
    }
    return false;
  }
  function findVersion(vid) {
    var docs = D().documents;
    for (var i = 0; i < docs.length; i++) {
      for (var j = 0; j < docs[i].versions.length; j++) {
        if (docs[i].versions[j].id === +vid) return { doc: docs[i], ver: docs[i].versions[j] };
      }
    }
    return null;
  }
  /* 구간 id — 버전 id 와 문단 번호로 만든다(서버의 segment id 자리) */
  function segId(verId, idx) { return +verId * 100 + idx; }
  function segmentsOf(ver) {
    return (ver.blocks || []).map(function (t, i) {
      return { id: segId(ver.id, i), idx: i, text: t, block_ref: 'p' + (i + 1) };
    });
  }
  function itemById(id) {
    return D().items.filter(function (x) { return x.id === +id; })[0] || null;
  }

  /* ── 검색 ────────────────────────────────────────────────
     범위(현재 사업장 / 전체 / 팀 공통)와 버전(최신만 / 이전 포함)이
     실제로 동작해야 한다 — 예시 데이터 위에서 진짜로 거른다. */
  function search(q, includeOld, scope, projectId) {
    var needle = String(q || '').trim();
    var out = { ok: true, scope: scope || 'all', projectId: projectId ? +projectId : null,
                documents: [], segments: [], knowledgeItems: [] };
    if (!needle) return out;
    var lower = needle.toLowerCase();
    var hit = function (s) { return String(s || '').toLowerCase().indexOf(lower) >= 0; };

    D().documents.forEach(function (doc) {
      if (scope === 'project' && projectId && !docInProject(doc, projectId)) return;
      if (scope === 'team') return;                 /* 팀 공통 지식에는 문서를 두지 않는다 */
      var latest = doc.versions[0];
      if (hit(doc.title)) {
        out.documents.push({
          id: doc.id, title: doc.title, scope: doc.scope, kind: doc.kind,
          fund_name: doc.fund_id ? (D().funds.filter(function (f) { return f.id === doc.fund_id; })[0] || {}).name : (fundOf(doc.project_id) || {}).name,
          project_names: doc.project_id ? projectName(doc.project_id) : '',
          version_count: doc.versions.length,
          latest_version_no: latest ? latest.version_no : null,
          latest_extract_status: latest ? latest.extract_status : null
        });
      }
      doc.versions.forEach(function (ver) {
        if (!includeOld && ver.version_no !== latest.version_no) return;
        segmentsOf(ver).forEach(function (sg) {
          if (!hit(sg.text)) return;
          out.segments.push({
            id: sg.id, text: sg.text, doc_version_id: ver.id, document_id: doc.id,
            version_no: ver.version_no, current_version_no: latest.version_no,
            uploaded_at: ver.uploaded_at, title: doc.title, scope: doc.scope,
            fund_name: doc.fund_id ? (D().funds.filter(function (f) { return f.id === doc.fund_id; })[0] || {}).name : (fundOf(doc.project_id) || {}).name,
            project_names: doc.project_id ? projectName(doc.project_id) : ''
          });
        });
      });
    });

    D().items.forEach(function (it) {
      if (scope === 'project' && projectId) { if (!(it.scope === 'project' && it.project_id === +projectId)) return; }
      else if (scope === 'team') { if (it.scope !== 'team') return; }
      if (!hit(it.value)) return;
      out.knowledgeItems.push({
        id: it.id, key: it.key, scope: it.scope, project_id: it.project_id, fund_id: it.fund_id,
        value: it.value, project_name: it.project_id ? projectName(it.project_id) : '',
        fund_name: it.project_id ? (fundOf(it.project_id) || {}).name : '',
        pending_proposals: D().proposals.filter(function (p) { return p.item_id === it.id && p.status === 'pending'; }).length
      });
    });
    return out;
  }

  /* ── 라우터 ──────────────────────────────────────────── */
  function route(method, path, qs, body) {
    var q = new URLSearchParams(qs || '');
    var m;

    if (method === 'GET') {
      if (path === '/api/tree') {
        return { funds: D().funds.map(function (f) {
          var o = { id: f.id, name: f.name, created_at: f.created_at,
            projects: D().projects.filter(function (p) { return p.fund_id === f.id; })
              .map(function (p) { return { id: p.id, name: p.name, type: p.type, created_at: p.created_at }; }) };
          return o;
        }) };
      }
      if (path === '/api/team-status') {
        return { ok: true, status: D().items.filter(function (i) { return i.scope === 'team'; }).map(pubItem) };
      }
      if ((m = path.match(/^\/api\/projects\/(\d+)$/))) {
        var pid = +m[1];
        var p = D().projects.filter(function (x) { return x.id === pid; })[0];
        if (!p) return { ok: false, reason: 'NOT_FOUND' };
        return { ok: true, project: p, funds: [fundOf(pid)].filter(Boolean),
          status: D().items.filter(function (i) { return i.scope === 'project' && i.project_id === pid; }).map(pubItem) };
      }
      if (path === '/api/documents') {
        var dpid = q.get('projectId'), dscope = q.get('scope');
        var docs = D().documents.filter(function (doc) {
          if (dscope === 'team') return false;
          if (dpid) return docInProject(doc, dpid);
          return true;
        }).map(pubDoc);
        return { ok: true, documents: docs, lastSyncFinishedAt: lastSyncAt() };
      }
      if ((m = path.match(/^\/api\/knowledge-items\/(\d+)\/history$/))) {
        var hi = itemById(m[1]);
        return { ok: true, versions: hi ? hi.history.slice().reverse() : [] };
      }
      if ((m = path.match(/^\/api\/doc-versions\/(\d+)$/))) {
        var fv = findVersion(m[1]);
        if (!fv) return { ok: false, reason: 'NOT_FOUND' };
        return { ok: true,
          docVersion: {
            id: fv.ver.id, document_id: fv.doc.id, version_no: fv.ver.version_no,
            original_filename: fv.ver.original_filename, size: fv.ver.size,
            uploaded_at: fv.ver.uploaded_at, uploaded_by: '예시 데이터',
            extract_status: fv.ver.extract_status, extract_error: null,
            extract_warnings: fv.ver.extract_warnings,
            /* 실제 파일 경로를 담지 않는다 — 체험판에는 원본 파일이 없다 */
            file_path: '(체험판 — 원본 파일 없음)'
          },
          document: { id: fv.doc.id, title: fv.doc.title, kind: fv.doc.kind },
          segments: segmentsOf(fv.ver) };
      }
      if ((m = path.match(/^\/api\/doc-versions\/(\d+)\/view-html$/))) {
        var vv = findVersion(m[1]);
        if (!vv) return { ok: false, reason: 'NOT_FOUND' };
        if (!(vv.ver.blocks || []).length) return { ok: true, html: '<p>추출된 본문이 없습니다.</p>' };
        return { ok: true, html: vv.ver.blocks.map(function (t, i) {
          return '<p data-block-ref="p' + (i + 1) + '">' + escapeHtml(t) + '</p>';
        }).join('') };
      }
      if (path === '/api/activities') {
        var apid = +q.get('projectId');
        return { ok: true, activities: D().activities.filter(function (a) { return a.project_id === apid; })
          .sort(function (a, b) { return String(b.occurred_at).localeCompare(String(a.occurred_at)); }) };
      }
      if (path === '/api/source-connections') {
        return { ok: true, connections: [{
          id: 1, name: '가상 공유폴더 (체험판)',
          /* 개인 PC 경로를 담지 않는다 */
          root_path: '(체험판 — 실제 공유폴더에 연결되어 있지 않습니다)',
          connection_type: 'demo',
          fund_names: D().funds.map(function (f) { return f.name; }),
          project_names: D().projects.map(function (p) { return p.name; })
        }] };
      }
      if (path === '/api/folder-mappings') {
        var mp = [];
        D().funds.forEach(function (f, i) {
          mp.push({ id: 100 + i, source_connection_id: 1, relative_path: f.name,
            fund_id: f.id, project_id: null, review_status: 'confirmed', allowed_kind: 'fund', unsupported: 0 });
        });
        D().projects.forEach(function (p, i) {
          var f = D().funds.filter(function (x) { return x.id === p.fund_id; })[0];
          mp.push({ id: 200 + i, source_connection_id: 1, relative_path: (f ? f.name + '/' : '') + p.name,
            fund_id: null, project_id: p.id, review_status: 'confirmed', allowed_kind: 'project', unsupported: 0 });
        });
        return { ok: true, mappings: mp };
      }
      if (path === '/api/sync/runs') return { ok: true, runs: D().runs.slice().reverse(), running: false };
      if (path === '/api/sync/failures') return { ok: true, failures: [] };
      if (path === '/api/supported-formats') return { ok: true, formats: FORMATS };
      if (path === '/api/proposals') {
        var ppid = +q.get('projectId');
        return { ok: true, proposals: D().proposals.filter(function (p) {
          var it = itemById(p.item_id); return it && it.project_id === ppid;
        }).map(pubProposal).reverse() };
      }
      if (path === '/api/ai/status') return { ok: true, connected: false };
      if (path === '/api/search') {
        return search(q.get('q'), q.get('includeOldVersions') === '1', q.get('scope') || 'all', q.get('projectId'));
      }
    }

    if (method === 'POST') {
      if ((m = path.match(/^\/api\/knowledge-items\/(\d+)\/edit$/))) return editItem(+m[1], body);
      if ((m = path.match(/^\/api\/documents\/(\d+)\/discard-empty$/))) {
        var dd = D().documents.filter(function (x) { return x.id === +m[1]; })[0];
        if (!dd) return { ok: false, reason: 'NOT_FOUND' };
        if (dd.versions.length) return { ok: false, reason: 'HAS_VERSIONS', versionCount: dd.versions.length };
        db.documents = D().documents.filter(function (x) { return x.id !== +m[1]; }); save();
        return { ok: true };
      }
      if (path === '/api/documents/text-memo') return addMemo(body);
      if (path === '/api/activities') {
        var a = { id: D().nextIds.act++, project_id: +body.projectId, occurred_at: body.occurredAt,
          author: body.author || '담당자(체험)', category: body.category,
          fact: body.fact, decision: body.decision, reason: body.reason, next_action: body.nextAction };
        D().activities.push(a); save();
        return { ok: true, id: a.id };
      }
      if (path === '/api/sync/run') return runDemoSync();
      if ((m = path.match(/^\/api\/folder-mappings\/(\d+)\/confirm/))) return { ok: true };
      if (path === '/api/proposals') return addProposal(body);
      if ((m = path.match(/^\/api\/proposals\/(\d+)\/approve$/))) return decide(+m[1], 'approve', body);
      if ((m = path.match(/^\/api\/proposals\/(\d+)\/reject$/))) return decide(+m[1], 'reject', body);
      if (path === '/api/ai/ask') return { ok: false, reason: 'AI_NOT_CONNECTED' };
    }

    return { ok: false, reason: 'DEMO_ENDPOINT_NOT_IMPLEMENTED', path: path };
  }

  function pubItem(it) {
    return { id: it.id, key: it.key, label: it.label, value: it.value, versionNo: it.versionNo,
      effectiveDate: it.effectiveDate, updatedBy: it.updatedBy, updatedAt: it.updatedAt,
      note: it.note, evidence: it.evidence };
  }
  function pubDoc(doc) {
    return { id: doc.id, scope: doc.scope, fund_id: doc.fund_id, kind: doc.kind, title: doc.title,
      created_at: doc.created_at,
      versions: doc.versions.map(function (v) {
        return { id: v.id, version_no: v.version_no, original_filename: v.original_filename,
          extract_status: v.extract_status, extract_warnings: v.extract_warnings, uploaded_at: v.uploaded_at };
      }) };
  }
  function pubProposal(p) {
    var it = itemById(p.item_id) || {};
    var ev = p.evidence_doc_version_id ? findVersion(p.evidence_doc_version_id) : null;
    var evText = '';
    if (ev && p.evidence_segment_id != null) {
      var sg = segmentsOf(ev.ver).filter(function (s) { return s.id === p.evidence_segment_id; })[0];
      evText = sg ? sg.text : '';
    }
    return { id: p.id, key: it.key, label: it.label, status: p.status, source_type: p.source_type,
      current_value: it.value, current_version_no: it.versionNo,
      proposed_value: p.proposed_value, base_version_no: p.base_version_no, reason: p.reason,
      evidence_doc_version_id: p.evidence_doc_version_id,
      evidence_segment_id: p.evidence_segment_id,
      evidence_filename: ev ? ev.ver.original_filename : '',
      evidence_text: evText,
      proposed_by: p.proposed_by, proposed_at: p.proposed_at,
      decided_by: p.decided_by, decided_at: p.decided_at, decision_note: p.decision_note };
  }
  function lastSyncAt() {
    var r = D().runs;
    return r.length ? r[r.length - 1].finished_at : null;
  }

  function bump(it, value, actor, note, evidenceDocVersionId, evidenceSegmentId) {
    it.versionNo += 1;
    it.value = value;
    it.updatedBy = actor || '담당자(체험)';
    it.updatedAt = new Date().toISOString();
    it.note = note || '';
    if (evidenceDocVersionId) {
      var fv = findVersion(evidenceDocVersionId);
      it.evidence = fv ? [{ doc_version_id: fv.ver.id, segment_id: evidenceSegmentId || null,
        original_filename: fv.ver.original_filename, page_no: null, cell_range: null }] : [];
    }
    it.history.push({ version_no: it.versionNo, value: value,
      created_by: it.updatedBy, created_at: it.updatedAt, note: it.note });
  }

  function editItem(id, body) {
    var it = itemById(id);
    if (!it) return { ok: false, reason: 'NOT_FOUND' };
    if (+body.baseVersionNo !== it.versionNo) return { ok: false, reason: 'CONFLICT', currentNo: it.versionNo };
    bump(it, String(body.value == null ? '' : body.value), body.actor, body.note,
      body.evidenceDocVersionId, null);
    save();
    return { ok: true, versionNo: it.versionNo };
  }

  function addMemo(body) {
    var text = String(body.text || '').trim();
    if (!text) return { ok: false, reason: 'EMPTY' };
    var vid = D().nextIds.ver++;
    var doc = {
      id: D().nextIds.doc++, scope: 'project', fund_id: null, project_id: +body.projectId, kind: 'memo',
      title: String(body.title || '메모') + ' (체험 메모)', created_at: new Date().toISOString(),
      versions: [{ id: vid, version_no: 1, original_filename: String(body.title || '메모') + '.txt',
        extract_status: 'ready', extract_warnings: null, uploaded_at: new Date().toISOString(),
        size: text.length, blocks: text.split(/\n+/).filter(Boolean) }]
    };
    D().documents.push(doc); save();
    return { ok: true, documentId: doc.id, versionId: vid };
  }

  /* 「예시 문서 갱신 체험」 — 실제 파일을 읽지 않는다.
     준비된 v2 를 붙이고, 그 v2 를 근거로 하는 변경 제안을 만든다.
     이 둘이 곧 «근거 확인 → 승인» 흐름의 출발점이다. */
  function runDemoSync() {
    var started = new Date().toISOString();
    var doc = D().documents.filter(function (d) { return d.id === PREPARED_V2.documentId; })[0];
    var added = 0, proposed = 0;
    if (doc && !D().demoSyncDone) {
      doc.versions.unshift(JSON.parse(JSON.stringify(PREPARED_V2.version)));
      added = 1;
      var target = D().items.filter(function (i) { return i.project_id === 1 && i.key === 'unresolved'; })[0];
      if (target) {
        var segs = segmentsOf(PREPARED_V2.version);
        var ev = segs.filter(function (s) { return /10월 8일/.test(s.text); })[0] || segs[0];
        D().proposals.push({
          id: D().nextIds.prop++, item_id: target.id, status: 'pending', source_type: 'synthetic',
          proposed_value: '회신 확인 완료 — 미결사항 해소 (근거: 인수인계 노트 v2)',
          base_version_no: target.versionNo,
          reason: '예시 문서 갱신 체험 — 미리 준비된 v2 본문을 근거로 든 예시 제안입니다. AI 분석 결과가 아닙니다.',
          proposed_by: '담당자(체험)', proposed_at: new Date().toISOString(),
          evidence_doc_version_id: PREPARED_V2.version.id,
          evidence_segment_id: ev ? ev.id : null,
          decided_by: null, decided_at: null, decision_note: null
        });
        proposed = 1;
      }
      db.demoSyncDone = true;
    }
    var run = { id: D().nextIds.run++, source_connection_id: 1, started_at: started,
      finished_at: new Date().toISOString(), status: 'success',
      items_total: added, items_ok: added, items_failed: 0 };
    D().runs.push(run); save();
    return { ok: true, runId: run.id, status: 'success',
      total: added, itemsOk: added, itemsFailed: 0,
      demo: true, addedVersions: added, addedProposals: proposed,
      items: [] };
  }

  function addProposal(body) {
    var it = itemById(body.itemId);
    if (!it) return { ok: false, reason: 'ITEM_NOT_FOUND' };
    var p = { id: D().nextIds.prop++, item_id: it.id, status: 'pending',
      source_type: body.sourceType || 'synthetic',
      proposed_value: body.proposedValue, base_version_no: it.versionNo,
      reason: body.reason, proposed_by: body.proposedBy || '담당자(체험)',
      proposed_at: new Date().toISOString(),
      evidence_doc_version_id: body.evidenceDocVersionId || null,
      evidence_segment_id: body.evidenceSegmentId || null,
      decided_by: null, decided_at: null, decision_note: null };
    D().proposals.push(p); save();
    return { ok: true, id: p.id };
  }

  function decide(id, how, body) {
    var p = D().proposals.filter(function (x) { return x.id === id; })[0];
    if (!p) return { ok: false, reason: 'NOT_FOUND' };
    if (p.status !== 'pending') return { ok: false, reason: 'ALREADY_DECIDED' };
    var it = itemById(p.item_id);
    if (!it) return { ok: false, reason: 'ITEM_NOT_FOUND' };
    if (how === 'reject') {
      p.status = 'rejected'; p.decided_by = body.rejecter || '팀장(체험)';
      p.decided_at = new Date().toISOString(); p.decision_note = body.reason || '';
      save(); return { ok: true };
    }
    /* 승인 — 기준 버전 이후 현황이 바뀌었으면 충돌로 남긴다(로컬 앱과 같은 규칙) */
    if (p.base_version_no !== it.versionNo) {
      p.status = 'conflict';
      p.decision_note = '기준 버전 v' + p.base_version_no + ' 이후 현황이 v' + it.versionNo + ' 로 바뀌었습니다.';
      save();
      return { ok: false, reason: 'CONFLICT', currentNo: it.versionNo };
    }
    bump(it, p.proposed_value, body.approver || '팀장(체험)', '변경 제안 승인',
      p.evidence_doc_version_id, p.evidence_segment_id);
    p.status = 'approved'; p.decided_by = body.approver || '팀장(체험)';
    p.decided_at = new Date().toISOString();
    save();
    return { ok: true, versionNo: it.versionNo };
  }

  /* 지원 형식 — 로컬 앱과 같은 판정을 쓰되, 근거는 «체험판» 기준으로 짧게 적는다. */
  var FORMATS = [
    { ext: '.docx', level: 'supported', can: '본문·표·목록을 추출해 검색·근거 인용에 쓴다. 원문 보기 제공.',
      limit: '없음(예시 자료 범위 안에서)', basis: '체험판은 미리 추출해 둔 예시 본문을 보여 줍니다.' },
    { ext: '.xlsx', level: 'supported', can: '셀·병합범위를 추출한다(셀 주소 단위 근거).',
      limit: '숨긴 시트·행·열 제외 · 계산값이 없는 수식은 값 없이 경고만',
      basis: '체험판은 미리 추출해 둔 예시 본문을 보여 줍니다.' },
    { ext: '.pdf', level: 'supported', can: '텍스트를 페이지 단위로 추출한다.',
      limit: '표 구조는 보존되지 않음 · 스캔(이미지) PDF는 텍스트가 없어 추출되지 않음',
      basis: '체험판은 미리 추출해 둔 예시 본문을 보여 줍니다.' },
    { ext: '.hwp / .hwpx', level: 'unsupported', can: '원본 보관·다운로드만 된다.',
      limit: '본문을 추출하지 않는다 — 검색·근거 인용 대상이 아니다',
      basis: '로컬 앱에도 추출 경로가 없습니다(체험판도 같게 표시).' },
    { ext: '.doc / .xls (레거시)', level: 'unsupported', can: '원본 보관·다운로드만 된다.',
      limit: '본문을 추출하지 않는다', basis: '로컬 앱에도 추출 경로가 없습니다.' },
    { ext: '암호가 걸린 파일', level: 'unknown', can: '확인되지 않았다.',
      limit: '지원·미지원 어느 쪽으로도 확인하지 않았다',
      basis: '실제 암호 파일로 시험한 적이 없어 «미확인»으로 둡니다.' }
  ];

  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  /* ── fetch 가로채기 ──────────────────────────────────────
     app.js 는 오직 fetch 로만 서버와 이야기한다. /api/ 로 시작하는 요청만
     여기서 답하고 나머지(정적 파일 등)는 원래 fetch 로 넘긴다. */
  var realFetch = window.fetch ? window.fetch.bind(window) : null;
  window.fetch = function (input, init) {
    var url = (typeof input === 'string') ? input : (input && input.url) || '';
    var method = ((init && init.method) || (input && input.method) || 'GET').toUpperCase();
    var idx = url.indexOf('/api/');
    if (idx < 0) return realFetch ? realFetch(input, init) : Promise.reject(new Error('no fetch'));

    var rest = url.slice(idx);
    var qi = rest.indexOf('?');
    var path = qi >= 0 ? rest.slice(0, qi) : rest;
    var qs = qi >= 0 ? rest.slice(qi + 1) : '';
    var body = {};
    try { if (init && init.body) body = JSON.parse(init.body); } catch (e) {}

    var data;
    try { data = route(method, path, qs, body); }
    catch (e) { data = { ok: false, reason: 'DEMO_ERROR', message: String(e && e.message || e) }; }

    /* 실제 통신처럼 한 박자 뒤에 답한다 — 화면의 «불러오는 중» 표시가 살아 있게 */
    return new Promise(function (resolve) {
      setTimeout(function () {
        resolve({
          ok: true, status: 200,
          json: function () { return Promise.resolve(data); },
          text: function () { return Promise.resolve(JSON.stringify(data)); }
        });
      }, 40);
    });
  };

  /* ── 내려받기 ────────────────────────────────────────────
     app.js 는 location.href 로 서버 파일을 받으러 간다. 체험판에는 원본
     파일이 없으므로, 캡처 단계에서 가로채 예시 본문을 텍스트로 내려 준다.
     (캡처 단계라 app.js 의 클릭 처리보다 먼저 실행된다 — 화면 코드는 그대로 둔다) */
  document.addEventListener('click', function (e) {
    var btn = e.target && e.target.closest && e.target.closest('[data-download]');
    if (!btn) return;
    e.preventDefault();
    e.stopPropagation();
    var fv = findVersion(btn.getAttribute('data-download'));
    if (!fv) { alert('체험판에는 이 자료의 원본이 없습니다.'); return; }
    var body = (fv.ver.blocks || []).join('\n\n') ||
      '이 형식은 본문을 추출하지 않습니다 — 체험판에는 원본 파일이 없습니다.';
    var txt = '[가상 자료 체험판] ' + fv.ver.original_filename + ' (v' + fv.ver.version_no + ')\n' +
      '실제 사내 자료가 아닙니다.\n\n' + body + '\n';
    var blob = new Blob([txt], { type: 'text/plain;charset=utf-8' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = fv.ver.original_filename.replace(/\.[^.]+$/, '') + ' (체험판).txt';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }, true);

  load();
})();
