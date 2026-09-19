/* R04(3차 검토) — 이전 버전은 "값이 있는 최대 행·열"까지 빈 <td>를 전부 만들었다. 값
   두 개가 A1과 XFD10에만 있어도 163,840개의 <td>(1.47MB짜리 HTML)가 나왔다 — 실제
   Excel 파일에서 멀리 떨어진 셀 하나만 있어도 화면이 터질 수 있는 구조였다. 상한을
   넘으면 "실제 값이 있는 셀만" 나열하는 희소 모드로 전환해 고쳤었다.

   P02(5차 검토, 재현됨) — 그런데 희소 모드 자체엔 상한이 없었다(값이 있는 셀 전부를
   그대로 나열 — 값이 10만 개면 10만 줄). 게다가 maxCol/maxRow를 구할 때 쓰던
   `Math.max(...cells.map(...))`는 배열을 함수 인자로 펼치는 방식이라, 셀이 아주 많으면
   (실측 15만 개) 엔진의 인자 개수 한도에 걸려 RangeError로 죽었다 — 밀집/희소를 가르는
   기준 자체를 계산하다가 죽어버리는 문제였다. 순수 함수로 뺀 이유는 그대로다: 브라우저
   DOM 없이 Node에서 직접 실행해 반례를 재현·확인하기 위해서다. */

export const MAX_DENSE_CELLS = 2000; // 이 이하면 기존처럼 사각 표로(문맥이 자연스럽게 보임)
export const MAX_SPARSE_ROWS_PER_SHEET = 500; // 희소 모드에서 시트 하나당 실제로 나열할 값의 상한
export const MAX_TOTAL_OUTPUT_ROWS = 2000; // 시트를 몇 개 합쳐도 이 줄 수를 넘으면 나머지는 생략

export function esc(s) {
  return (s ?? '').toString().replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

export function parseCellAddress(addr) {
  const m = /^([A-Z]+)(\d+)$/.exec(addr);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { col, row: Number(m[2]) };
}

/* Math.max(...arr) 대신 순수 루프 — 배열을 함수 인자로 펼치지 않으므로 원소가
   몇 개든(실측 15만 개 포함) 호출 스택/인자 개수 한도에 걸리지 않는다. */
function maxOf(arr, pick) {
  let m = 0;
  for (const x of arr) { const v = pick(x); if (v > m) m = v; }
  return m;
}

export function renderExcelSegmentsAsGrid(segments) {
  const bySheet = new Map();
  for (const s of segments) {
    if (s.kind !== 'excel_cell' && s.kind !== 'excel_range') continue;
    const anchor = (s.cell_range || '').split(':')[0];
    const pos = parseCellAddress(anchor);
    if (!pos) continue;
    if (!bySheet.has(s.sheet_name)) bySheet.set(s.sheet_name, []);
    bySheet.get(s.sheet_name).push({ ...pos, text: s.text, cellRange: s.cell_range });
  }

  let out = '';
  let totalRowsRendered = 0; // 시트를 합쳐도 이 값이 상한을 넘으면 이후 내용은 생략(전체를 한 번에 DOM에 안 만듦)
  let truncatedSheets = 0;
  for (const [sheetName, cells] of bySheet) {
    if (totalRowsRendered >= MAX_TOTAL_OUTPUT_ROWS) { truncatedSheets++; continue; }

    const maxCol = maxOf(cells, (c) => c.col);
    const maxRow = maxOf(cells, (c) => c.row);
    out += `<h4>시트: ${esc(sheetName)}</h4>`;

    if (maxRow * maxCol > MAX_DENSE_CELLS) {
      /* 값이 있는 셀만 행 단위로 나열한다 — 멀리 떨어진 셀 하나 때문에 빈 칸을 전부
         만들지 않는다. 그래도 값 자체가 아주 많을 수 있으니(실측 15만 개) 시트당·전체
         출력 줄 수에 상한을 두고, 생략된 나머지가 있다는 사실을 숨기지 않는다. */
      const sorted = [...cells].sort((a, b) => (a.row - b.row) || (a.col - b.col));
      const budget = Math.max(0, Math.min(MAX_SPARSE_ROWS_PER_SHEET, MAX_TOTAL_OUTPUT_ROWS - totalRowsRendered));
      const shown = sorted.slice(0, budget);
      const omitted = sorted.length - shown.length;
      out += `<p class="hint">전체 범위(최대 ${maxRow}행 × ${maxCol}열)가 커서 값이 있는 셀만 나열합니다`
        + (omitted > 0 ? ` — ${sorted.length}개 중 ${shown.length}개만 표시, 나머지 ${omitted}개는 생략됨` : '')
        + '. 제목·단위 등 원문 전체는 다운로드로 확인하세요.</p>';
      out += '<div style="overflow-x:auto"><table><tbody>';
      for (const c of shown) {
        out += `<tr><td>${esc(c.cellRange)}</td><td>${esc(c.text)}</td></tr>`;
      }
      out += '</tbody></table></div>';
      totalRowsRendered += shown.length;
    } else {
      const grid = new Map(cells.map((c) => [`${c.col},${c.row}`, c]));
      out += '<div style="overflow-x:auto"><table><tbody>';
      for (let r = 1; r <= maxRow; r++) {
        out += '<tr>';
        for (let c = 1; c <= maxCol; c++) {
          const cell = grid.get(`${c},${r}`);
          out += `<td>${cell ? esc(cell.text) : ''}</td>`;
        }
        out += '</tr>';
      }
      out += '</tbody></table></div>';
      totalRowsRendered += maxRow;
    }
  }
  if (truncatedSheets > 0) {
    out += `<p class="hint">시트 ${truncatedSheets}개는 표시 상한을 넘어 생략됨 — 원문은 다운로드로 확인하세요.</p>`;
  }
  return out || '<p class="hint">추출된 셀이 없습니다.</p>';
}
