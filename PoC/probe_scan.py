"""
probe_scan.py — InitScan/GetText 실측 스크립트

사용법:
    python probe_scan.py <hwp_file_path>

출력:
    probe_result.json  — 전체 이벤트 로그
    probe_summary.txt  — 사람이 읽기 편한 요약 (state 분포, 필드 감지 여부 등)
"""
import json
import sys
import os

def main():
    if len(sys.argv) < 2:
        print("사용법: python probe_scan.py <hwp_file_path>")
        sys.exit(1)

    hwp_path = os.path.abspath(sys.argv[1])
    if not os.path.exists(hwp_path):
        print(f"파일 없음: {hwp_path}")
        sys.exit(1)

    print(f"[1] HWP 연결 중...")
    from hwp_controller import HwpController
    hwp = HwpController()
    hwp.connect(visible=True)

    print(f"[2] 파일 열기: {hwp_path}")
    hwp.open_file(hwp_path)

    print(f"[3] probe_scan 실행 중 (최대 500 이벤트)...")
    events = hwp.probe_scan(max_events=500)
    print(f"    → {len(events)}개 이벤트 수집 완료")

    # ── JSON 저장 ──────────────────────────────────────────
    result_path = os.path.join(os.path.dirname(__file__), "probe_result.json")
    with open(result_path, "w", encoding="utf-8") as f:
        json.dump(events, f, ensure_ascii=False, indent=2)
    print(f"[4] 전체 결과 저장: {result_path}")

    # ── 요약 분석 ──────────────────────────────────────────
    lines = []
    lines.append("=" * 60)
    lines.append("PROBE SUMMARY")
    lines.append("=" * 60)

    # state 분포
    from collections import Counter
    state_counter = Counter(ev["state"] for ev in events)
    lines.append("\n[state 분포]")
    for st, cnt in sorted(state_counter.items()):
        lines.append(f"  state={st} : {cnt}회")

    # CurFieldName 감지 여부
    field_events = [ev for ev in events if ev["cur_field_name"]]
    lines.append(f"\n[CurFieldName 감지]")
    if field_events:
        lines.append(f"  총 {len(field_events)}개 이벤트에서 필드 이름 확인됨")
        seen_fields = {}
        for ev in field_events:
            fn = ev["cur_field_name"]
            if fn not in seen_fields:
                seen_fields[fn] = ev
        for fn, ev in seen_fields.items():
            lines.append(f"  필드명={fn!r}  state={ev['state']}  text={ev['text']!r}")
    else:
        lines.append("  ❌ CurFieldName이 비어있음 — InitScan 중 필드 감지 불가")

    # CurCtrl 감지 여부
    ctrl_events = [ev for ev in events if ev["cur_ctrl_id"]]
    lines.append(f"\n[CurCtrl.CtrlID 감지]")
    if ctrl_events:
        ctrl_ids = Counter(ev["cur_ctrl_id"] for ev in ctrl_events)
        for cid, cnt in ctrl_ids.most_common():
            lines.append(f"  CtrlID={cid!r} : {cnt}회")
    else:
        lines.append("  ❌ CurCtrl을 읽을 수 없음")

    # 에러 목록
    err_events = [ev for ev in events if ev["notes"]]
    if err_events:
        lines.append(f"\n[에러/경고 ({len(err_events)}개)]")
        seen_notes = set()
        for ev in err_events:
            for note in ev["notes"]:
                if note not in seen_notes:
                    lines.append(f"  {note}")
                    seen_notes.add(note)

    # 전체 이벤트 테이블 (처음 60개)
    lines.append(f"\n[이벤트 테이블 (처음 60개)]")
    lines.append(f"  {'seq':>4}  {'st':>3}  {'field_name':16}  {'ctrl_id':8}  text")
    lines.append("  " + "-" * 72)
    for ev in events[:60]:
        fn = (ev["cur_field_name"] or "")[:16]
        cid = (ev["cur_ctrl_id"] or "")[:8]
        txt = repr(ev["text"])[:40]
        lines.append(f"  {ev['seq']:>4}  {ev['state']:>3}  {fn:16}  {cid:8}  {txt}")

    summary = "\n".join(lines)
    summary_path = os.path.join(os.path.dirname(__file__), "probe_summary.txt")
    with open(summary_path, "w", encoding="utf-8") as f:
        f.write(summary)

    print(f"[5] 요약 저장: {summary_path}")
    print()
    print(summary)


if __name__ == "__main__":
    main()
