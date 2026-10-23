import json
import re
from pathlib import Path
from openpyxl import load_workbook

INPUT_FILE = "好き声.xlsx"
OUTPUT_FILE = "data.js"


def rows_as_dicts(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(v).strip() if v is not None else "" for v in rows[0]]
    out = []
    for row in rows[1:]:
        item = {headers[i]: row[i] for i in range(min(len(headers), len(row))) if headers[i]}
        if any(v is not None and v != "" for v in item.values()):
            out.append(item)
    return out


def clean(value):
    return "" if value is None else str(value).strip()


def parse_representative_works(value):
    """Convert the Excel multiline representative-work cell into displayable entries."""
    if value is None:
        return []
    text = str(value).strip()
    if not text:
        return []
    # Usually each work is on a separate line, often prefixed by a bullet.
    parts = [p.strip() for p in re.split(r"\r?\n", text) if p.strip()]
    cleaned = []
    for part in parts:
        part = re.sub(r"^[•・\-*\s]+", "", part).strip()
        if part:
            cleaned.append(part)
    return cleaned


def main():
    path = Path(INPUT_FILE)
    if not path.exists():
        raise SystemExit(f"Excelファイルが見つかりません: {INPUT_FILE}（リポジトリ直下に置いてください）")

    wb = load_workbook(path, data_only=True, read_only=True)
    required = ["作品一覧", "声優一覧", "キャラクター名", "定番声優一覧"]
    missing = [name for name in required if name not in wb.sheetnames]
    if missing:
        raise SystemExit("必要なシートがありません: " + ", ".join(missing))

    work_rows = rows_as_dicts(wb["作品一覧"])
    actor_rows = rows_as_dicts(wb["声優一覧"])
    appearance_rows = rows_as_dicts(wb["キャラクター名"])
    standard_rows = rows_as_dicts(wb["定番声優一覧"])

    works = []
    work_names = {}
    for row in work_rows:
        wid, name = clean(row.get("作品ID")), clean(row.get("作品名"))
        if wid and name:
            works.append({"id": wid, "name": name, "reading": clean(row.get("読み仮名"))})
            work_names[wid] = name

    voice_actors = []
    actor_names = {}
    for row in actor_rows:
        aid, name = clean(row.get("声優ID")), clean(row.get("声優名"))
        if not aid or not name:
            continue
        actor_names[aid] = name
        actor = {"id": aid, "name": name, "showInList": clean(row.get("声優一覧に表示")) or "T"}
        reps = parse_representative_works(row.get("代表作"))
        if reps:
            actor["representativeWorks"] = reps
        voice_actors.append(actor)

    characters = []
    seen = set()
    for row in appearance_rows:
        aid, wid = clean(row.get("声優ID")), clean(row.get("作品ID"))
        if not aid or not wid:
            continue
        cname = clean(row.get("キャラクター名"))
        key = (aid, wid, cname)
        if key in seen:
            continue
        seen.add(key)
        characters.append({
            "voiceActorId": aid,
            "voiceActorName": clean(row.get("声優名")) or actor_names.get(aid, ""),
            "workId": wid,
            "workName": clean(row.get("作品名")) or work_names.get(wid, ""),
            "characterName": cname,
        })

    standard_voice_actors = []
    seen_ids = set()
    for row in standard_rows:
        aid = clean(row.get("声優ID"))
        # Only include rows marked as visible, when that column is present.
        visible = clean(row.get("定番声優に表示"))
        if aid and aid not in seen_ids and visible != "非表示":
            standard_voice_actors.append(aid)
            seen_ids.add(aid)

    output = (
        "const works = " + json.dumps(works, ensure_ascii=False, separators=(",", ":")) + ";\n\n"
        "const voiceActors = " + json.dumps(voice_actors, ensure_ascii=False, separators=(",", ":")) + ";\n\n"
        "const characters = " + json.dumps(characters, ensure_ascii=False, separators=(",", ":")) + ";\n\n"
        "const standardVoiceActors = " + json.dumps(standard_voice_actors, ensure_ascii=False, separators=(",", ":")) + ";\n"
    )
    Path(OUTPUT_FILE).write_text(output, encoding="utf-8")
    print(f"data.js を生成しました：作品 {len(works)}件 / 声優 {len(voice_actors)}人 / 出演関係 {len(characters)}件 / 定番声優 {len(standard_voice_actors)}人")


if __name__ == "__main__":
    main()
