import json

with open(r"logs\sira_raw_dump.jsonl", encoding="utf-8") as f:
    for line in f:
        d = json.loads(line)
        if d["folio"] == "56579":
            print("=" * 70)
            print(f"FOLIO 56579 - body completo:")
            print("=" * 70)
            print(d["body_raw"])
            print("=" * 70)
            break
