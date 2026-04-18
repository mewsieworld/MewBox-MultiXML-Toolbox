import re
from collections import defaultdict

def strip_cdata(text):
    """Strip CDATA wrapper from text."""
    match = re.match(r'<!\[CDATA\[(.*?)\]\]>', text, re.DOTALL)
    if match:
        return match.group(1).strip()
    return text.strip()

def parse_xml(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    rows = re.findall(r'<ROW>(.*?)</ROW>', content, re.DOTALL)

    records = []
    for row in rows:
        def get_field(name):
            match = re.search(rf'<{name}>(.*?)</{name}>', row, re.DOTALL)
            return match.group(1).strip() if match else ''

        record_id = get_field('Id')
        chr_type_raw = get_field('ChrType')
        title_raw = get_field('Title')

        title = strip_cdata(title_raw) if title_raw else title_raw

        # Normalize the ChrType combo as a sorted tuple of values
        values = [v.strip() for v in chr_type_raw.split('/') if v.strip()]
        # Keep original order for display but sort numerically for grouping key
        combo_key = tuple(sorted(values, key=lambda x: int(x) if x.isdigit() else x))
        combo_display = '/'.join(str(v) for v in sorted(values, key=lambda x: int(x) if x.isdigit() else x))

        records.append({
            'id': record_id,
            'chr_type_raw': chr_type_raw,
            'combo_key': combo_key,
            'combo_display': combo_display,
            'title': title,
        })

    return records

def analyze(records):
    # Group by unique ChrType combo
    combo_map = defaultdict(list)  # combo_key -> list of (id, title)
    for r in records:
        combo_map[r['combo_key']].append((r['id'], r['title'], r['combo_display']))

    print("=" * 80)
    print("CHRTYPE ANALYSIS REPORT")
    print("=" * 80)
    print(f"Total records: {len(records)}")
    print(f"Unique ChrType combinations: {len(combo_map)}")
    print()

    # Sort combos by number of values, then lexicographically
    sorted_combos = sorted(combo_map.items(), key=lambda x: (len(x[0]), x[0]))

    for i, (combo_key, entries) in enumerate(sorted_combos, 1):
        display = entries[0][2]  # combo_display from first entry (all same)
        print(f"Combination #{i}: {display}")
        print(f"  Values ({len(combo_key)}): {', '.join(combo_key)}")
        print(f"  Used by {len(entries)} record(s):")
        for (rid, title, _) in entries:
            print(f"    ID {rid:>4} | {title}")
        print()

    # Summary table
    print("=" * 80)
    print("SUMMARY TABLE (sorted by combo)")
    print("=" * 80)
    print(f"{'#':<4} {'IDs':<30} {'# Values':<10} {'ChrType Combo'}")
    print("-" * 80)
    for i, (combo_key, entries) in enumerate(sorted_combos, 1):
        ids = ', '.join(e[0] for e in entries)
        display = entries[0][2]
        print(f"{i:<4} {ids:<30} {len(combo_key):<10} {display}")

if __name__ == '__main__':
    import os
    script_dir = os.path.dirname(os.path.abspath(__file__))
    filepath = os.path.join(script_dir, 'PostBoxLevelUpGiftInfo.xml')
    output_path = os.path.join(script_dir, 'PostBoxLevelUpGiftInfo_ChrType_Report.txt')

    records = parse_xml(filepath)

    import io, sys
    buffer = io.StringIO()
    sys.stdout = buffer
    analyze(records)
    sys.stdout = sys.__stdout__

    report = buffer.getvalue()
    print(report)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(report)
    print(f"Report saved to: {output_path}")
