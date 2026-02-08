# Assignment Input Audit (2026-04/05)

Target files:
- Survey: `2026年4-5月_挨拶・旗振り当番シフトアンケート.xlsx`
- Shift: `2026年4-5月_旗振りシフト日と場所.xlsx`

Audit basis:
- Current implementation in `js/script.js` (no code changes)
- Checklist in `md/ASSIGNMENT_INPUT_CHECKLIST.md`

---

## 1) Header Mapping Check

### 1-1 Required fields

| Check item | Result | Evidence |
|---|---|---|
| `count` points to participation count column | NG | mapped index=`14` (Excel col 15), header=`<6>...参加可能回数を超えて追加...` |
| `pref1` mapped correctly | OK | index=`10` (Excel col 11), header=`第1希望` |
| `pref2` mapped correctly | OK | index=`11` (Excel col 12), header=`第2希望` |
| `pref3` mapped correctly | NG | index=`13` (Excel col 14), header=`第4希望` |
| `fullName` or `last+first` detected | OK | `fullName` index=`4` (Excel col 5) |
| `grade` detected | OK | index=`2` (Excel col 3) |

### 1-2 Misrecognition risk checks

| Check item | Result | Evidence |
|---|---|---|
| `count` and `additionalCount` are different | NG | both index=`14` |
| `count` is not yes/no additional support column | NG | sample values are `はい/いいえ` |
| `preferredDates` and `ngDates` are not swapped | NG | `preferredDates=null`, `ngDates=index 9 (Excel col 10: この日しか参加できない)` |
| `pref3` is not overwritten by `pref4` | NG | both `pref3` and `pref4` are index=`13` |
| `firstName` is not misdetected | NG (warning) | `firstName=index 1` (Excel col 2: `ユーザー名`) |

---

## 2) Row Parsing Check

Survey rows: `291`

### 2-1 Skip counters

| Metric | Result | Status |
|---|---:|---|
| valid participant count | 0 | NG |
| skip-by-zero-count | 284 | NG |
| exempt count | 0 | NG |
| skip-by-empty-row | 7 | OK |

Reason:
- `count` misrecognized to Excel col 15 (`はい/いいえ/空`) -> `maxAssignments` becomes 0 for almost all rows.

### 2-2 Participant field sanity

Status: `N/A` (current logic produced 0 participants)

---

## 3) Shift Sheet Parsing Check

| Check item | Result | Status |
|---|---:|---|
| location count | 14 | OK |
| weekday date count | 35 | OK |
| total slots | 490 | OK |
| unassigned slots at load | 490 | OK |

---

## 4) Location Matching Check

| Check item | Result | Status |
|---|---:|---|
| survey unique locations (first 100 rows by current logic) | 14 | OK |
| shift unique locations | 14 | OK |
| unmatched survey locations | 0 | OK |

---

## 5) Candidate Extraction Check (pre-assignment)

| Check item | Result | Status |
|---|---:|---|
| participants > 0 | 0 | NG |
| total slots > 0 | 490 | OK |
| slots with at least one candidate | 0 | NG |

---

## 6) Per-Column Audit (actual workbook columns)

| Logic field | Current mapped index (0-based) | Excel col | Current header | Non-empty rows | Status |
|---|---:|---:|---|---:|---|
| count | 14 | 15 | `<6>...参加可能回数を超えて追加...` | 198 | NG |
| pref1 | 10 | 11 | `第1希望` | 198 | OK |
| pref2 | 11 | 12 | `第2希望` | 113 | OK |
| pref3 | 13 | 14 | `第4希望` | 8 | NG |
| pref4 | 13 | 14 | `第4希望` | 8 | Warning |
| fullName | 4 | 5 | `<1-3> ... 氏名` | 291 | OK |
| grade | 2 | 3 | `<1-1> ... 学年` | 291 | OK |
| preferredDates | null | - | (not detected) | - | NG |
| ngDates | 9 | 10 | `<4-4>この日しか参加できない` | 21 | NG |
| additionalCount | 14 | 15 | `<6>...追加...` | 198 | OK |

Reference (expected raw columns in this workbook):
- participation count is Excel col 6 (`<2>期間中の参加可能回数...`) with `284` non-empty
- NG date question is Excel col 7 with `50` non-empty
- strict preferred date question is Excel col 10 with `21` non-empty
- 3rd preference is Excel col 13 with `32` non-empty

---

## 7) Diagnostic Control Run (forced expected mapping, no code changes)

Purpose:
- Confirm file content itself is assignable when mapped to intended columns.

Forced mapping used for diagnosis:
- `count=col6`, `pref1=11`, `pref2=12`, `pref3=13`, `pref4=14`, `ngDates=col7`, `preferredDates=col10`

Observed:
- participants=`198`
- exempt=`86`
- skip-by-zero=`0`
- skip-by-empty=`7`
- weekday slots=`490`
- slots with at least one candidate=`490`

Interpretation:
- The input file is not fundamentally broken.
- The failure is caused by header-to-logic misrecognition in current `js/script.js`.
