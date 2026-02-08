# Shift Assignment Input Recognition Checklist

Target code paths:
- `js/script.js:536` `getColumnIndices`
- `js/script.js:779` `parseSurveyResponses`
- `js/script.js:277` `validateLocationMatching`
- `js/script.js:1183` `prepareShiftSlots`
- `js/script.js:1565` `findCandidatesFlexible`

Goal:
- Verify each survey column is recognized as the intended logic field.
- Detect why assignment becomes `0` even when file load itself succeeds.

---

## 1. Header Mapping Check (most important)

How to check:
- Open browser console after upload.
- Confirm `column index map` log values.

### 1-1. Required fields
- [ ] `count` points to the "participation count" column
- [ ] `pref1` points to 1st location preference column
- [ ] `pref2` points to 2nd location preference column (optional in operation)
- [ ] `pref3` points to 3rd location preference column (optional in operation)
- [ ] `fullName` or (`lastName` + `firstName`) is detected
- [ ] `grade` is detected

### 1-2. Common misrecognition risks
- [ ] `count` and `additionalCount` are NOT the same column
- [ ] `count` is NOT mapped to yes/no additional support question
- [ ] `preferredDates` and `ngDates` are not swapped
- [ ] `pref3` is not overwritten by `pref4`
- [ ] `firstName` is not misdetected as email/username column

---

## 2. Row Parsing Check (participant creation)

How to check:
- Run assignment and read `survey parse summary` logs.

### 2-1. Skip counters
- [ ] `valid participant count` is not 0
- [ ] `skip-by-zero-count` is not abnormally high
- [ ] `exempt count` is close to expected count
- [ ] `skip-by-empty-row` is close to expected count

### 2-2. Participant field sanity (sample 5-10 rows)
- [ ] `displayName` is non-empty
- [ ] `maxAssignments >= 1` for non-exempt rows
- [ ] `preferredLocations` has at least 1 location
- [ ] `preferredMonths` / `preferredDays` / `preferredDates` match raw input meaning
- [ ] `ngMonths` / `ngDays` / `ngDates` match raw input meaning

---

## 3. Shift Sheet Parsing Check

How to check:
- Review shift upload logs.

- [ ] `state.locations.length > 0`
- [ ] `state.dates.length > 0`
- [ ] school-day count (excluding Sat/Sun/Holiday) is > 0
- [ ] `total slots > 0`

---

## 4. Location Matching Check

How to check:
- After both files are loaded, review location matching result.

- [ ] `unmatchedSurvey = 0`
- [ ] survey location values match shift-side `baseName` exactly
- [ ] capacity suffix differences (example `(2-person)`) are normalized correctly

---

## 5. Candidate Extraction Check (pre-assignment)

How to check:
- Review run logs for participants, slots, and assignment result.

- [ ] `participants > 0`
- [ ] `total slots > 0`
- [ ] `assignedCount > 0` (new assignments, not only pre-assigned rows)
- [ ] if unassigned is high, verify constraints are not too strict

---

## 6. Per-Column Audit Template (for next step)

Fill one line per logic field:

| Logic field | Actual Excel column | Header text | `getColumnIndices` result | Sample cell value | Expected interpretation | OK/NG | Notes |
|---|---:|---|---|---|---|---|---|
| count |  |  |  |  | Example: `1x` -> `maxAssignments=1` |  |  |
| pref1 |  |  |  |  | Must match shift location name |  |  |
| pref2 |  |  |  |  | Must match shift location name |  |  |
| pref3 |  |  |  |  | Must match shift location name |  |  |
| pref4 |  |  |  |  | Optional by operation |  |  |
| fullName / last+first |  |  |  |  | Used for displayName |  |  |
| grade |  |  |  |  | Used in prioritization |  |  |
| preferredDates |  |  |  |  | If set, only those dates are allowed |  |  |
| preferredMonth |  |  |  |  | Month filter (when set) |  |  |
| preferredDay |  |  |  |  | Day-of-week filter (when set) |  |  |
| ngDates |  |  |  |  | Absolute exclusion |  |  |
| ngMonths |  |  |  |  | Absolute exclusion |  |  |
| ngDays |  |  |  |  | Absolute exclusion |  |  |
| additionalSupport / additionalCount |  |  |  |  | Additional support upper bound |  |  |

---

## 7. Temporary pass/fail criteria

Fail now:
- [ ] `count` misrecognized
- [ ] `pref1` not recognized
- [ ] participants = 0
- [ ] location mismatch exists

Warning:
- [ ] `pref3/pref4` overwrite
- [ ] `preferredDates/ngDates` swap
- [ ] sudden increase in skipped rows
