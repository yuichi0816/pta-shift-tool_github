/**
 * PTA旗振りデータ検証ツール
 * アンケートデータの不備を検出する
 */

document.addEventListener('DOMContentLoaded', () => {
    // ========================================================================
    // State
    // ========================================================================
    const state = {
        surveyData: null,
        shiftData: null,
        validLocations: [], // Shift data places
        validDates: [],     // Shift data dates
        issues: []          // List of found issues
    };

    // ========================================================================
    // DOM Elements
    // ========================================================================
    const surveyFile = document.getElementById('surveyFile');
    const shiftFile = document.getElementById('shiftFile');
    const surveyUploadBox = document.getElementById('surveyUploadBox');
    const shiftUploadBox = document.getElementById('shiftUploadBox');
    const surveyStatus = document.getElementById('surveyStatus');
    const shiftStatus = document.getElementById('shiftStatus');
    const resultSection = document.getElementById('resultSection');
    const validationSummary = document.getElementById('validationSummary');
    const issuesList = document.getElementById('issuesList');
    const tabs = document.querySelectorAll('.tab');

    // ========================================================================
    // Event Listeners
    // ========================================================================
    surveyFile.addEventListener('change', async (e) => {
        await handleFileUpload(e.target.files[0], 'survey');
    });

    shiftFile.addEventListener('change', async (e) => {
        await handleFileUpload(e.target.files[0], 'shift');
    });

    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            tabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            renderIssues(tab.dataset.filter);
        });
    });

    // ========================================================================
    // File Handling
    // ========================================================================
    async function handleFileUpload(file, type) {
        if (!file) return;

        try {
            const data = await parseExcelFile(file);

            if (type === 'survey') {
                state.surveyData = data;
                surveyUploadBox.classList.add('uploaded');
                surveyStatus.textContent = `✓ ${file.name} を読み込みました (${data.length - 1}件)`;
            } else {
                state.shiftData = data;
                processShiftData();
                shiftUploadBox.classList.add('uploaded');
                shiftStatus.textContent = `✓ ${file.name} を読み込みました (場所・日付情報を取得)`;
            }

            // Always run validation if survey data exists
            if (state.surveyData) {
                runValidation();
            }

        } catch (error) {
            console.error(error);
            const statusEl = type === 'survey' ? surveyStatus : shiftStatus;
            statusEl.textContent = `❌ エラー: ${error.message}`;
            statusEl.style.color = 'var(--color-danger)';
        }
    }

    async function parseExcelFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array', codepage: 932 });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
                    resolve(jsonData);
                } catch (error) {
                    reject(error);
                }
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    }

    // ========================================================================
    // Data Processing
    // ========================================================================
    function processShiftData() {
        if (!state.shiftData || state.shiftData.length < 2) return;

        // 1. Get Headers
        let headerRowIndex = -1;
        for (let i = 0; i < Math.min(5, state.shiftData.length); i++) {
            const row = state.shiftData[i];
            if (row && row[0] && row[0].toString().includes('日')) {
                headerRowIndex = i;
                break;
            }
        }
        if (headerRowIndex === -1) headerRowIndex = 1;

        const headerRow = state.shiftData[headerRowIndex];

        // 2. Extract Valid Locations
        state.validLocations = [];
        for (let i = 2; i < headerRow.length; i++) {
            if (headerRow[i]) {
                const name = headerRow[i].toString().trim();
                // Exclude metadata cols
                if (!['更新', '変更', '削除'].some(k => name.includes(k))) {
                    state.validLocations.push(name);
                }
            }
        }
        console.log("Valid Locations:", state.validLocations);
    }

    // ========================================================================
    // Validation Logic
    // ========================================================================
    function runValidation() {
        state.issues = [];
        const headers = state.surveyData[0];
        const colIndices = getColumnIndices(headers);
        const seenParticipants = new Map(); // For duplicate check

        // 1. Column Check
        if (colIndices.count === null) {
            addIssue('global', 'critical', '「参加可能回数」のカラムが見つかりません', 'ヘッダーに「参加可能回数」または「希望回数」を含む列が必要です');
        }
        if (colIndices.pref1 === null) {
            addIssue('global', 'warning', '「第1希望」のカラムが見つかりません', '希望場所のロジックが正常に動作しない可能性があります');
        }

        // 2. Row Checks
        for (let i = 1; i < state.surveyData.length; i++) {
            const row = state.surveyData[i];
            const rowIndex = i + 1; // 1-based index for display

            // Helper to get raw val
            const getVal = (idx) => (idx !== null && row[idx]) ? row[idx].toString().trim() : '';

            // Skip empty rows
            if (!row[1] && !getVal(colIndices.participation)) continue;

            // 氏名の取得（新形式: fullName, 旧形式: lastName + firstName）
            const fullName = getVal(colIndices.fullName);
            const lastName = getVal(colIndices.lastName);
            const firstName = getVal(colIndices.firstName);
            const grade = getVal(colIndices.grade);
            const email = row[1] ? row[1].toString() : ''; // Email is usually col 1 in Google Forms

            // 氏名を統一的に取得
            const displayName = fullName || `${lastName}${firstName}` || email || `NoName_Row${rowIndex}`;
            const userKey = `${grade}_${displayName}`;

            // Check A: Duplicates
            // Only check if we have a name
            if (fullName || (lastName && firstName)) {
                // 学年を含めず、氏名のみでチェック (学年の入力ミスがあっても検知できるようにする)
                const uniqueKey = fullName || `${lastName}_${firstName}`;
                if (seenParticipants.has(uniqueKey)) {
                    // 同じ氏名が見つかった場合
                    const prevRow = seenParticipants.get(uniqueKey);
                    addIssue(i, 'warning', '重複回答の可能性があります', `"${displayName}" さんは行 ${prevRow} でも見つかりました。(学年に関わらず氏名が一致しています)`);
                } else {
                    seenParticipants.set(uniqueKey, rowIndex);
                }
            }

            // Check B: Missing Required Info
            if (!fullName && !lastName && !firstName) {
                addIssue(i, 'error', '氏名が入力されていません', '割り当て時に誰の希望か特定できません。');
            }
            if (!grade) {
                addIssue(i, 'warning', '学年が入力されていません', '同姓同名の判別が難しくなる可能性があります。');
            }

            // Check C: Count Validity (希望回数のチェック)
            const participationStr = getVal(colIndices.participation);
            // If they are exempted, skip other checks
            if (participationStr.includes('免除')) continue;


            // 希望回数のチェック
            const countValue = getVal(colIndices.count);

            // 「免除」が含まれている場合は正常な免除希望者として扱う（エラーにしない）
            if (countValue !== '' && countValue.includes('免除')) {
                // 免除希望者は正常なのでチェックをスキップ
                // （割り当て時に自動的に除外される）
            } else if (countValue === '') {
                addIssue(i, 'error', '希望回数が未入力です', '参加回数を入力してください。「期間中○回」のように、期間全体での合計回数を数値で入力する必要があります。免除希望の場合は「免除希望」と入力してください。');
            } else {
                // 数値抽出の試み
                const normalizedValue = countValue.replace(/[０-９]/g, s =>
                    String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
                );
                const numberMatch = normalizedValue.match(/(\d+)/);

                if (!numberMatch) {
                    addIssue(i, 'error', '希望回数から数値を抽出できません', `入力値: "${countValue}" → 数値が含まれていません。「期間中5回」のように数値を含めてください。免除希望の場合は「免除希望」と入力してください。`);
                } else {
                    const countVal = parseInt(numberMatch[1]);
                    if (countVal === 0) {
                        addIssue(i, 'error', '希望回数が0です', '参加希望の場合、1以上の回数を入力してください。参加しない場合は希望回数欄に「免除希望」と入力してください。');
                    }
                }
            }


            // Check D: Locations
            if (state.validLocations.length > 0) {
                const prefs = [
                    getVal(colIndices.pref1),
                    getVal(colIndices.pref2),
                    getVal(colIndices.pref3)
                ];

                prefs.forEach((pref, pIdx) => {
                    if (pref && !state.validLocations.includes(pref)) {
                        addIssue(i, 'error', `存在しない場所名: 第${pIdx + 1}希望`, `"${pref}" はシフト表のヘッダーに存在しません。表記揺れを確認してください。`);
                    }
                });
            }

            // Check E: Date Formats
            // Just check if parsing function returns anything useful if string is present
            if (colIndices.preferredDates !== null) {
                const dateStr = getVal(colIndices.preferredDates);
                if (dateStr) {
                    const { dates: parsed, errors: dateErrors } = parsePreferredDatesWithValidation(dateStr);
                    if (parsed.length === 0 && dateErrors.length === 0) {
                        addIssue(i, 'warning', '希望日の形式を解析できませんでした', `入力値: "${dateStr}" -> 日付として認識されませんでした。`);
                    }
                    // 過去の年エラーを表示
                    dateErrors.forEach(err => {
                        addIssue(i, 'error', err.title, err.description);
                    });
                }
            }
        }

        renderResult();
    }

    function addIssue(rowIndex, level, title, description) {
        state.issues.push({
            row: rowIndex,
            level: level, // 'error' or 'warning' or 'critical'
            title: title,
            description: description
        });
    }

    function getColumnIndices(headers) {
        const indices = {
            participation: null,
            count: null,
            preferredMonth: null,
            preferredDay: null,
            preferredDates: null,
            ngDates: null,
            ngDays: null,
            ngMonths: null,
            pref1: null,
            pref2: null,
            pref3: null,
            additionalSupport: null,
            additionalCount: null,
            grade: null,
            fullName: null,        // 氏名（新形式）
            lastName: null,        // 姓（旧形式）
            firstName: null,       // 名（旧形式）
            classNum: null,
            freeText: null
        };

        headers.forEach((h, i) => {
            if (!h) return;
            const header = h.toString();

            // === NG条件の検出（最優先・絶対条件） ===
            if (header.includes('NG') || header.includes('参加できない') || header.includes('不参加')) {
                if (header.includes('日') && !header.includes('曜日')) indices.ngDates = i;
                else if (header.includes('曜日')) indices.ngDays = i;
                else if (header.includes('月')) indices.ngMonths = i;
                return;
            }

            // === 回数系 ===
            if (header.includes('参加可能回数') || header.includes('希望回数') || (header.includes('回数') && header.includes('希望'))) {
                // 追加回数と区別
                if (!header.includes('追加') && !header.includes('欠員')) {
                    indices.count = i;
                }
            }

            // 参加/免除判定
            if (header.includes('旗振り当番について') || header.includes('参加確認')) {
                indices.participation = i;
            }

            // === 日付・月・曜日 ===
            if (header.includes('特定') || header.includes('希望する日にち') || header.includes('特に') ||
                ((header.includes('希望') || header.includes('参加可能')) && header.includes('日') && !header.includes('曜日'))) {
                if (!header.includes('NG') && !header.includes('参加できない')) {
                    indices.preferredDates = i;
                }
            }
            if (header.includes('月') && (header.includes('希望') || header.includes('選択') || header.includes('参加可能'))) {
                if (!header.includes('NG') && !header.includes('参加できない')) {
                    indices.preferredMonth = i;
                }
            }
            if (header.includes('曜日') && (header.includes('希望') || header.includes('選択') || header.includes('参加可能'))) {
                if (!header.includes('NG') && !header.includes('参加できない')) {
                    indices.preferredDay = i;
                }
            }

            // === 場所の希望 ===
            if (!header.includes('追加') && !header.includes('欠員')) {
                if (header.includes('第1希望') || header.includes('第１希望') || header.includes('第I希望') ||
                    ((header.includes('場所') || header.includes('地点')) && (header.includes('1') || header.includes('１')))) {
                    indices.pref1 = i;
                }
                if (header.includes('第2希望') || header.includes('第２希望') || header.includes('第II希望') ||
                    ((header.includes('場所') || header.includes('地点')) && (header.includes('2') || header.includes('２')))) {
                    indices.pref2 = i;
                }
                if (header.includes('第3希望') || header.includes('第３希望') || header.includes('第III希望') ||
                    ((header.includes('場所') || header.includes('地点')) && (header.includes('3') || header.includes('３')))) {
                    indices.pref3 = i;
                }
            }

            // === 追加対応 ===
            if (header.includes('追加') || header.includes('欠員')) {
                if (header.includes('回数') || header.includes('何回')) {
                    indices.additionalCount = i;
                } else {
                    indices.additionalSupport = i;
                }
            }

            // === 学年・氏名・クラス ===
            if (header.includes('学年')) {
                indices.grade = i;
            }
            // 氏名の検出（新形式: 氏名, 旧形式: 姓・名）
            // 新形式: 「氏名」または「名前」を含むが、「姓」「名」を含まない場合 → fullName
            if ((header.includes('氏名') || header.includes('名前')) && !header.includes('姓') && !header.includes('（名）')) {
                indices.fullName = i;
            }
            // 旧形式: 「姓」「名」別カラム
            else if (header.includes('姓') || header.includes('名字')) {
                indices.lastName = i;
            }
            else if (header.includes('名') && !header.includes('氏名') && !header.includes('校') && !header.includes('宛') && !header.includes('名字')) {
                if (indices.firstName === null) indices.firstName = i;
            }
            if (header.includes('クラス') || header.includes('組')) {
                indices.classNum = i;
            }

            // === 自由記述・備考 ===
            if (header.includes('備考') || header.includes('自由') || header.includes('その他') || header.includes('連絡')) {
                indices.freeText = i;
            }
        });

        return indices;
    }

    function findCol(headers, keywords) {
        for (let i = 0; i < headers.length; i++) {
            if (!headers[i]) continue;
            const h = headers[i].toString();
            if (keywords.some(k => h.includes(k))) return i;
        }
        return null;
    }

    // Copied from main script for consistency
    function parsePreferredDates(str) {
        if (!str) return [];
        const dates = [];
        const strVal = str.toString()
            .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
            .replace(/[、，\s]+/g, ',');

        const segments = strVal.split(',');
        segments.forEach(segment => {
            let match = segment.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/) ||
                segment.match(/(\d{1,2})\/(\d{1,2})/) ||
                segment.match(/(\d{1,2})月(\d{1,2})日/);
            if (match) {
                dates.push(match[0]); // Just push string to say "valid format found"
            }
        });
        return dates;
    }

    // ========================================================================
    // UI Rendering
    // ========================================================================
    function renderResult() {
        resultSection.classList.remove('hidden');
        resultSection.style.display = 'block';

        const errorCount = state.issues.filter(i => i.level === 'error' || i.level === 'critical').length;
        const warningCount = state.issues.filter(i => i.level === 'warning').length;

        // Summary
        validationSummary.innerHTML = `
            <div class="summary-card ${errorCount > 0 ? 'danger' : 'valid'}">
                <div class="summary-number">${errorCount}</div>
                <div class="summary-label">エラー<br>(修正必須)</div>
            </div>
            <div class="summary-card ${warningCount > 0 ? 'warning' : 'valid'}">
                <div class="summary-number">${warningCount}</div>
                <div class="summary-label">警告<br>(要確認)</div>
            </div>
            <div class="summary-card">
                <div class="summary-number">${state.surveyData.length - 1}</div>
                <div class="summary-label">チェック済み<br>行数</div>
            </div>
        `;

        // Default: Show all
        renderIssues('all');
    }

    function renderIssues(filter) {
        issuesList.innerHTML = '';

        const filtered = state.issues.filter(i => {
            if (filter === 'all') return true;
            if (filter === 'error') return i.level === 'error' || i.level === 'critical';
            if (filter === 'warning') return i.level === 'warning';
            return true;
        });

        if (filtered.length === 0) {
            issuesList.innerHTML = `<div style="text-align: center; color: #aaa; padding: 20px;">
                ${state.issues.length === 0 ? '🎉 問題は見つかりませんでした！' : '該当する項目はありません'}
            </div>`;
            return;
        }

        // Group by title
        const grouped = filtered.reduce((acc, issue) => {
            if (!acc[issue.title]) acc[issue.title] = [];
            acc[issue.title].push(issue);
            return acc;
        }, {});

        // Render groups
        Object.keys(grouped).sort().forEach(title => {
            const groupIssues = grouped[title];
            const sampleIssue = groupIssues[0];
            const isWarning = sampleIssue.level === 'warning';

            // Group Container
            const groupDiv = document.createElement('div');
            groupDiv.style.marginBottom = '24px';
            groupDiv.style.border = `1px solid ${isWarning ? '#fcd34d' : '#f87171'}`;
            groupDiv.style.borderRadius = '8px';
            groupDiv.style.overflow = 'hidden';

            // Group Header
            const headerColor = isWarning ? '#fffbeb' : '#fef2f2';
            const badgeColor = isWarning ? '#f59e0b' : '#ef4444';

            groupDiv.innerHTML = `
                <div style="background: ${headerColor}; padding: 12px 16px; border-bottom: 1px solid ${isWarning ? '#fcd34d' : '#f87171'}; display: flex; align-items: center; justify-content: space-between;">
                    <div style="font-weight: bold; color: ${isWarning ? '#92400e' : '#991b1b'}; display: flex; align-items: center;">
                        <span style="background: ${badgeColor}; color: white; padding: 2px 8px; border-radius: 4px; font-size: 0.75rem; margin-right: 8px;">
                            ${isWarning ? 'WARNING' : 'ERROR'}
                        </span>
                        ${title} (${groupIssues.length}件)
                    </div>
                </div>
            `;

            // List of items
            const listDiv = document.createElement('div');
            listDiv.style.padding = '8px 0';
            listDiv.style.backgroundColor = '#fff';

            groupIssues.forEach(issue => {
                const itemDiv = document.createElement('div');
                itemDiv.style.padding = '8px 16px';
                itemDiv.style.borderBottom = '1px solid #eee';

                const rowText = issue.row === 'global' ? '全体' : `行 ${Number(issue.row) + 1}`;

                itemDiv.innerHTML = `
                    <div style="display: flex; align-items: baseline;">
                        <span style="font-family: monospace; font-weight: bold; color: #666; margin-right: 12px; min-width: 60px;">${rowText}</span>
                        <span style="color: #333;">${issue.description}</span>
                    </div>
                `;
                listDiv.appendChild(itemDiv);
            });

            groupDiv.appendChild(listDiv);
            issuesList.appendChild(groupDiv);
        });
    }

    // Expose render for tab switching
    window.renderIssues = renderIssues;

    // ========================================================================
    // Data Export
    // ========================================================================
    const downloadCleanBtn = document.getElementById('downloadCleanBtn');
    if (downloadCleanBtn) {
        downloadCleanBtn.addEventListener('click', exportCleanedData);
    }

    function exportCleanedData() {
        if (!state.surveyData) return;

        const headers = state.surveyData[0];
        const colIndices = getColumnIndices(headers);

        // Filter and Clean
        const cleanData = [headers]; // Always include headers

        for (let i = 1; i < state.surveyData.length; i++) {
            const row = state.surveyData[i];

            // Filter Exemptions
            const participationStr = colIndices.participation !== null && row[colIndices.participation]
                ? row[colIndices.participation].toString()
                : '';

            if (participationStr.includes('免除')) {
                continue; // Skip this row
            }

            const cleanRow = row.map((cell, colIndex) => {
                let val = cell;

                if (typeof val === 'string') {
                    // 1. Trim
                    val = val.trim();

                    // 2. Full-width to Half-width (Numbers & Alpha)
                    val = val.replace(/[０-９Ａ-Ｚａ-ｚ]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));

                    // 3. Normalize spaces
                    val = val.replace(/\s+/g, ' ');
                }

                // 4. Specific Date Column Formatting
                if (colIndex === colIndices.preferredDates && val !== undefined && val !== null && val !== '') {
                    const dates = parsePreferredDates(val);
                    if (dates.length > 0) {
                        return dates.join(','); // Separator is ","
                    }
                }

                return val;
            });

            cleanData.push(cleanRow);
        }

        // Generate Excel
        const ws = XLSX.utils.aoa_to_sheet(cleanData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "CleanedData");

        // Filename with timestamp
        const now = new Date();
        const timestamp = `${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}`;
        XLSX.writeFile(wb, `survey_cleaned_${timestamp}.xlsx`);
    }

    // 日付解析・整形ロジック (年補完・ゼロ埋め・シリアル値対応版)
    function parsePreferredDates(str, baseDate = new Date()) {
        if (!str) return [];

        // Excelシリアル値 (数値のみ) の場合
        if (!isNaN(str) && Number(str) > 20000 && Number(str) < 60000) {
            const serial = Number(str);
            const date = new Date((serial - 25569) * 86400 * 1000);
            return [formatDate(date.getFullYear(), date.getMonth() + 1, date.getDate())];
        }

        // 前処理
        let strVal = str.toString();

        // 1. 全角半角変換 (数字のみ)
        strVal = strVal.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));

        // 2. 曜日の除去: (月), （月）, ㈪ などを空文字に
        // 括弧付きの曜日、丸囲み文字などを削除
        strVal = strVal.replace(/[\(（][月火水木金土日][\)）]/g, '');
        strVal = strVal.replace(/[㈪㈫㈬㈭㈮㈯㈰]/g, '');

        // 3. 区切り文字の統一
        // 年月日などの文字を適切な区切りに置換
        strVal = strVal
            .replace(/年/g, '/')
            .replace(/月/g, '/')
            .replace(/日/g, '');

        // 4. セパレーターの統一
        // カンマ、読点、スペース、ドットなどを全てカンマにする
        strVal = strVal.replace(/[、，\s・\.]+/g, ','); // Added dot and center dot

        strVal = strVal.trim();

        const dates = [];
        const segments = strVal.split(',');

        const currentYear = baseDate.getFullYear();
        // Date比較用に時刻をリセットした基準日を作成
        const base = new Date(currentYear, baseDate.getMonth(), baseDate.getDate());

        segments.forEach(segment => {
            if (!segment.trim()) return;

            // シリアル値文字列のケア
            if (segment.match(/^\d{5}$/)) {
                const serial = Number(segment);
                if (serial > 20000 && serial < 60000) {
                    const date = new Date((serial - 25569) * 86400 * 1000);
                    date.setSeconds(date.getSeconds() + 10);
                    dates.push(formatDate(date.getFullYear(), date.getMonth() + 1, date.getDate()));
                    return;
                }
            }

            // パターン1: YYYY/MM/DD
            let match = segment.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
            if (match) {
                const y = parseInt(match[1]);
                const m = parseInt(match[2]);
                const d = parseInt(match[3]);
                dates.push(formatDate(y, m, d));
                return;
            }

            // パターン2: MM/DD -> 年補完
            match = segment.match(/^(\d{1,2})\/(\d{1,2})$/);
            if (match) {
                const m = parseInt(match[1]);
                const d = parseInt(match[2]);

                let y = currentYear;
                const targetThisYear = new Date(y, m - 1, d);

                if (targetThisYear < base) {
                    y++;
                }

                dates.push(formatDate(y, m, d));
                return;
            }
        });

        return dates;
    }

    function formatDate(y, m, d) {
        // ゼロパディングなしで出力（js/script.jsと一貫性を持たせる）
        return `${y}/${m}/${d}`;
    }

    // 日付解析・検証用（エラーも返す）
    function parsePreferredDatesWithValidation(str, baseDate = new Date()) {
        if (!str) return { dates: [], errors: [] };

        const currentYear = baseDate.getFullYear();
        const errors = [];

        // Excelシリアル値 (数値のみ) の場合
        if (!isNaN(str) && Number(str) > 20000 && Number(str) < 60000) {
            const serial = Number(str);
            const date = new Date((serial - 25569) * 86400 * 1000);
            const y = date.getFullYear();
            if (y < currentYear) {
                errors.push({
                    title: '過去の年が入力されています',
                    description: `入力値: "${str}" → ${y}年は過去の年です。正しい年を入力してください。`
                });
            }
            return { dates: [formatDate(y, date.getMonth() + 1, date.getDate())], errors };
        }

        // 前処理
        let strVal = str.toString();
        strVal = strVal.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
        strVal = strVal.replace(/[\(（][月火水木金土日][\)）]/g, '');
        strVal = strVal.replace(/[㉾㉿㊀㊁㊂㊃㊄]/g, '');
        strVal = strVal
            .replace(/年/g, '/')
            .replace(/月/g, '/')
            .replace(/日/g, '');
        strVal = strVal.replace(/[、，\s・\.]+/g, ',');
        strVal = strVal.trim();

        const dates = [];
        const segments = strVal.split(',');
        const base = new Date(currentYear, baseDate.getMonth(), baseDate.getDate());

        segments.forEach(segment => {
            if (!segment.trim()) return;

            // シリアル値文字列のケア
            if (segment.match(/^\d{5}$/)) {
                const serial = Number(segment);
                if (serial > 20000 && serial < 60000) {
                    const date = new Date((serial - 25569) * 86400 * 1000);
                    date.setSeconds(date.getSeconds() + 10);
                    const y = date.getFullYear();
                    if (y < currentYear) {
                        errors.push({
                            title: '過去の年が入力されています',
                            description: `入力値: "${segment}" → ${y}年は過去の年です。正しい年を入力してください。`
                        });
                    }
                    dates.push(formatDate(y, date.getMonth() + 1, date.getDate()));
                    return;
                }
            }

            // パターン1: YYYY/MM/DD
            let match = segment.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
            if (match) {
                const y = parseInt(match[1]);
                const m = parseInt(match[2]);
                const d = parseInt(match[3]);
                if (y < currentYear) {
                    errors.push({
                        title: '過去の年が入力されています',
                        description: `入力値: "${segment}" → ${y}年は過去の年です。正しい年を入力してください。`
                    });
                }
                dates.push(formatDate(y, m, d));
                return;
            }

            // パターン2: MM/DD -> 年補完
            match = segment.match(/^(\d{1,2})\/(\d{1,2})$/);
            if (match) {
                const m = parseInt(match[1]);
                const d = parseInt(match[2]);
                let y = currentYear;
                const targetThisYear = new Date(y, m - 1, d);
                if (targetThisYear < base) {
                    y++;
                }
                dates.push(formatDate(y, m, d));
                return;
            }
        });

        return { dates, errors };
    }
});
