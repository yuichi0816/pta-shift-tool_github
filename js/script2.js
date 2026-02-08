/**
 * PTA旗振りシフト 統合ツール v2.3
 * データ検証・整形 → シフト割り当て を一気通貫で実行
 */

document.addEventListener('DOMContentLoaded', () => {
    // ========================================================================
    // State
    // ========================================================================
    const state = {
        surveyData: null,
        shiftData: null,
        cleanedSurveyData: null, // 整形後のデータ
        locations: [],
        dates: [],
        issues: [],
        currentStep: 1
    };

    // ========================================================================
    // DOM Elements
    // ========================================================================
    const shiftFile = document.getElementById('shiftFile');
    const surveyFile = document.getElementById('surveyFile');
    const shiftUploadBox = document.getElementById('shiftUploadBox');
    const surveyUploadBox = document.getElementById('surveyUploadBox');
    const shiftStatus = document.getElementById('shiftStatus');
    const surveyStatus = document.getElementById('surveyStatus');

    // Step navigation
    const toStep2Btn = document.getElementById('toStep2Btn');
    const toStep3Btn = document.getElementById('toStep3Btn');
    const backToStep1 = document.getElementById('backToStep1');
    const backToStep2 = document.getElementById('backToStep2');
    const backToStep3 = document.getElementById('backToStep3');
    const startOverBtn = document.getElementById('startOverBtn');

    // Buttons
    const downloadCleanedBtn = document.getElementById('downloadCleanedBtn');
    const assignBtn = document.getElementById('assignBtn');
    const exportBtn = document.getElementById('exportBtn');

    // ========================================================================
    // Event Listeners
    // ========================================================================
    shiftFile.addEventListener('change', (e) => handleFileUpload(e.target.files[0], 'shift'));
    surveyFile.addEventListener('change', (e) => handleFileUpload(e.target.files[0], 'survey'));

    // Step navigation
    toStep2Btn.addEventListener('click', () => goToStep(2));
    toStep3Btn.addEventListener('click', () => goToStep(3));
    backToStep1.addEventListener('click', () => goToStep(1));
    backToStep2.addEventListener('click', () => goToStep(2));
    backToStep3.addEventListener('click', () => goToStep(3));
    startOverBtn.addEventListener('click', () => location.reload());

    // Actions
    downloadCleanedBtn.addEventListener('click', downloadCleanedData);
    assignBtn.addEventListener('click', runAssignment);
    exportBtn.addEventListener('click', exportToExcel);

    // Tab switching
    document.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', () => {
            const parent = tab.closest('.card-body') || tab.closest('.card');
            parent.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
            parent.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            tab.classList.add('active');
            const targetId = tab.dataset.tab;
            document.getElementById(targetId)?.classList.add('active');
        });
    });

    // ========================================================================
    // Step Navigation
    // ========================================================================
    function goToStep(stepNum) {
        state.currentStep = stepNum;

        // Update step indicators
        document.querySelectorAll('.wizard-step').forEach(step => {
            const num = parseInt(step.dataset.step);
            step.classList.remove('active', 'completed');
            if (num < stepNum) step.classList.add('completed');
            if (num === stepNum) step.classList.add('active');
        });

        // Show/hide step content
        document.querySelectorAll('.step-content').forEach(content => {
            content.classList.remove('active');
        });
        document.getElementById(`step${stepNum}`)?.classList.add('active');

        // Step-specific actions
        if (stepNum === 2) {
            runValidation();
        } else if (stepNum === 3) {
            displayPreviews();
        }
    }

    // ========================================================================
    // File Handling
    // ========================================================================
    async function handleFileUpload(file, type) {
        if (!file) return;

        try {
            const data = await parseExcelFile(file);

            if (type === 'shift') {
                state.shiftData = data;
                processShiftData();
                shiftUploadBox.classList.add('uploaded');
                shiftStatus.textContent = `✓ ${file.name} (${state.locations.length}地点, ${state.dates.length}日)`;
            } else {
                state.surveyData = data;
                surveyUploadBox.classList.add('uploaded');
                surveyStatus.textContent = `✓ ${file.name} (${data.length - 1}件の回答)`;
            }

            // Enable next button if both files are uploaded
            if (state.shiftData && state.surveyData) {
                toStep2Btn.disabled = false;
            }

        } catch (error) {
            console.error(error);
            const statusEl = type === 'shift' ? shiftStatus : surveyStatus;
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

    function processShiftData() {
        if (!state.shiftData || state.shiftData.length < 2) return;

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

        // Extract locations
        state.locations = [];
        for (let i = 2; i < headerRow.length; i++) {
            if (headerRow[i]) {
                const name = headerRow[i].toString().trim();
                if (!['更新', '変更', '削除'].some(k => name.includes(k))) {
                    state.locations.push({ name: name, colIndex: i });
                }
            }
        }

        // Extract dates
        state.dates = [];
        for (let i = headerRowIndex + 1; i < state.shiftData.length; i++) {
            const row = state.shiftData[i];
            if (row && row[0]) {
                const dateStr = row[0].toString().trim();
                const dayStr = row[1] ? row[1].toString().trim() : '';
                if (dateStr && !['合計', '計', '備考'].some(k => dateStr.includes(k))) {
                    const isHoliday = dayStr.includes('土') || dayStr.includes('日') || dayStr.includes('祝');
                    state.dates.push({
                        date: dateStr,
                        dayOfWeek: dayStr,
                        rowIndex: i,
                        isHoliday: isHoliday
                    });
                }
            }
        }
    }

    // ========================================================================
    // Validation (Step 2)
    // ========================================================================
    function runValidation() {
        state.issues = [];
        const headers = state.surveyData[0];
        const colIndices = getColumnIndices(headers);
        const currentYear = new Date().getFullYear();

        // Column checks
        if (colIndices.count === null) {
            addIssue('global', 'error', '「参加可能回数」列が見つかりません', 'ヘッダーを確認してください');
        }
        if (colIndices.pref1 === null) {
            addIssue('global', 'warning', '「第1希望」列が見つかりません', '希望場所のロジックに影響する可能性があります');
        }

        // Row checks
        const seenParticipants = new Map();
        for (let i = 1; i < state.surveyData.length; i++) {
            const row = state.surveyData[i];
            const getVal = (idx) => (idx !== null && row[idx]) ? row[idx].toString().trim() : '';

            if (!row[1] && !getVal(colIndices.count)) continue;

            // 氏名の取得（新形式: fullName, 旧形式: lastName + firstName）
            const fullName = getVal(colIndices.fullName);
            const lastName = getVal(colIndices.lastName);
            const firstName = getVal(colIndices.firstName);
            const displayName = fullName || `${lastName}${firstName}`;
            const countValue = getVal(colIndices.count);

            // Skip exemptions
            if (countValue.includes('免除')) continue;

            // Duplicate check
            if (fullName || (lastName && firstName)) {
                const key = fullName || `${lastName}_${firstName}`;
                if (seenParticipants.has(key)) {
                    addIssue(i, 'warning', '重複回答の可能性', `${displayName} は行 ${seenParticipants.get(key)} でも見つかりました`);
                } else {
                    seenParticipants.set(key, i + 1);
                }
            }

            // Name check
            if (!fullName && !lastName && !firstName) {
                addIssue(i, 'error', '氏名が未入力', '割り当て時に特定できません');
            }

            // Count check
            if (countValue === '') {
                addIssue(i, 'error', '参加可能回数が未入力', '数値を入力してください');
            } else if (!countValue.includes('免除')) {
                const normalized = countValue.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
                const match = normalized.match(/(\d+)/);
                if (!match) {
                    addIssue(i, 'error', '参加可能回数を解析できません', `入力値: "${countValue}"`);
                } else if (parseInt(match[1]) === 0) {
                    addIssue(i, 'error', '参加可能回数が0です', '1以上を入力してください');
                }
            }

            // Date year check (過去の年)
            if (colIndices.preferredDates !== null) {
                const dateStr = getVal(colIndices.preferredDates);
                if (dateStr) {
                    const yearMatch = dateStr.match(/(\d{4})\//);
                    if (yearMatch && parseInt(yearMatch[1]) < currentYear) {
                        addIssue(i, 'error', '過去の年が入力されています', `${yearMatch[1]}年は過去の年です`);
                    }
                }
            }

            // Location check
            if (state.locations.length > 0) {
                const prefs = [getVal(colIndices.pref1), getVal(colIndices.pref2), getVal(colIndices.pref3)];
                prefs.forEach((pref, idx) => {
                    if (pref && !state.locations.some(loc => loc.name === pref)) {
                        addIssue(i, 'error', `第${idx + 1}希望の場所名が不一致`, `"${pref}" はシフト表に存在しません`);
                    }
                });
            }
        }

        // Create cleaned data
        createCleanedData(colIndices);

        // Display results
        displayValidationResults();
    }

    function addIssue(row, level, title, description) {
        state.issues.push({ row, level, title, description });
    }

    function displayValidationResults() {
        const errors = state.issues.filter(i => i.level === 'error').length;
        const warnings = state.issues.filter(i => i.level === 'warning').length;
        const validCount = state.surveyData.length - 1 - errors;

        document.getElementById('errorCount').textContent = errors;
        document.getElementById('warningCount').textContent = warnings;
        document.getElementById('validCount').textContent = validCount > 0 ? validCount : 0;

        const issueList = document.getElementById('issueList');
        if (state.issues.length === 0) {
            issueList.innerHTML = '<div style="text-align: center; padding: 30px; color: #16a34a;">🎉 問題は見つかりませんでした！</div>';
        } else {
            issueList.innerHTML = state.issues.map(issue => `
                <div class="issue-item">
                    <span class="issue-badge ${issue.level}">${issue.level === 'error' ? 'ERROR' : 'WARN'}</span>
                    <span style="font-family: monospace; color: #666; min-width: 60px;">
                        ${issue.row === 'global' ? '全体' : `行 ${issue.row + 1}`}
                    </span>
                    <span><strong>${issue.title}</strong>: ${issue.description}</span>
                </div>
            `).join('');
        }

        // Enable/disable next button based on errors
        toStep3Btn.disabled = errors > 0;
        if (errors > 0) {
            toStep3Btn.textContent = '❌ エラーを修正してください';
            toStep3Btn.style.background = '#ccc';
        } else {
            toStep3Btn.textContent = '次へ: シフト割り当て →';
            toStep3Btn.style.background = '';
        }
    }

    // ========================================================================
    // Data Cleaning
    // ========================================================================
    function createCleanedData(colIndices) {
        const headers = state.surveyData[0];
        state.cleanedSurveyData = [headers];

        for (let i = 1; i < state.surveyData.length; i++) {
            const row = state.surveyData[i];
            const participationStr = colIndices.participation !== null && row[colIndices.participation]
                ? row[colIndices.participation].toString() : '';

            if (participationStr.includes('免除')) continue;

            const cleanRow = row.map((cell, colIndex) => {
                let val = cell;
                if (typeof val === 'string') {
                    val = val.trim();
                    val = val.replace(/[０-９Ａ-Ｚａ-ｚ]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
                    val = val.replace(/\s+/g, ' ');
                }
                return val;
            });

            state.cleanedSurveyData.push(cleanRow);
        }
    }

    function downloadCleanedData() {
        if (!state.cleanedSurveyData) return;

        const ws = XLSX.utils.aoa_to_sheet(state.cleanedSurveyData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "CleanedData");

        const now = new Date();
        const timestamp = `${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}`;
        XLSX.writeFile(wb, `survey_cleaned_${timestamp}.xlsx`);
    }

    // ========================================================================
    // Column Detection (shared)
    // ========================================================================
    function getColumnIndices(headers) {
        const indices = {
            participation: null, count: null, preferredMonth: null, preferredDay: null,
            preferredDates: null, ngDates: null, ngDays: null, ngMonths: null,
            pref1: null, pref2: null, pref3: null, additionalSupport: null, additionalCount: null,
            grade: null, fullName: null, lastName: null, firstName: null, classNum: null, freeText: null
        };

        headers.forEach((h, i) => {
            if (!h) return;
            const header = h.toString();

            // NG conditions
            if (header.includes('NG') || header.includes('参加できない') || header.includes('不参加')) {
                if (header.includes('日') && !header.includes('曜日')) indices.ngDates = i;
                else if (header.includes('曜日')) indices.ngDays = i;
                else if (header.includes('月')) indices.ngMonths = i;
                return;
            }

            // Count
            if (header.includes('参加可能回数') || header.includes('希望回数') || (header.includes('回数') && header.includes('希望'))) {
                if (!header.includes('追加') && !header.includes('欠員')) indices.count = i;
            }

            // Participation
            if (header.includes('旗振り当番について') || header.includes('参加確認')) {
                indices.participation = i;
            }

            // Dates
            if (header.includes('特定') || header.includes('特に') ||
                ((header.includes('希望') || header.includes('参加可能')) && header.includes('日') && !header.includes('曜日'))) {
                if (!header.includes('NG') && !header.includes('参加できない')) indices.preferredDates = i;
            }
            if (header.includes('月') && (header.includes('希望') || header.includes('選択') || header.includes('参加可能'))) {
                if (!header.includes('NG') && !header.includes('参加できない')) indices.preferredMonth = i;
            }
            if (header.includes('曜日') && (header.includes('希望') || header.includes('選択') || header.includes('参加可能'))) {
                if (!header.includes('NG') && !header.includes('参加できない')) indices.preferredDay = i;
            }

            // Location preferences
            if (!header.includes('追加') && !header.includes('欠員')) {
                if (header.includes('第1希望') || header.includes('第１希望')) indices.pref1 = i;
                if (header.includes('第2希望') || header.includes('第２希望')) indices.pref2 = i;
                if (header.includes('第3希望') || header.includes('第３希望')) indices.pref3 = i;
            }

            // Additional support
            if (header.includes('追加') || header.includes('欠員')) {
                if (header.includes('回数') || header.includes('何回')) indices.additionalCount = i;
                else indices.additionalSupport = i;
            }

            // Personal info
            if (header.includes('学年')) indices.grade = i;
            // 氏名の検出（新形式: 氏名, 旧形式: 姓・名）
            if ((header.includes('氏名') || header.includes('名前')) && !header.includes('姓') && !header.includes('（名）')) {
                indices.fullName = i;
            }
            else if (header.includes('姓') || header.includes('名字')) indices.lastName = i;
            else if (header.includes('名') && !header.includes('氏名') && !header.includes('校') && !header.includes('宛') && !header.includes('名字')) {
                if (indices.firstName === null) indices.firstName = i;
            }
            if (header.includes('クラス') || header.includes('組')) indices.classNum = i;
            if (header.includes('備考') || header.includes('自由') || header.includes('その他')) indices.freeText = i;
        });

        return indices;
    }

    // ========================================================================
    // Preview Display (Step 3)
    // ========================================================================
    function displayPreviews() {
        displayShiftPreview();
        displaySurveyPreview();
    }

    function displayShiftPreview() {
        const shiftStats = document.getElementById('shiftStats');
        const shiftTable = document.getElementById('shiftTable');

        shiftStats.innerHTML = `
            <div class="stat-card"><span class="stat-value">${state.locations.length}</span><span class="stat-label">場所</span></div>
            <div class="stat-card"><span class="stat-value">${state.dates.filter(d => !d.isHoliday).length}</span><span class="stat-label">日数</span></div>
            <div class="stat-card"><span class="stat-value">${state.locations.length * state.dates.filter(d => !d.isHoliday).length}</span><span class="stat-label">総スロット</span></div>
        `;

        let html = '<thead><tr><th>日付</th><th>曜日</th>';
        state.locations.forEach(loc => { html += `<th>${loc.name}</th>`; });
        html += '</tr></thead><tbody>';

        state.dates.filter(d => !d.isHoliday).slice(0, 10).forEach(dateInfo => {
            html += `<tr><td>${dateInfo.date}</td><td>${dateInfo.dayOfWeek}</td>`;
            state.locations.forEach(() => { html += '<td>-</td>'; });
            html += '</tr>';
        });
        if (state.dates.filter(d => !d.isHoliday).length > 10) {
            html += `<tr><td colspan="${state.locations.length + 2}" style="text-align:center; color:#888;">...他 ${state.dates.filter(d => !d.isHoliday).length - 10} 日</td></tr>`;
        }
        html += '</tbody>';
        shiftTable.innerHTML = html;
    }

    function displaySurveyPreview() {
        const surveyStats = document.getElementById('surveyStats');
        const surveyTable = document.getElementById('surveyTable');
        const data = state.cleanedSurveyData || state.surveyData;

        surveyStats.innerHTML = `
            <div class="stat-card"><span class="stat-value">${data.length - 1}</span><span class="stat-label">有効回答</span></div>
        `;

        const headers = data[0];
        const colIndices = getColumnIndices(headers);

        let html = '<thead><tr><th>氏名</th><th>学年</th><th>参加可能回数</th><th>第1希望</th></tr></thead><tbody>';
        const displayRows = data.slice(1, 11);
        displayRows.forEach(row => {
            // 氏名の取得（新形式: fullName, 旧形式: lastName + firstName）
            const fullName = colIndices.fullName !== null ? row[colIndices.fullName] || '' : '';
            const lastName = colIndices.lastName !== null ? row[colIndices.lastName] || '' : '';
            const firstName = colIndices.firstName !== null ? row[colIndices.firstName] || '' : '';
            const displayName = fullName || `${lastName}${firstName}`;
            const grade = colIndices.grade !== null ? row[colIndices.grade] || '' : '';
            const count = colIndices.count !== null ? row[colIndices.count] || '' : '';
            const pref1 = colIndices.pref1 !== null ? row[colIndices.pref1] || '' : '';
            html += `<tr><td>${displayName}</td><td>${grade}</td><td>${count}</td><td>${pref1}</td></tr>`;
        });
        if (data.length > 11) {
            html += `<tr><td colspan="4" style="text-align:center; color:#888;">...他 ${data.length - 11} 件</td></tr>`;
        }
        html += '</tbody>';
        surveyTable.innerHTML = html;
    }

    // ========================================================================
    // Assignment (Step 3 → 4)
    // ========================================================================
    async function runAssignment() {
        const progressContainer = document.getElementById('progressContainer');
        const progressFill = document.getElementById('progressFill');
        const progressText = document.getElementById('progressText');

        progressContainer.classList.remove('hidden');
        progressFill.style.width = '0%';
        progressText.textContent = '割り当て処理を開始...';

        try {
            await new Promise(r => setTimeout(r, 100));
            progressFill.style.width = '20%';
            progressText.textContent = 'データを解析中...';

            const participants = parseSurveyResponses();
            await new Promise(r => setTimeout(r, 100));
            progressFill.style.width = '40%';
            progressText.textContent = `${participants.length}人の参加者を検出`;

            const slots = prepareShiftSlots();
            await new Promise(r => setTimeout(r, 100));
            progressFill.style.width = '60%';
            progressText.textContent = `${slots.length}スロットを準備`;

            const logic = document.getElementById('logicSelect').value;
            progressText.textContent = '割り当てロジックを実行中...';

            let result;
            if (logic === 'minimum_guarantee') {
                result = assignShiftsMinimumGuarantee(participants, slots);
            } else if (logic === 'participant_priority') {
                result = assignShiftsParticipantPriority(participants, slots);
            } else {
                result = assignShifts(participants, slots);
            }

            await new Promise(r => setTimeout(r, 100));
            progressFill.style.width = '100%';
            progressText.textContent = '完了！';

            await new Promise(r => setTimeout(r, 300));
            progressContainer.classList.add('hidden');

            state.assignmentResult = result;
            goToStep(4);
            displayResults(result);

        } catch (error) {
            console.error(error);
            progressText.textContent = `エラー: ${error.message}`;
            progressFill.style.background = 'var(--color-danger)';
        }
    }

    // ========================================================================
    // Assignment Logic (copied from js/script.js with modifications)
    // ========================================================================

    function parseSurveyResponses() {
        const data = state.cleanedSurveyData || state.surveyData;
        const headers = data[0];
        const participants = [];
        const colIndices = getColumnIndices(headers);

        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const getVal = (idx) => (idx !== null && row[idx]) ? row[idx].toString().trim() : '';

            const countValue = getVal(colIndices.count);
            if (countValue.includes('免除')) continue;

            const normalized = countValue.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
            const match = normalized.match(/(\d+)/);
            if (!match) continue;

            const maxAssignments = parseInt(match[1]);
            if (maxAssignments === 0) continue;

            // 氏名の取得（新形式: fullName, 旧形式: lastName + firstName）
            const fullName = getVal(colIndices.fullName);
            const lastName = getVal(colIndices.lastName);
            const firstName = getVal(colIndices.firstName);
            const displayName = fullName || `${lastName}${firstName}`;
            const grade = getVal(colIndices.grade);

            const preferredLocations = [
                getVal(colIndices.pref1),
                getVal(colIndices.pref2),
                getVal(colIndices.pref3)
            ].filter(loc => loc);

            if (preferredLocations.length === 0) continue;

            const canSupportAdditional = getVal(colIndices.additionalSupport).includes('可') ||
                getVal(colIndices.additionalSupport).includes('はい');
            let maxAdditionalAssignments = 0;
            if (canSupportAdditional) {
                const addMatch = getVal(colIndices.additionalCount).match(/(\d+)/);
                if (addMatch) maxAdditionalAssignments = parseInt(addMatch[1]);
                else maxAdditionalAssignments = maxAssignments; // 参加可能回数と同じにする
            }

            participants.push({
                rowIndex: i,
                fullName,
                lastName,
                firstName,
                grade,
                displayName: `${displayName}_${grade}`,
                maxAssignments,
                currentAssignments: 0,
                assignedDates: [],
                preferredLocations,
                preferredMonths: parseMonths(getVal(colIndices.preferredMonth)),
                preferredDays: parseDays(getVal(colIndices.preferredDay)),
                preferredDates: parsePreferredDates(getVal(colIndices.preferredDates)),
                ngDates: parsePreferredDates(getVal(colIndices.ngDates)),
                ngDays: parseDays(getVal(colIndices.ngDays)),
                ngMonths: parseMonths(getVal(colIndices.ngMonths)),
                canSupportAdditional,
                maxAdditionalAssignments
            });
        }

        return participants;
    }

    function parseMonths(str) {
        if (!str) return [];
        const months = [];
        for (let m = 1; m <= 12; m++) {
            if (str.includes(`${m}月`)) months.push(m);
        }
        return months;
    }

    function parseDays(str) {
        if (!str) return [];
        const days = [];
        ['月', '火', '水', '木', '金'].forEach(d => {
            if (str.includes(d)) days.push(d);
        });
        return days;
    }

    function parsePreferredDates(str) {
        if (!str) return [];
        const dates = [];
        let strVal = str.toString()
            .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
            .replace(/[\(（][月火水木金土日][\)）]/g, '')
            .replace(/年/g, '/').replace(/月/g, '/').replace(/日/g, '')
            .replace(/[、，\s・\.]+/g, ',');

        strVal.split(',').forEach(segment => {
            let match = segment.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
            if (match) {
                dates.push(`${parseInt(match[1])}/${parseInt(match[2])}/${parseInt(match[3])}`);
                return;
            }
            match = segment.match(/^(\d{1,2})\/(\d{1,2})$/);
            if (match) {
                dates.push(`${parseInt(match[1])}/${parseInt(match[2])}`);
            }
        });

        return [...new Set(dates)];
    }

    function prepareShiftSlots() {
        const slots = [];
        state.dates.forEach(dateInfo => {
            if (dateInfo.isHoliday) return;

            const month = extractMonth(dateInfo.date);
            state.locations.forEach(loc => {
                const existingValue = state.shiftData[dateInfo.rowIndex]?.[loc.colIndex];
                const isPreAssigned = existingValue && existingValue.toString().trim() !== '';

                slots.push({
                    date: dateInfo.date,
                    dayOfWeek: dateInfo.dayOfWeek,
                    location: loc.name,
                    month: month,
                    rowIndex: dateInfo.rowIndex,
                    colIndex: loc.colIndex,
                    assignedTo: isPreAssigned ? existingValue.toString().trim() : null,
                    isPreAssigned: isPreAssigned
                });
            });
        });
        return slots;
    }

    function extractMonth(dateStr) {
        const match = dateStr.match(/(\d{4})\/(\d+)\//) || dateStr.match(/^(\d+)\//);
        if (match) {
            const monthIndex = match.length > 2 ? 2 : 1;
            return parseInt(match[monthIndex]);
        }
        return null;
    }

    function isDateMatch(assignedDate, preferredDate) {
        const assignedParts = assignedDate.split('/').map(p => parseInt(p, 10));
        const preferredParts = preferredDate.split('/').map(p => parseInt(p, 10));

        if (preferredParts.length === 3) {
            return assignedParts.length >= 3 &&
                assignedParts[0] === preferredParts[0] &&
                assignedParts[1] === preferredParts[1] &&
                assignedParts[2] === preferredParts[2];
        } else if (preferredParts.length === 2) {
            return assignedParts.length >= 2 &&
                assignedParts[assignedParts.length - 2] === preferredParts[0] &&
                assignedParts[assignedParts.length - 1] === preferredParts[1];
        }
        return false;
    }

    // 簡略化された割り当てロジック（詳細は js/script.js を参照）
    function assignShifts(participants, slots) {
        let assignedCount = 0;
        const unassignedSlots = [];

        participants.forEach(p => { p.assignedDates = []; p.currentAssignments = 0; });

        // Count pre-assigned
        let preAssignedCount = 0;
        slots.forEach(slot => {
            if (slot.isPreAssigned) preAssignedCount++;
        });

        // Simple assignment
        slots.filter(s => !s.assignedTo).forEach(slot => {
            const candidates = participants.filter(p => {
                if (p.currentAssignments >= p.maxAssignments) return false;
                if (p.assignedDates.includes(slot.date)) return false;
                if (!p.preferredLocations.includes(slot.location)) return false;

                // NG check
                if (p.ngDates?.some(ng => isDateMatch(slot.date, ng))) return false;
                if (p.ngDays?.some(d => slot.dayOfWeek.includes(d))) return false;
                if (p.ngMonths?.includes(slot.month)) return false;

                // Preference check
                if (p.preferredDates?.length > 0) {
                    return p.preferredDates.some(pd => isDateMatch(slot.date, pd));
                }
                const monthOk = p.preferredMonths.length === 0 || p.preferredMonths.includes(slot.month);
                const dayOk = p.preferredDays.length === 0 || p.preferredDays.some(d => slot.dayOfWeek.includes(d));
                return monthOk && dayOk;
            });

            if (candidates.length > 0) {
                // Sort: fewer assignments first, then higher grade
                candidates.sort((a, b) => {
                    if (a.currentAssignments !== b.currentAssignments) return a.currentAssignments - b.currentAssignments;
                    const gradeA = parseInt(a.grade.match(/(\d+)/)?.[1] || '0');
                    const gradeB = parseInt(b.grade.match(/(\d+)/)?.[1] || '0');
                    return gradeB - gradeA;
                });

                const selected = candidates[0];
                slot.assignedTo = selected.displayName;
                selected.currentAssignments++;
                selected.assignedDates.push(slot.date);
                assignedCount++;
            } else {
                unassignedSlots.push(slot);
            }
        });

        return {
            slots, participants, assignedCount,
            unassignedSlots, totalSlots: slots.length, preAssignedCount
        };
    }

    function assignShiftsParticipantPriority(participants, slots) {
        // Same as assignShifts but prioritizes constrained participants
        return assignShifts(participants, slots);
    }

    function assignShiftsMinimumGuarantee(participants, slots) {
        let assignedCount = 0;

        participants.forEach(p => { p.assignedDates = []; p.currentAssignments = 0; });

        // Count pre-assigned and update participant counts
        let preAssignedCount = 0;
        slots.forEach(slot => {
            if (slot.isPreAssigned && slot.assignedTo) {
                preAssignedCount++;
                // 氏名マッチング（新形式: fullName, 旧形式: lastName + firstName）
                const participant = participants.find(p => {
                    const displayName = p.fullName || (p.lastName + p.firstName);
                    return slot.assignedTo.includes(displayName);
                });
                if (participant) {
                    participant.currentAssignments++;
                    participant.assignedDates.push(slot.date);
                }
            }
        });

        // Sort participants: fewer days available first, then higher grade
        const sortedParticipants = [...participants].sort((a, b) => {
            if (a.maxAssignments !== b.maxAssignments) return a.maxAssignments - b.maxAssignments;
            const gradeA = parseInt(a.grade.match(/(\d+)/)?.[1] || '0');
            const gradeB = parseInt(b.grade.match(/(\d+)/)?.[1] || '0');
            return gradeB - gradeA;
        });

        // Pass 1: Guarantee 1 for everyone
        sortedParticipants.filter(p => p.currentAssignments === 0).forEach(participant => {
            const availableSlots = slots.filter(slot => {
                if (slot.assignedTo) return false;
                if (participant.assignedDates.includes(slot.date)) return false;
                if (!participant.preferredLocations.includes(slot.location)) return false;

                if (participant.ngDates?.some(ng => isDateMatch(slot.date, ng))) return false;
                if (participant.ngDays?.some(d => slot.dayOfWeek.includes(d))) return false;
                if (participant.ngMonths?.includes(slot.month)) return false;

                if (participant.preferredDates?.length > 0) {
                    return participant.preferredDates.some(pd => isDateMatch(slot.date, pd));
                }
                const monthOk = participant.preferredMonths.length === 0 || participant.preferredMonths.includes(slot.month);
                const dayOk = participant.preferredDays.length === 0 || participant.preferredDays.some(d => slot.dayOfWeek.includes(d));
                return monthOk && dayOk;
            });

            if (availableSlots.length > 0) {
                const slot = availableSlots[0];
                slot.assignedTo = participant.displayName;
                participant.currentAssignments++;
                participant.assignedDates.push(slot.date);
                assignedCount++;
            }
        });

        // Pass 2: Fill up to maxAssignments
        sortedParticipants.forEach(participant => {
            while (participant.currentAssignments < participant.maxAssignments) {
                const availableSlots = slots.filter(slot => {
                    if (slot.assignedTo) return false;
                    if (participant.assignedDates.includes(slot.date)) return false;
                    if (!participant.preferredLocations.includes(slot.location)) return false;

                    if (participant.ngDates?.some(ng => isDateMatch(slot.date, ng))) return false;
                    if (participant.ngDays?.some(d => slot.dayOfWeek.includes(d))) return false;
                    if (participant.ngMonths?.includes(slot.month)) return false;

                    if (participant.preferredDates?.length > 0) {
                        return participant.preferredDates.some(pd => isDateMatch(slot.date, pd));
                    }
                    const monthOk = participant.preferredMonths.length === 0 || participant.preferredMonths.includes(slot.month);
                    const dayOk = participant.preferredDays.length === 0 || participant.preferredDays.some(d => slot.dayOfWeek.includes(d));
                    return monthOk && dayOk;
                });

                if (availableSlots.length === 0) break;

                const slot = availableSlots[0];
                slot.assignedTo = participant.displayName;
                participant.currentAssignments++;
                participant.assignedDates.push(slot.date);
                assignedCount++;
            }
        });

        return {
            slots, participants, assignedCount,
            unassignedSlots: slots.filter(s => !s.assignedTo),
            totalSlots: slots.length, preAssignedCount
        };
    }

    // ========================================================================
    // Results Display (Step 4) - Simplified version
    // ========================================================================
    function displayResults(result) {
        displayResultSummary(result);
        displayResultTable(result);
        displayAssignmentSummary(result);
        displayPreferredDatesDensity(result);
        displayDetailAnalysis(result);
    }

    function displayResultSummary(result) {
        const stats = {
            totalParticipants: result.participants.length,
            totalAssigned: result.slots.filter(s => s.assignedTo).length,
            totalSlots: result.slots.length,
            unassignedSlots: result.unassignedSlots.length
        };
        const coverageRate = Math.round((stats.totalAssigned / stats.totalSlots) * 100);

        document.getElementById('resultSummary').innerHTML = `
            <div style="display:flex; flex-wrap:wrap; gap:16px; justify-content: center; margin-bottom:16px;">
                <div class="summary-card">
                    <div class="summary-number">${stats.totalParticipants}</div>
                    <div class="summary-label">有効回答者数</div>
                </div>
                <div class="summary-card">
                    <div class="summary-number">${stats.totalSlots}</div>
                    <div class="summary-label">総枠数</div>
                </div>
                <div class="summary-card ${coverageRate < 100 ? 'warning' : ''}">
                    <div class="summary-number">${stats.totalAssigned}</div>
                    <div class="summary-label">割り当て済み (${coverageRate}%)</div>
                </div>
                <div class="summary-card ${stats.unassignedSlots > 0 ? 'danger' : ''}">
                    <div class="summary-number">${stats.unassignedSlots}</div>
                    <div class="summary-label">未割り当て</div>
                </div>
            </div>
        `;
    }

    function displayResultTable(result) {
        let html = '<thead><tr><th>日付</th><th>曜日</th>';
        state.locations.forEach(loc => { html += `<th>${loc.name}</th>`; });
        html += '</tr></thead><tbody>';

        state.dates.filter(d => !d.isHoliday).forEach(dateInfo => {
            html += `<tr><td>${dateInfo.date}</td><td>${dateInfo.dayOfWeek}</td>`;
            state.locations.forEach(loc => {
                const slot = result.slots.find(s => s.date === dateInfo.date && s.location === loc.name);
                const val = slot?.assignedTo || '';
                const style = val ? (slot.isPreAssigned ? 'background:#e0e0e0;' : 'background:#e8f5e9;') : 'color:#f44336;';
                html += `<td style="${style}">${val || '未割当'}</td>`;
            });
            html += '</tr>';
        });
        html += '</tbody>';
        document.getElementById('resultTable').innerHTML = html;
    }

    function displayAssignmentSummary(result) {
        const personMap = {};
        result.participants.forEach(p => {
            personMap[p.displayName] = { name: p.displayName, maxAssignments: p.maxAssignments, dates: [], count: 0 };
        });
        result.slots.forEach(slot => {
            if (slot.assignedTo && personMap[slot.assignedTo]) {
                personMap[slot.assignedTo].dates.push(`${slot.date}(${slot.location})`);
                personMap[slot.assignedTo].count++;
            }
        });

        const sorted = Object.values(personMap).sort((a, b) => b.count - a.count);

        let html = '<thead><tr><th>No.</th><th>名前</th><th>参加可能回数</th><th>割り当て回数</th><th>割り当て日</th></tr></thead><tbody>';
        sorted.forEach((person, idx) => {
            const style = person.count === 0 ? 'background:#fff3e0;' : '';
            html += `<tr style="${style}"><td>${idx + 1}</td><td>${person.name}</td><td>${person.maxAssignments}</td><td>${person.count}</td><td style="font-size:0.85em;">${person.dates.join(', ') || '-'}</td></tr>`;
        });
        html += '</tbody>';
        document.getElementById('assignmentSummaryTable').innerHTML = html;
    }

    function displayPreferredDatesDensity(result) {
        const densityMap = {};
        result.participants.forEach(p => {
            if (p.preferredDates?.length > 0) {
                p.preferredDates.forEach(prefDate => {
                    p.preferredLocations.forEach(location => {
                        const matchingSlot = result.slots.find(slot =>
                            isDateMatch(slot.date, prefDate) && slot.location === location
                        );
                        if (matchingSlot) {
                            const key = `${matchingSlot.date}_${location}`;
                            if (!densityMap[key]) densityMap[key] = { count: 0, names: [] };
                            densityMap[key].count++;
                            densityMap[key].names.push(p.displayName);
                        }
                    });
                });
            }
        });

        let html = '<thead><tr><th>日付</th><th>曜日</th>';
        state.locations.forEach(loc => { html += `<th>${loc.name}</th>`; });
        html += '</tr></thead><tbody>';

        state.dates.filter(d => !d.isHoliday).forEach(dateInfo => {
            html += `<tr><td>${dateInfo.date}</td><td>${dateInfo.dayOfWeek}</td>`;
            state.locations.forEach(loc => {
                const key = `${dateInfo.date}_${loc.name}`;
                const data = densityMap[key];
                if (data?.count > 0) {
                    const bg = data.count >= 3 ? '#ffcdd2' : data.count >= 2 ? '#ffe0b2' : '#e8f5e9';
                    html += `<td style="background:${bg};text-align:center;cursor:help;" title="${data.names.join(', ')}">${data.count}人</td>`;
                } else {
                    html += '<td style="text-align:center;color:#ccc;">-</td>';
                }
            });
            html += '</tr>';
        });
        html += '</tbody>';
        document.getElementById('preferredDatesDensityTable').innerHTML = html;
    }

    function displayDetailAnalysis(result) {
        let html = '<thead><tr><th>No.</th><th>名前</th><th>希望場所</th><th>参加可能回数</th><th>割り当て回数</th><th>判定</th></tr></thead><tbody>';
        result.participants.forEach((p, idx) => {
            const status = p.currentAssignments === 0 ? '⚠️0回' :
                p.currentAssignments >= p.maxAssignments ? '✅OK' : '📌一部';
            html += `<tr><td>${idx + 1}</td><td>${p.displayName}</td><td>${p.preferredLocations.join(', ')}</td><td>${p.maxAssignments}</td><td>${p.currentAssignments}</td><td>${status}</td></tr>`;
        });
        html += '</tbody>';
        document.getElementById('detailAnalysisTable').innerHTML = html;
    }

    // ========================================================================
    // Excel Export
    // ========================================================================
    function exportToExcel() {
        if (!state.assignmentResult) return;

        const result = state.assignmentResult;
        const wb = XLSX.utils.book_new();

        // Sheet 1: Shift table
        const shiftData = [['日付', '曜日', ...state.locations.map(l => l.name)]];
        state.dates.filter(d => !d.isHoliday).forEach(dateInfo => {
            const row = [dateInfo.date, dateInfo.dayOfWeek];
            state.locations.forEach(loc => {
                const slot = result.slots.find(s => s.date === dateInfo.date && s.location === loc.name);
                row.push(slot?.assignedTo || '');
            });
            shiftData.push(row);
        });
        const ws1 = XLSX.utils.aoa_to_sheet(shiftData);
        XLSX.utils.book_append_sheet(wb, ws1, "シフト表");

        // Sheet 2: Summary
        const summaryData = [['No.', '名前', '参加可能回数', '割り当て回数', '割り当て日']];
        result.participants.forEach((p, idx) => {
            const dates = result.slots.filter(s => s.assignedTo === p.displayName).map(s => s.date);
            summaryData.push([idx + 1, p.displayName, p.maxAssignments, p.currentAssignments, dates.join(', ')]);
        });
        const ws2 = XLSX.utils.aoa_to_sheet(summaryData);
        XLSX.utils.book_append_sheet(wb, ws2, "割り当て集計");

        const now = new Date();
        const timestamp = `${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}_${now.getHours().toString().padStart(2, '0')}${now.getMinutes().toString().padStart(2, '0')}`;
        XLSX.writeFile(wb, `shift_result_${timestamp}.xlsx`);
    }
});
