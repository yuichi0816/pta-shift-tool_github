/**
 * PTA旗振りシフト自動割り当てツール v3.0
 * アンケート結果を基にシフトを自動で割り当てる
 * 
 * v3.0新機能: 前後NG期間制約
 * - 同一人物が短期間（デフォルト7日）に連続して割り当てられることを防ぐ
 */

// ========================================================================
// v3.0 変数
// ========================================================================
let NG_PERIOD_DAYS = 7; // 前後NG期間（デフォルト: 7日間）
// 同一参加者が割り当てられた日から前後この日数の間は、別のシフトに割り当てない

/**
 * ユーザーが入力したNG期間日数を取得する
 * @returns {number} NG期間日数（デフォルト: 7）
 */
function getUserNgPeriodDays() {
    const input = document.getElementById('ngPeriodDays');
    if (input) {
        const value = parseInt(input.value, 10);
        return !isNaN(value) && value >= 0 ? value : 7;
    }
    return 7;
}

document.addEventListener('DOMContentLoaded', () => {
    // ========================================================================
    // State
    // ========================================================================
    const state = {
        shiftData: null,      // シフト表データ
        surveyData: null,     // アンケートデータ
        locations: [],        // 場所リスト
        dates: [],            // 日付リスト
        assignmentResult: null, // 割り当て結果
        surveyMapping: {
            headers: [],
            columnTags: [],
            isValid: false,
            resolvedIndices: null,
            issues: [],
            filter: {
                issuesOnly: false
            }
        },
        surveyFileName: '',
        lastMappingSnapshot: null,
        assignmentBlockReason: '',
        uploadDragState: {
            survey: false,
            shift: false
        },
        tabsInitialized: false,
        isAssigning: false
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
    const previewSection = document.getElementById('previewSection');
    const assignSection = document.getElementById('assignSection');
    const resultSection = document.getElementById('resultSection');
    const unassignedSection = document.getElementById('unassignedSection');
    const assignBtn = document.getElementById('assignBtn');
    const assignBlockedReason = document.getElementById('assignBlockedReason');
    const exportBtn = document.getElementById('exportBtn');
    const progressContainer = document.getElementById('progressContainer');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    const mappingSection = document.getElementById('mappingSection');
    const mappingMeta = document.getElementById('mappingMeta');
    const mappingSummary = document.getElementById('mappingSummary');
    const mappingIssueNav = document.getElementById('mappingIssueNav');
    const mappingTableBody = document.getElementById('mappingTableBody');
    const mappingConfirmStatus = document.getElementById('mappingConfirmStatus');
    const showIssueOnly = document.getElementById('showIssueOnly');
    const unmappedCountBadge = document.getElementById('unmappedCountBadge');
    const mappingJsonFile = document.getElementById('mappingJsonFile');
    const mappingRestoreNotice = document.getElementById('mappingRestoreNotice');
    const mappingRestoreText = document.getElementById('mappingRestoreText');
    const restoreLastMappingBtn = document.getElementById('restoreLastMappingBtn');

    const SURVEY_TAGS = [
        { id: 'UNMAPPED', label: '-- 未使用 --', required: false, type: 'ignore' },
        { id: 'TIMESTAMP', label: 'TIMESTAMP: タイムスタンプ', required: false, type: 'single' },
        { id: 'EMAIL', label: 'EMAIL: メールアドレス', required: false, type: 'single' },
        { id: 'PARTICIPATION', label: 'PARTICIPATION: 参加可否/参加確認', required: false, type: 'single' },
        { id: 'COUNT', label: 'COUNT: 参加可能回数', required: true, type: 'single' },
        { id: 'NG_DATE', label: 'NG_DATE: 参加できない日', required: false, type: 'single' },
        { id: 'NG_DAY', label: 'NG_DAY: 参加できない曜日', required: false, type: 'single' },
        { id: 'NG_MONTH', label: 'NG_MONTH: 参加できない月', required: false, type: 'single' },
        { id: 'PREF_DAY', label: 'PREF_DAY: 参加可能曜日', required: false, type: 'single' },
        { id: 'PREF_MONTH', label: 'PREF_MONTH: 参加可能月', required: false, type: 'single' },
        { id: 'PREF_DATE_ONLY', label: 'PREF_DATE_ONLY: この日なら参加できる日', required: false, type: 'single' },
        { id: 'LOC_1', label: 'LOC_1: 第1希望場所', required: true, type: 'single' },
        { id: 'LOC_2', label: 'LOC_2: 第2希望場所', required: false, type: 'single' },
        { id: 'LOC_3', label: 'LOC_3: 第3希望場所', required: false, type: 'single' },
        { id: 'LOC_4', label: 'LOC_4: 第4希望場所', required: false, type: 'single' },
        { id: 'LOC_5', label: 'LOC_5: 第5希望場所', required: false, type: 'single' },
        { id: 'ADD_SUPPORT', label: 'ADD_SUPPORT: 欠員時の追加対応可否', required: false, type: 'single' },
        { id: 'ADD_COUNT', label: 'ADD_COUNT: 追加対応可能回数', required: false, type: 'single' },
        { id: 'GRADE', label: 'GRADE: 学年', required: false, type: 'single' },
        { id: 'NAME', label: 'NAME: 氏名（単一列）', required: false, type: 'single' },
        { id: 'LAST_NAME', label: 'LAST_NAME: 姓', required: false, type: 'single' },
        { id: 'FIRST_NAME', label: 'FIRST_NAME: 名', required: false, type: 'single' },
        { id: 'CLASS', label: 'CLASS: クラス', required: false, type: 'single' },
        { id: 'FREE_TEXT', label: 'FREE_TEXT: 連絡事項', required: false, type: 'single' }
    ];
    const SURVEY_TAG_MAP = new Map(SURVEY_TAGS.map(tag => [tag.id, tag]));
    const SURVEY_REQUIRED_TAGS = ['COUNT', 'LOC_1'];
    const SUPPORTED_UPLOAD_EXTENSIONS = ['.xlsx', '.xls', '.csv'];
    const LOCAL_STORAGE_MAPPING_KEY = 'pta_shift:last_mapping';
    const LOCAL_STORAGE_TEST_KEY = 'pta_shift:storage_check';
    const MAPPING_STORAGE_VERSION = 1;
    const localStorageAvailable = checkLocalStorageAvailable();

    // ========================================================================
    // File Upload Handlers
    // ========================================================================
    shiftFile.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        await handleShiftFileUpload(file);
        shiftFile.value = '';
    });

    surveyFile.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) return;
        await handleSurveyFileUpload(file);
        surveyFile.value = '';
    });

    setupUploadDragAndDrop();

    if (showIssueOnly) {
        showIssueOnly.addEventListener('change', (event) => {
            state.surveyMapping.filter.issuesOnly = !!event.target.checked;
            renderSurveyMappingTable();
        });
    }
    if (restoreLastMappingBtn) {
        restoreLastMappingBtn.addEventListener('click', restoreLastMappingFromStorage);
    }

    async function handleShiftFileUpload(file) {
        if (!isSupportedUploadFile(file)) {
            shiftStatus.textContent = `❌ 非対応形式です（対応: ${SUPPORTED_UPLOAD_EXTENSIONS.join(', ')}）`;
            updateAssignmentButtonState();
            return;
        }

        try {
            state.shiftData = await parseExcelFile(file);
            console.log('=== シフト表データ ===');
            console.log('行数:', state.shiftData.length);
            console.log('最初の5行:', state.shiftData.slice(0, 5));
            processShiftData();
            shiftUploadBox.classList.add('uploaded');
            shiftStatus.textContent = `✓ ${file.name} を読み込みました (${state.dates.length}日, ${state.locations.length}場所)`;
            checkBothFilesLoaded();
        } catch (error) {
            console.error('シフト表の読み込みエラー:', error);
            shiftStatus.textContent = `❌ 読み込みエラー: ${error.message}`;
            updateAssignmentButtonState();
        }
    }

    async function handleSurveyFileUpload(file) {
        if (!isSupportedUploadFile(file)) {
            surveyStatus.textContent = `❌ 非対応形式です（対応: ${SUPPORTED_UPLOAD_EXTENSIONS.join(', ')}）`;
            updateAssignmentButtonState();
            return;
        }

        try {
            state.surveyFileName = file.name;
            state.surveyData = await parseExcelFile(file);
            console.log('=== アンケートデータ ===');
            console.log('行数:', state.surveyData.length);
            console.log('ヘッダー:', state.surveyData[0]);
            console.log('最初のデータ行:', state.surveyData[1]);
            processSurveyData();
            surveyUploadBox.classList.add('uploaded');
            surveyStatus.textContent = `✓ ${file.name} を読み込みました (${state.surveyData.length - 1}件)`;
            checkBothFilesLoaded();
        } catch (error) {
            console.error('アンケートの読み込みエラー:', error);
            surveyStatus.textContent = `❌ 読み込みエラー: ${error.message}`;
            updateAssignmentButtonState();
        }
    }

    function setupUploadDragAndDrop() {
        bindDropZone(shiftUploadBox, handleShiftFileUpload);
        bindDropZone(surveyUploadBox, handleSurveyFileUpload);

        document.addEventListener('dragover', (event) => {
            if (isFileDragEvent(event)) {
                event.preventDefault();
            }
        });
        document.addEventListener('drop', (event) => {
            if (isFileDragEvent(event)) {
                event.preventDefault();
            }
        });
    }

    function bindDropZone(uploadBox, handler) {
        if (!uploadBox) return;
        let dragDepth = 0;
        const dragStateKey = uploadBox === surveyUploadBox ? 'survey' : 'shift';

        uploadBox.addEventListener('dragenter', (event) => {
            if (!isFileDragEvent(event)) return;
            event.preventDefault();
            event.stopPropagation();
            dragDepth += 1;
            uploadBox.classList.add('drag-active');
            state.uploadDragState[dragStateKey] = true;
        });

        uploadBox.addEventListener('dragover', (event) => {
            if (!isFileDragEvent(event)) return;
            event.preventDefault();
            event.stopPropagation();
            uploadBox.classList.add('drag-active');
            state.uploadDragState[dragStateKey] = true;
        });

        uploadBox.addEventListener('dragleave', (event) => {
            event.preventDefault();
            event.stopPropagation();
            dragDepth = Math.max(dragDepth - 1, 0);
            if (dragDepth === 0) {
                uploadBox.classList.remove('drag-active');
                state.uploadDragState[dragStateKey] = false;
            }
        });

        uploadBox.addEventListener('dragend', () => {
            dragDepth = 0;
            uploadBox.classList.remove('drag-active');
            state.uploadDragState[dragStateKey] = false;
        });

        uploadBox.addEventListener('drop', async (event) => {
            if (!isFileDragEvent(event)) return;
            event.preventDefault();
            event.stopPropagation();
            dragDepth = 0;
            uploadBox.classList.remove('drag-active');
            state.uploadDragState[dragStateKey] = false;

            const file = event.dataTransfer.files && event.dataTransfer.files[0];
            if (!file) return;
            await handler(file);
        });
    }

    function isFileDragEvent(event) {
        if (!event || !event.dataTransfer) return false;
        if (!event.dataTransfer.types) return true;
        return Array.from(event.dataTransfer.types).includes('Files');
    }

    function isSupportedUploadFile(file) {
        if (!file || !file.name) return false;
        const name = file.name.toLowerCase();
        return SUPPORTED_UPLOAD_EXTENSIONS.some(ext => name.endsWith(ext));
    }

    function checkLocalStorageAvailable() {
        try {
            if (typeof window === 'undefined' || !window.localStorage) return false;
            window.localStorage.setItem(LOCAL_STORAGE_TEST_KEY, 'ok');
            window.localStorage.removeItem(LOCAL_STORAGE_TEST_KEY);
            return true;
        } catch (error) {
            return false;
        }
    }

    // ========================================================================
    // Excel File Parser
    // ========================================================================
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
        if (!state.shiftData || state.shiftData.length < 2) {
            console.error('シフトデータが不足しています');
            return;
        }

        // ヘッダー行を探す（「日」を含む行）
        let headerRowIndex = -1;
        for (let i = 0; i < Math.min(5, state.shiftData.length); i++) {
            const row = state.shiftData[i];
            if (row && row[0] && row[0].toString().includes('日')) {
                headerRowIndex = i;
                break;
            }
        }

        // ヘッダー行が見つからない場合は1行目を使用
        if (headerRowIndex === -1) {
            headerRowIndex = 1;
        }

        console.log('ヘッダー行インデックス:', headerRowIndex);
        const headerRow = state.shiftData[headerRowIndex];
        console.log('ヘッダー行:', headerRow);

        // 場所を抽出（日付と曜日の後のカラム）
        // 定員情報の抽出: 「地点A（2人体制）」形式から定員を取得
        state.locations = [];
        for (let i = 2; i < headerRow.length; i++) {
            if (headerRow[i] && headerRow[i].toString().trim()) {
                const fullName = headerRow[i].toString().trim();
                // 「更新」などのメタデータカラムを除外
                if (!fullName.includes('更新') && !fullName.includes('変更') && !fullName.includes('削除')) {
                    // 定員を抽出: 「（N人体制）」「(N人体制)」「（N名）」「(N名)」などのパターン
                    // 全角・半角両方に対応
                    const capacityMatch = fullName.match(/[（(](\d+|[０-９]+)\s*人?\s*(体制|名)?[）)]/);
                    let capacity = 1; // デフォルト1名
                    let baseName = fullName;

                    if (capacityMatch) {
                        // 数字を抽出（全角数字を半角に変換）
                        let numStr = capacityMatch[1].replace(/[０-９]/g, s =>
                            String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
                        );
                        capacity = parseInt(numStr, 10) || 1;
                        // 定員部分を除いた基本名を取得（アンケートとのマッチング用）
                        // 「（N人体制）」「(N人体制)」「（N名）」などを全て除去
                        baseName = fullName.replace(/[（(](\d+|[０-９]+)\s*人?\s*(体制|名)?[）)]/g, '').trim();
                    }

                    state.locations.push({
                        index: i,
                        name: fullName,      // 元の名前（定員情報含む）
                        baseName: baseName,  // アンケートとのマッチング用の基本名
                        capacity: capacity   // 定員（1以上）
                    });

                    // デバッグログ
                    if (capacity > 1) {
                        console.log(`📍 場所「${fullName}」→ 基本名「${baseName}」、定員: ${capacity}名`);
                    }
                }
            }
        }

        console.log('場所リスト:', state.locations);

        // 日付を抽出（ヘッダー行の次の行以降）
        state.dates = [];
        for (let i = headerRowIndex + 1; i < state.shiftData.length; i++) {
            const row = state.shiftData[i];
            if (row && row[0]) {
                const dateValue = row[0];
                let dateStr = '';

                // Excelの日付シリアル値の場合
                if (typeof dateValue === 'number') {
                    const date = new Date((dateValue - 25569) * 86400 * 1000);
                    dateStr = `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`;
                } else {
                    dateStr = dateValue.toString();
                }

                // 日付の形式チェック
                if (dateStr.match(/\d+\/\d+/) || dateStr.match(/\d{4}/)) {
                    const dayOfWeek = row[1] ? row[1].toString() : '';

                    // 土日祝は除外
                    const isHoliday = ['土', '日', '祝'].some(d => dayOfWeek.includes(d));

                    state.dates.push({
                        rowIndex: i,
                        date: dateStr,
                        dayOfWeek: dayOfWeek,
                        isHoliday: isHoliday
                    });
                }
            }
        }

        console.log('日付リスト:', state.dates.slice(0, 10));
        console.log('登校日数:', state.dates.filter(d => !d.isHoliday).length);
    }

    function processSurveyData() {
        if (!state.surveyData || state.surveyData.length < 2) {
            console.error('アンケートデータが不足しています');
            return;
        }

        // ヘッダー行を解析
        const headers = state.surveyData[0];
        console.log('アンケートヘッダー数:', headers.length);

        // 各カラムの内容を確認
        headers.forEach((h, i) => {
            if (h) console.log(`  [${i}]: ${h.toString().substring(0, 50)}`);
        });

        state.surveyMapping.headers = headers.map((header, idx) => {
            const text = header === null || header === undefined ? '' : header.toString().trim();
            return text || `無題列_${idx + 1}`;
        });
        state.surveyMapping.columnTags = new Array(state.surveyMapping.headers.length).fill('UNMAPPED');
        state.surveyMapping.isValid = false;
        state.surveyMapping.resolvedIndices = null;
        state.surveyMapping.issues = [];
        state.surveyMapping.filter.issuesOnly = false;
        if (showIssueOnly) {
            showIssueOnly.checked = false;
        }

        if (mappingSection) {
            mappingSection.classList.remove('hidden');
            mappingSection.style.display = 'block';
        }

        renderSurveyMappingMeta();
        applyAutoTagMapping();
        refreshMappingRestoreCandidate();
        renderMappingStatus();
    }

    function renderSurveyMappingMeta() {
        if (!mappingMeta) return;

        const rowCount = Math.max(state.surveyData.length - 1, 0);
        mappingMeta.innerHTML = `
            <div class="stat-item">📄 ファイル: <strong>${escapeHtml(state.surveyFileName || '-')}</strong></div>
            <div class="stat-item">📊 データ行: <strong>${rowCount}件</strong></div>
            <div class="stat-item">🧱 列数: <strong>${state.surveyMapping.headers.length}列</strong></div>
        `;
    }

    function applyAutoTagMapping(options = {}) {
        if (!state.surveyMapping.headers || state.surveyMapping.headers.length === 0) return;
        const nextTags = state.surveyMapping.headers.map(header => suggestSurveyTagFromHeader(header));
        applyMappingTags(nextTags);
        if (options.persist && state.surveyMapping.isValid) {
            saveCurrentMappingToStorage();
        }
    }

    function applyMappingTags(nextTags) {
        if (!Array.isArray(nextTags)) return;
        state.surveyMapping.columnTags = [...nextTags];
        renderSurveyMappingTable();
    }

    function collectDuplicateTagColumnIndices(columnTags) {
        const tagToColumns = new Map();
        columnTags.forEach((tagId, colIndex) => {
            if (!tagId || tagId === 'UNMAPPED') return;
            const tagDef = SURVEY_TAG_MAP.get(tagId);
            if (!tagDef || tagDef.type !== 'single') return;
            if (!tagToColumns.has(tagId)) tagToColumns.set(tagId, []);
            tagToColumns.get(tagId).push(colIndex);
        });

        const duplicateColumnIndices = new Set();
        tagToColumns.forEach((columns) => {
            if (columns.length <= 1) return;
            columns.forEach(colIndex => duplicateColumnIndices.add(colIndex));
        });
        return duplicateColumnIndices;
    }

    function renderSurveyMappingTable() {
        if (!mappingTableBody) return;

        if (!state.surveyMapping.headers || state.surveyMapping.headers.length === 0) {
            mappingTableBody.innerHTML = '<tr><td colspan="4" style="color:#64748b;">アンケートファイルを読み込むと列一覧が表示されます。</td></tr>';
            renderMappingIssueNav();
            renderMappingStatus();
            return;
        }

        const totalColumns = state.surveyMapping.headers.length;
        const unmappedCount = state.surveyMapping.columnTags.filter(tagId => tagId === 'UNMAPPED').length;
        const duplicateColumnIndices = collectDuplicateTagColumnIndices(state.surveyMapping.columnTags);
        const duplicateCount = duplicateColumnIndices.size;
        if (unmappedCountBadge) {
            unmappedCountBadge.textContent = `未対応 ${unmappedCount}件 / 重複 ${duplicateCount}列`;
        }

        const indicesToRender = [];
        for (let colIndex = 0; colIndex < totalColumns; colIndex++) {
            const isUnmapped = state.surveyMapping.columnTags[colIndex] === 'UNMAPPED';
            const isIssueColumn = isUnmapped || duplicateColumnIndices.has(colIndex);
            if (state.surveyMapping.filter.issuesOnly && !isIssueColumn) {
                continue;
            }
            indicesToRender.push(colIndex);
        }

        if (indicesToRender.length === 0) {
            mappingTableBody.innerHTML = '<tr><td colspan="4" style="color:#64748b;">未対応・重複列はありません（表示対象なし）</td></tr>';
            validateSurveyTagMapping();
            return;
        }

        const rows = [];
        for (const colIndex of indicesToRender) {
            rows.push(`
                <tr data-col-index="${colIndex}">
                    <td>${colIndex + 1}</td>
                    <td class="mapping-header-cell">${escapeHtml(state.surveyMapping.headers[colIndex])}</td>
                    <td style="white-space: pre-wrap; line-height: 1.4;">${escapeHtml(getSurveyColumnSample(colIndex, 3))}</td>
                    <td>
                        <select class="mapping-tag-select" data-col-index="${colIndex}" style="width: 100%; max-width: 280px; padding: 6px; border-radius: 6px; border: 1px solid #cbd5e1;">
                            ${SURVEY_TAGS.map(tag => {
                                const selected = state.surveyMapping.columnTags[colIndex] === tag.id ? 'selected' : '';
                                return `<option value="${tag.id}" ${selected}>${escapeHtml(tag.label)}</option>`;
                            }).join('')}
                        </select>
                    </td>
                </tr>
            `);
        }
        mappingTableBody.innerHTML = rows.join('');

        mappingTableBody.querySelectorAll('select[data-col-index]').forEach(select => {
            select.addEventListener('change', (event) => {
                const colIndex = Number(event.target.dataset.colIndex);
                const nextTag = event.target.value;
                state.surveyMapping.columnTags[colIndex] = nextTag;
                validateSurveyTagMapping();
                if (state.surveyMapping.isValid) {
                    saveCurrentMappingToStorage();
                }
            });
        });

        validateSurveyTagMapping();
    }

    function renderMappingStatus() {
        if (!mappingConfirmStatus) return;

        if (!state.surveyMapping.headers || state.surveyMapping.headers.length === 0) {
            mappingConfirmStatus.textContent = 'アンケートファイルを読み込んでください';
            mappingConfirmStatus.style.color = '#475569';
            mappingConfirmStatus.style.borderColor = '#e2e8f0';
            mappingConfirmStatus.style.background = '#f8fafc';
            return;
        }

        if (!state.surveyMapping.isValid) {
            mappingConfirmStatus.textContent = 'マッピング未完了（必須タグ不足または重複を修正してください）';
            mappingConfirmStatus.style.color = '#b91c1c';
            mappingConfirmStatus.style.borderColor = '#fecaca';
            mappingConfirmStatus.style.background = '#fef2f2';
            return;
        }

        mappingConfirmStatus.textContent = 'マッピング完了。STEP 2 へ進めます。';
        mappingConfirmStatus.style.color = '#166534';
        mappingConfirmStatus.style.borderColor = '#bbf7d0';
        mappingConfirmStatus.style.background = '#f0fdf4';
    }

    function getSurveyColumnSample(colIndex, maxCount = 3) {
        const samples = [];
        const upper = Math.min(state.surveyData.length, 80);
        for (let rowIndex = 1; rowIndex < upper; rowIndex++) {
            const row = state.surveyData[rowIndex] || [];
            const value = row[colIndex] === null || row[colIndex] === undefined ? '' : row[colIndex].toString().trim();
            if (!value) continue;
            samples.push(value);
            if (samples.length >= maxCount) break;
        }
        return samples.length > 0 ? samples.join('\n---\n') : '(空欄)';
    }

    function suggestSurveyTagFromHeader(header) {
        const h = normalizeTagHeaderText(header);

        if (h.includes('第1希望') || h.includes('第１希望')) return 'LOC_1';
        if (h.includes('第2希望') || h.includes('第２希望')) return 'LOC_2';
        if (h.includes('第3希望') || h.includes('第３希望')) return 'LOC_3';
        if (h.includes('第4希望') || h.includes('第４希望')) return 'LOC_4';
        if (h.includes('第5希望') || h.includes('第５希望')) return 'LOC_5';

        if (h.includes('タイムスタンプ')) return 'TIMESTAMP';
        if (h.includes('メール') || h.includes('email') || h.includes('mail') || h.includes('ユーザー名')) return 'EMAIL';
        if (h.includes('参加確認') || h.includes('旗振り当番について')) return 'PARTICIPATION';

        const hasAdditionalKeyword = h.includes('追加') || h.includes('欠員');
        if ((h.includes('参加可能回数') || h.includes('希望回数')) && !hasAdditionalKeyword) return 'COUNT';
        if (hasAdditionalKeyword) {
            const looksLikeSupportQuestion =
                h.includes('対応') || h.includes('可否') || h.includes('できますか') || h.includes('いただけますか');
            const looksLikeCountField =
                h.includes('追加回数') || h.includes('対応回数') || h.includes('何回') || h.includes('回まで');

            if (looksLikeCountField || (h.includes('回数') && !looksLikeSupportQuestion && !h.includes('超えて'))) {
                return 'ADD_COUNT';
            }
            return 'ADD_SUPPORT';
        }

        if ((h.includes('ng') || h.includes('参加できない') || h.includes('不参加')) && h.includes('曜日')) return 'NG_DAY';
        if ((h.includes('ng') || h.includes('参加できない') || h.includes('不参加')) && h.includes('月')) return 'NG_MONTH';
        if ((h.includes('ng') || h.includes('参加できない') || h.includes('不参加')) && h.includes('日')) return 'NG_DATE';
        if ((h.includes('この日なら') || h.includes('特定') || h.includes('参加可能日')) && !h.includes('参加できない')) return 'PREF_DATE_ONLY';

        if (h.includes('曜日') && (h.includes('参加可能') || h.includes('希望') || h.includes('選択'))) return 'PREF_DAY';
        if (h.includes('月') && (h.includes('参加可能') || h.includes('希望') || h.includes('選択'))) return 'PREF_MONTH';

        if ((h.includes('氏名') || h.includes('名前')) && (h.includes('姓') || h.includes('名字'))) return 'LAST_NAME';
        if ((h.includes('氏名') || h.includes('名前')) && (h.includes('（名）') || h.includes('名（'))) return 'FIRST_NAME';
        if (h.includes('姓') || h.includes('名字')) return 'LAST_NAME';
        if (h.includes('名') && !h.includes('校') && !h.includes('宛') && !h.includes('名字') && !h.includes('氏名')) return 'FIRST_NAME';
        if ((h.includes('氏名') || h.includes('名前')) && !h.includes('姓')) return 'NAME';
        if (h.includes('学年')) return 'GRADE';
        if (h.includes('クラス') || h.includes('組')) return 'CLASS';
        if (h.includes('連絡') || h.includes('意見') || h.includes('その他') || h.includes('自由記述')) return 'FREE_TEXT';

        return 'UNMAPPED';
    }

    function normalizeTagHeaderText(text) {
        return (text || '')
            .toString()
            .toLowerCase()
            .replace(/[Ａ-Ｚａ-ｚ０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
            .replace(/\s+/g, ' ')
            .trim();
    }

    function buildMappingHeaderSignature(headers) {
        const normalized = (headers || []).map(header => normalizeTagHeaderText(header)).join('|');
        let hash = 2166136261;
        for (let i = 0; i < normalized.length; i++) {
            hash ^= normalized.charCodeAt(i);
            hash += (hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24);
        }
        const hashHex = (hash >>> 0).toString(16).padStart(8, '0');
        return `${headers.length}:${hashHex}`;
    }

    function buildCurrentMappingSnapshot() {
        if (!state.surveyMapping.headers || state.surveyMapping.headers.length === 0) return null;

        return {
            version: MAPPING_STORAGE_VERSION,
            sourceFileName: state.surveyFileName || '',
            generatedAt: new Date().toISOString(),
            columnCount: state.surveyMapping.headers.length,
            headerSignature: buildMappingHeaderSignature(state.surveyMapping.headers),
            columns: state.surveyMapping.headers.map((header, colIndex) => ({
                colIndex: colIndex,
                colNumber: colIndex + 1,
                sourceHeader: header,
                tagId: state.surveyMapping.columnTags[colIndex] || 'UNMAPPED'
            }))
        };
    }

    function saveMappingSnapshotToStorage(snapshot) {
        if (!localStorageAvailable || !snapshot) return false;

        try {
            window.localStorage.setItem(LOCAL_STORAGE_MAPPING_KEY, JSON.stringify(snapshot));
            state.lastMappingSnapshot = snapshot;
            return true;
        } catch (error) {
            return false;
        }
    }

    function loadMappingSnapshotFromStorage() {
        if (!localStorageAvailable) return null;

        try {
            const raw = window.localStorage.getItem(LOCAL_STORAGE_MAPPING_KEY);
            if (!raw) return null;
            const parsed = JSON.parse(raw);
            if (!parsed || !Array.isArray(parsed.columns)) return null;
            if (typeof parsed.columnCount !== 'number') return null;
            if (typeof parsed.headerSignature !== 'string') return null;
            return parsed;
        } catch (error) {
            return null;
        }
    }

    function saveCurrentMappingToStorage() {
        const snapshot = buildCurrentMappingSnapshot();
        return saveMappingSnapshotToStorage(snapshot);
    }

    function isCompatibleMappingSnapshot(snapshot) {
        if (!snapshot || !state.surveyMapping.headers || state.surveyMapping.headers.length === 0) return false;
        const currentColumnCount = state.surveyMapping.headers.length;
        if (snapshot.columnCount !== currentColumnCount) return false;
        const currentSignature = buildMappingHeaderSignature(state.surveyMapping.headers);
        return snapshot.headerSignature === currentSignature;
    }

    function renderMappingRestoreNotice(options = {}) {
        if (!mappingRestoreNotice || !mappingRestoreText || !restoreLastMappingBtn) return;

        if (!options.message) {
            mappingRestoreNotice.classList.add('hidden');
            return;
        }

        mappingRestoreNotice.classList.remove('hidden');
        mappingRestoreText.textContent = options.message;
        restoreLastMappingBtn.classList.toggle('hidden', !options.showRestoreButton);
    }

    function refreshMappingRestoreCandidate() {
        if (!state.surveyMapping.headers || state.surveyMapping.headers.length === 0) {
            renderMappingRestoreNotice();
            return;
        }

        if (!localStorageAvailable) {
            renderMappingRestoreNotice();
            return;
        }

        const snapshot = loadMappingSnapshotFromStorage();
        if (!snapshot) {
            renderMappingRestoreNotice();
            return;
        }

        state.lastMappingSnapshot = snapshot;

        if (isCompatibleMappingSnapshot(snapshot)) {
            renderMappingRestoreNotice({
                message: '同じ形式の前回マッピング設定が見つかりました。必要なら復元できます。',
                showRestoreButton: true
            });
            return;
        }

        renderMappingRestoreNotice({
            message: '前回設定はこのファイル形式に適合しません。'
        });
    }

    function restoreLastMappingFromStorage() {
        const snapshot = state.lastMappingSnapshot || loadMappingSnapshotFromStorage();
        if (!snapshot) {
            alert('前回設定が見つかりませんでした。');
            return;
        }

        if (!isCompatibleMappingSnapshot(snapshot)) {
            renderMappingRestoreNotice({
                message: '前回設定はこのファイル形式に適合しません。'
            });
            alert('前回設定はこのファイル形式に適合しません。');
            return;
        }

        const shouldRestore = window.confirm('前回設定を復元しますか？');
        if (!shouldRestore) return;

        const applied = applyMappingColumns(snapshot.columns);
        if (!applied) {
            alert('前回設定の復元に失敗しました。');
            return;
        }
        saveCurrentMappingToStorage();

        renderMappingRestoreNotice({
            message: '前回設定を復元しました。',
            showRestoreButton: false
        });
    }

    function createEmptyColumnIndices() {
        return {
            email: null,
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
            pref4: null,
            pref5: null,
            additionalSupport: null,
            additionalCount: null,
            grade: null,
            fullName: null,
            lastName: null,
            firstName: null,
            classNum: null,
            freeText: null
        };
    }

    function validateSurveyTagMapping() {
        const tagToColumns = new Map();
        state.surveyMapping.columnTags.forEach((tagId, colIndex) => {
            if (!tagId || tagId === 'UNMAPPED') return;
            if (!tagToColumns.has(tagId)) tagToColumns.set(tagId, []);
            tagToColumns.get(tagId).push(colIndex);
        });

        const missingRequired = SURVEY_REQUIRED_TAGS.filter(tagId => !tagToColumns.has(tagId));
        const hasFullName = tagToColumns.has('NAME');
        const hasLegacyName = tagToColumns.has('LAST_NAME') && tagToColumns.has('FIRST_NAME');
        const duplicateTags = [];
        tagToColumns.forEach((columns, tagId) => {
            const tagDef = SURVEY_TAG_MAP.get(tagId);
            if (!tagDef || tagDef.type !== 'single') return;
            if (columns.length > 1) duplicateTags.push({ tagId, columns });
        });

        const mappedCount = state.surveyMapping.columnTags.filter(tagId => tagId !== 'UNMAPPED').length;
        const problems = [];
        if (missingRequired.length > 0) {
            problems.push(`必須タグ不足: ${missingRequired.join(', ')}`);
        }
        if (!hasFullName && !hasLegacyName) {
            problems.push('氏名タグ不足: NAME または LAST_NAME+FIRST_NAME を設定してください');
        }
        if (duplicateTags.length > 0) {
            const msg = duplicateTags.map(item => `${item.tagId}(列:${item.columns.map(c => c + 1).join(',')})`).join(' / ');
            problems.push(`重複タグ: ${msg}`);
        }

        const issues = [];
        missingRequired.forEach(tagId => {
            issues.push({
                type: 'missing',
                tagId,
                colIndices: []
            });
        });
        duplicateTags.forEach(item => {
            issues.push({
                type: 'duplicate',
                tagId: item.tagId,
                colIndices: [...item.columns]
            });
        });

        const isValid = problems.length === 0;
        state.surveyMapping.isValid = isValid;
        state.surveyMapping.resolvedIndices = isValid ? buildColumnIndicesFromTagMapping(tagToColumns) : null;
        state.surveyMapping.issues = issues;

        if (mappingSummary) {
            let summaryText = `マッピング済み ${mappedCount}/${state.surveyMapping.headers.length} 列。`;
            mappingSummary.style.color = '#166534';
            mappingSummary.style.borderColor = '#bbf7d0';
            mappingSummary.style.background = '#f0fdf4';

            if (!isValid) {
                summaryText += ` ${problems.join(' | ')}`;
                mappingSummary.style.color = '#b91c1c';
                mappingSummary.style.borderColor = '#fecaca';
                mappingSummary.style.background = '#fef2f2';
            } else if (mappedCount < state.surveyMapping.headers.length) {
                summaryText += ' 未使用列があります。不要ならそのままで問題ありません。';
                mappingSummary.style.color = '#b45309';
                mappingSummary.style.borderColor = '#fde68a';
                mappingSummary.style.background = '#fffbeb';
            } else {
                summaryText += ' 必須タグはすべて揃っています。';
            }
            mappingSummary.textContent = summaryText;
        }

        renderMappingIssueNav();
        renderMappingStatus();
        checkBothFilesLoaded();
    }

    function renderMappingIssueNav() {
        if (!mappingIssueNav) return;

        const issues = state.surveyMapping.issues || [];
        if (issues.length === 0) {
            mappingIssueNav.classList.add('hidden');
            mappingIssueNav.innerHTML = '';
            return;
        }

        const buttons = [];
        issues.forEach(issue => {
            if (issue.type === 'missing') {
                buttons.push(`<button type="button" data-issue-type="missing" data-tag-id="${issue.tagId}">${issue.tagId} 未設定</button>`);
                return;
            }

            if (issue.type === 'duplicate' && Array.isArray(issue.colIndices)) {
                issue.colIndices.forEach(colIndex => {
                    buttons.push(`<button type="button" data-issue-type="duplicate" data-tag-id="${issue.tagId}" data-col-index="${colIndex}">${issue.tagId} 重複（${colIndex + 1}列）</button>`);
                });
            }
        });

        if (buttons.length === 0) {
            mappingIssueNav.classList.add('hidden');
            mappingIssueNav.innerHTML = '';
            return;
        }

        mappingIssueNav.classList.remove('hidden');
        mappingIssueNav.innerHTML = buttons.join('');
        mappingIssueNav.querySelectorAll('button').forEach(button => {
            button.addEventListener('click', () => {
                const issueType = button.dataset.issueType;
                if (issueType === 'duplicate') {
                    const colIndex = Number(button.dataset.colIndex);
                    focusMappingRow(colIndex);
                    return;
                }

                if (issueType === 'missing') {
                    const firstUnmapped = state.surveyMapping.columnTags.findIndex(tagId => tagId === 'UNMAPPED');
                    const targetIndex = firstUnmapped >= 0 ? firstUnmapped : 0;
                    focusMappingRow(targetIndex);
                }
            });
        });
    }

    function focusMappingRow(colIndex) {
        if (!mappingTableBody || Number.isNaN(colIndex)) return;

        let targetRow = mappingTableBody.querySelector(`tr[data-col-index="${colIndex}"]`);
        if (!targetRow && state.surveyMapping.filter.issuesOnly) {
            state.surveyMapping.filter.issuesOnly = false;
            if (showIssueOnly) {
                showIssueOnly.checked = false;
            }
            renderSurveyMappingTable();
            targetRow = mappingTableBody.querySelector(`tr[data-col-index="${colIndex}"]`);
        }

        if (!targetRow) return;

        targetRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
        targetRow.classList.remove('mapping-row-highlight');
        void targetRow.offsetWidth;
        targetRow.classList.add('mapping-row-highlight');
        setTimeout(() => {
            targetRow.classList.remove('mapping-row-highlight');
        }, 1600);
    }

    function buildColumnIndicesFromTagMapping(tagToColumns) {
        const first = (tagId) => tagToColumns.has(tagId) ? tagToColumns.get(tagId)[0] : null;
        const indices = createEmptyColumnIndices();

        indices.email = first('EMAIL');
        indices.participation = first('PARTICIPATION');
        indices.count = first('COUNT');
        if (indices.participation === null) {
            indices.participation = indices.count;
        }
        indices.preferredMonth = first('PREF_MONTH');
        indices.preferredDay = first('PREF_DAY');
        indices.preferredDates = first('PREF_DATE_ONLY');
        indices.ngDates = first('NG_DATE');
        indices.ngDays = first('NG_DAY');
        indices.ngMonths = first('NG_MONTH');
        indices.pref1 = first('LOC_1');
        indices.pref2 = first('LOC_2');
        indices.pref3 = first('LOC_3');
        indices.pref4 = first('LOC_4');
        indices.pref5 = first('LOC_5');
        indices.additionalSupport = first('ADD_SUPPORT');
        indices.additionalCount = first('ADD_COUNT');
        indices.grade = first('GRADE');
        indices.fullName = first('NAME');
        indices.lastName = first('LAST_NAME');
        indices.firstName = first('FIRST_NAME');
        indices.classNum = first('CLASS');
        indices.freeText = first('FREE_TEXT');

        return indices;
    }

    function onMappingJsonSelected(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = () => {
            try {
                const payload = JSON.parse(reader.result);
                applyMappingJsonPayload(payload);
            } catch (error) {
                alert(`マッピングJSONの読み込みに失敗しました: ${error.message}`);
            } finally {
                if (mappingJsonFile) {
                    mappingJsonFile.value = '';
                }
            }
        };
        reader.readAsText(file, 'utf-8');
    }

    function applyMappingJsonPayload(payload) {
        if (!payload || !Array.isArray(payload.columns)) {
            alert('マッピングJSON形式が不正です。');
            return;
        }
        if (!state.surveyMapping.headers || state.surveyMapping.headers.length === 0) {
            alert('先にアンケートファイルを読み込んでください。');
            return;
        }

        const applied = applyMappingColumns(payload.columns);
        if (!applied) {
            alert('マッピングJSONの適用に失敗しました。');
            return;
        }

        saveCurrentMappingToStorage();
        renderMappingRestoreNotice({
            message: 'マッピングJSONを適用しました。前回設定として保存しました。',
            showRestoreButton: false
        });
    }

    function applyMappingColumns(columns) {
        if (!Array.isArray(columns) || !state.surveyMapping.headers || state.surveyMapping.headers.length === 0) {
            return false;
        }

        const nextTags = new Array(state.surveyMapping.headers.length).fill('UNMAPPED');
        columns.forEach(col => {
            if (typeof col.colIndex !== 'number') return;
            if (col.colIndex < 0 || col.colIndex >= nextTags.length) return;
            if (!SURVEY_TAG_MAP.has(col.tagId)) return;
            nextTags[col.colIndex] = col.tagId;
        });

        applyMappingTags(nextTags);
        return true;
    }

    function downloadSurveyMappingJson() {
        if (!state.surveyMapping.headers || state.surveyMapping.headers.length === 0) {
            alert('先にアンケートファイルを読み込んでください。');
            return;
        }

        const payload = buildCurrentMappingSnapshot();
        if (!payload) return;

        const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `survey_mapping_${formatDateTimeTag()}.json`;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
    }

    function formatDateTimeTag() {
        const now = new Date();
        const y = now.getFullYear();
        const m = String(now.getMonth() + 1).padStart(2, '0');
        const d = String(now.getDate()).padStart(2, '0');
        const hh = String(now.getHours()).padStart(2, '0');
        const mm = String(now.getMinutes()).padStart(2, '0');
        return `${y}${m}${d}_${hh}${mm}`;
    }

    // ========================================================================
    // Preview Display
    // ========================================================================
    function evaluateAssignmentReadiness() {
        state.assignmentBlockReason = '';

        if (!state.shiftData) {
            state.assignmentBlockReason = 'シフトファイルを読み込んでください';
            return {
                isReady: false,
                reason: state.assignmentBlockReason
            };
        }

        if (!state.surveyData) {
            state.assignmentBlockReason = 'アンケートファイルを読み込んでください';
            return {
                isReady: false,
                reason: state.assignmentBlockReason
            };
        }

        if (!state.surveyMapping.isValid || !state.surveyMapping.resolvedIndices) {
            state.assignmentBlockReason = '設問タグマッピングを完了してください';
            return {
                isReady: false,
                reason: state.assignmentBlockReason
            };
        }

        const validationResult = validateLocationMatching();
        if (!validationResult.isValid) {
            state.assignmentBlockReason = '場所名不一致を解消してください';
            return {
                isReady: false,
                reason: state.assignmentBlockReason
            };
        }

        if (state.isAssigning) {
            state.assignmentBlockReason = '割当て処理中です';
            return {
                isReady: false,
                reason: state.assignmentBlockReason
            };
        }

        return {
            isReady: true,
            reason: ''
        };
    }

    function updateAssignmentButtonState() {
        const readiness = evaluateAssignmentReadiness();

        if (assignBtn) {
            assignBtn.disabled = !readiness.isReady;
        }

        if (assignBlockedReason) {
            if (readiness.isReady) {
                assignBlockedReason.textContent = '実行可能です';
                assignBlockedReason.classList.add('is-ready');
            } else {
                assignBlockedReason.textContent = `実行前に対応: ${readiness.reason}`;
                assignBlockedReason.classList.remove('is-ready');
            }
        }

        return readiness;
    }

    function checkBothFilesLoaded() {
        updateAssignmentButtonState();

        if (!(state.shiftData && state.surveyData)) {
            if (previewSection) previewSection.classList.add('hidden');
            if (assignSection) assignSection.classList.add('hidden');
            return;
        }

        if (!state.surveyMapping.isValid || !state.surveyMapping.resolvedIndices) {
            if (previewSection) previewSection.classList.add('hidden');
            if (assignSection) assignSection.classList.add('hidden');
            return;
        }

        // 場所名のマッチング検証
        const validationResult = validateLocationMatching();

        if (!validationResult.isValid) {
            displayValidationError(validationResult);
            updateAssignmentButtonState();
            return;
        }

        clearLocationValidationError();

        // セクションを表示
        const preview = document.getElementById('previewSection');
        const assign = document.getElementById('assignSection');

        if (preview) {
            preview.classList.remove('hidden');
            preview.style.display = 'block';
        }
        if (assign) {
            assign.classList.remove('hidden');
            assign.style.display = 'block';
        }

        displayPreviews();
        updateAssignmentButtonState();
    }

    // 場所名のマッチング検証
    // 定員表記「（N人体制）」を除いた基本名でマッチングを行う
    function validateLocationMatching() {
        const headers = state.surveyData[0];
        const colIndices = getColumnIndices(headers);

        // シフト表の場所名セット（基本名を使用）
        const shiftBaseNames = new Set(state.locations.map(loc => loc.baseName || loc.name));
        // 基本名→元の名前のマッピング
        const baseNameToFullName = {};
        state.locations.forEach(loc => {
            baseNameToFullName[loc.baseName || loc.name] = loc.name;
        });

        // アンケートの希望場所を収集
        const surveyLocations = new Set();
        for (let i = 1; i < Math.min(100, state.surveyData.length); i++) {
            const row = state.surveyData[i];
            [colIndices.pref1, colIndices.pref2, colIndices.pref3].forEach(idx => {
                if (idx !== null && row[idx]) {
                    const loc = row[idx].toString().trim();
                    if (loc) surveyLocations.add(loc);
                }
            });
        }

        // マッチしない場所を検出（基本名で比較）
        const unmatchedSurvey = [...surveyLocations].filter(loc => !shiftBaseNames.has(loc));
        const unmatchedShift = [...shiftBaseNames].filter(baseName => !surveyLocations.has(baseName));

        return {
            isValid: unmatchedSurvey.length === 0,
            shiftLocations: [...shiftBaseNames],  // 基本名を返す
            surveyLocations: [...surveyLocations],
            unmatchedSurvey: unmatchedSurvey,
            unmatchedShift: unmatchedShift
        };
    }


    // エラー表示
    function displayValidationError(result) {
        const previewSection = document.getElementById('previewSection');
        previewSection.classList.remove('hidden');
        previewSection.style.display = 'block';

        const previewBody = previewSection.querySelector('.card-body');
        if (!previewBody) return;

        let errorHost = document.getElementById('locationValidationError');
        if (!errorHost) {
            errorHost = document.createElement('div');
            errorHost.id = 'locationValidationError';
            errorHost.style.marginBottom = '12px';
            previewBody.prepend(errorHost);
        }

        let html = `
            <div style="border: 1px solid #fecaca; border-left: 4px solid #e74c3c; background: #fff; border-radius: 10px; padding: 14px;">
                <h3 style="color: #e74c3c; margin: 0 0 8px 0;">⚠️ 場所名の不一致エラー</h3>
                <p style="margin-bottom: 12px;">アンケートの希望場所と旗振りシフト日と場所の場所名が一致しません。入力データを修正してください。</p>
                <h4 style="margin: 12px 0 8px 0;">❌ マッチしないアンケートの場所名</h4>
                <ul style="background: #fff5f5; padding: 10px 24px; border-radius: 8px;">
        `;

        result.unmatchedSurvey.forEach(loc => {
            html += `<li style="color: #e74c3c;"><strong>"${loc}"</strong></li>`;
        });

        html += `
                    </ul>
                    
                    <h4 style="margin: 16px 0 8px 0;">📋 旗振りシフト日と場所の場所名（正解）</h4>
                    <ul style="background: #f0fff0; padding: 10px 24px; border-radius: 8px;">
        `;

        result.shiftLocations.forEach(loc => {
            html += `<li>"${loc}"</li>`;
        });

        html += `
                    </ul>
                    
                    <p style="margin-top: 12px;"><strong>修正方法:</strong> アンケートの希望場所の選択肢を、旗振りシフト日と場所のヘッダー行と完全に一致するよう修正してください。</p>
                </div>
            </div>
        `;

        errorHost.innerHTML = html;
    }

    function clearLocationValidationError() {
        const host = document.getElementById('locationValidationError');
        if (host) {
            host.remove();
        }
    }

    function displayPreviews() {
        displayShiftPreview();
        displaySurveyPreview();
        setupTabs();
    }

    function displayShiftPreview() {
        const shiftStats = document.getElementById('shiftStats');
        const shiftTable = document.getElementById('shiftTable');

        // 統計情報（定員を考慮）
        const workDays = state.dates.filter(d => !d.isHoliday).length;
        // 総枠数 = 各地点の定員 × 登校日数
        const totalCapacity = state.locations.reduce((sum, loc) => sum + (loc.capacity || 1), 0);
        const totalSlots = workDays * totalCapacity;

        // 既に入力済みの枠数をカウント（登校日のみ、定員を考慮）
        let filledSlots = 0;
        state.dates.forEach(dateInfo => {
            if (dateInfo.isHoliday) return;
            const row = state.shiftData[dateInfo.rowIndex];
            state.locations.forEach(loc => {
                const value = row[loc.index] ? row[loc.index].toString().trim() : '';
                if (value !== '') {
                    // 入力されている名前の数をカウント（カンマや改行で区切られている場合）
                    const names = value.split(/[,、\n\r]+/).filter(n => n.trim());
                    filledSlots += Math.min(names.length, loc.capacity || 1);
                }
            });
        });

        const emptySlots = totalSlots - filledSlots;

        // 複数人体制の地点があるかどうか
        const hasMultiCapacity = state.locations.some(loc => (loc.capacity || 1) > 1);
        const capacityNote = hasMultiCapacity ?
            ` <span style="font-size: 0.85em; color: #666;">(複数人体制含む)</span>` : '';

        shiftStats.innerHTML = `
            <div class="stat-item">📅 日数: <strong>${state.dates.length}日</strong></div>
            <div class="stat-item">🏫 登校日: <strong>${workDays}日</strong></div>
            <div class="stat-item">📍 場所: <strong>${state.locations.length}箇所</strong>${capacityNote}</div>
            <div class="stat-item">🎯 総枠数: <strong>${totalSlots}</strong> (済: ${filledSlots} / 残: <span style="color: #e74c3c; font-weight: bold;">${emptySlots}</span>)</div>
        `;


        // テーブル表示
        let html = '<thead><tr><th>日付</th><th>曜日</th>';
        state.locations.forEach(loc => {
            html += `<th>${loc.name}</th>`;
        });
        html += '</tr></thead><tbody>';

        const displayRows = state.dates; // 全件表示
        displayRows.forEach(dateInfo => {
            const row = state.shiftData[dateInfo.rowIndex];
            const rowClass = dateInfo.isHoliday ? 'style="background: #f5f5f5; color: #999;"' : '';
            html += `<tr ${rowClass}>`;
            html += `<td>${dateInfo.date}</td>`;
            html += `<td>${dateInfo.dayOfWeek}</td>`;
            state.locations.forEach(loc => {
                const value = row[loc.index] || '';
                html += `<td>${value}</td>`;
            });
            html += '</tr>';
        });

        // 省略表示コードを削除

        html += '</tbody>';
        shiftTable.innerHTML = html;
    }

    function displaySurveyPreview() {
        const surveyStats = document.getElementById('surveyStats');
        const surveyTable = document.getElementById('surveyTable');
        const headers = state.surveyData[0];

        // カラムインデックスを特定
        const colIndices = getColumnIndices(headers);
        console.log('カラムインデックス:', colIndices);

        // 統計情報
        const totalResponses = state.surveyData.length - 1;
        let exemptCount = 0;
        let activeCount = 0;

        for (let i = 1; i < state.surveyData.length; i++) {
            const row = state.surveyData[i];
            const participation = row[colIndices.participation] ? row[colIndices.participation].toString() : '';
            if (participation.includes('免除')) {
                exemptCount++;
            } else if (participation) {
                activeCount++;
            }
        }

        surveyStats.innerHTML = `
            <div class="stat-item">📋 総回答数: <strong>${totalResponses}件</strong></div>
            <div class="stat-item">✅ 参加者: <strong>${activeCount}人</strong></div>
            <div class="stat-item">🚫 免除希望: <strong>${exemptCount}人</strong></div>
        `;

        // テーブル表示
        let html = '<thead><tr>';

        // 氏名列の設定（新形式: fullName, 旧形式: lastName + firstName）
        const hasFullName = colIndices.fullName !== null;
        const hasLegacyName = colIndices.lastName !== null || colIndices.firstName !== null;

        const displayCols = [
            // 氏名は特別処理（fullName または lastName+firstName）
            { label: '氏名', index: hasFullName ? colIndices.fullName : colIndices.lastName, isName: true },
            { label: '学年', index: colIndices.grade },
            { label: '参加可能回数', index: colIndices.count, isCount: true },
            { label: '参加確認', index: colIndices.participation },
            { label: '特定参加可能日', index: colIndices.preferredDates },
            { label: '第1希望', index: colIndices.pref1 },
            { label: '備考', index: colIndices.freeText }
        ].filter(col => col.index !== null || (col.isName && hasLegacyName));

        displayCols.forEach(col => {
            html += `<th>${col.label}</th>`;
        });
        html += '</tr></thead><tbody>';


        // すべての行を表示（制限なし）
        for (let i = 1; i < state.surveyData.length; i++) {
            const row = state.surveyData[i];
            html += '<tr>';
            displayCols.forEach(col => {
                let value = '';

                // 氏名列の特別処理（新形式: fullName, 旧形式: lastName + firstName）
                if (col.isName) {
                    if (hasFullName && colIndices.fullName !== null) {
                        value = row[colIndices.fullName] ? row[colIndices.fullName].toString() : '';
                    } else {
                        // 旧形式: lastName + firstName を結合
                        const lastName = colIndices.lastName !== null && row[colIndices.lastName] ? row[colIndices.lastName].toString() : '';
                        const firstName = colIndices.firstName !== null && row[colIndices.firstName] ? row[colIndices.firstName].toString() : '';
                        value = lastName + firstName;
                    }
                } else {
                    value = col.index !== null && row[col.index] ? row[col.index].toString() : '';
                }

                // 希望回数の場合、処理後の数字を表示
                if (col.isCount && value !== '') {
                    if (value.includes('免除')) {
                        value = '免除';
                    } else {
                        // 数値を抽出
                        const normalizedValue = value.replace(/[０-９]/g, s =>
                            String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
                        );
                        const numberMatch = normalizedValue.match(/(\d+)/);
                        if (numberMatch) {
                            value = numberMatch[1] + '回';
                        } else {
                            value = '⚠️ ' + value; // 数値が見つからない場合は警告マーク
                        }
                    }
                }

                html += `<td>${truncateText(value, 20)}</td>`;
            });
            html += '</tr>';
        }


        html += '</tbody>';
        surveyTable.innerHTML = html;
    }

    function getColumnIndices(headers) {
        if (state.surveyMapping && state.surveyMapping.resolvedIndices) {
            return state.surveyMapping.resolvedIndices;
        }

        console.warn('設問タグマッピングが未確定のため、列認識を実行できません。');
        return createEmptyColumnIndices();
    }


    function truncateText(text, maxLength) {
        if (text.length <= maxLength) return text;
        return text.substring(0, maxLength) + '...';
    }

    // ========================================================================
    // Tab Navigation
    // ========================================================================
    function setupTabs() {
        if (state.tabsInitialized) return;
        const tabs = document.querySelectorAll('.tab');
        tabs.forEach(tab => {
            tab.addEventListener('click', () => {
                tabs.forEach(t => t.classList.remove('active'));
                document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
                tab.classList.add('active');
                const targetId = tab.dataset.tab;
                document.getElementById(targetId).classList.add('active');
            });
        });
        state.tabsInitialized = true;
    }

    // ========================================================================
    // Assignment Algorithm
    // ========================================================================
    assignBtn.addEventListener('click', runAssignment);

    async function runAssignment() {
        const readiness = updateAssignmentButtonState();
        if (!readiness.isReady) {
            alert(readiness.reason);
            return;
        }

        state.isAssigning = true;
        updateAssignmentButtonState();

        progressContainer.classList.remove('hidden');
        progressFill.style.width = '0%';
        progressText.textContent = 'データを準備中...';

        try {
            await sleep(300);

            // Step 1: Parse survey responses
            progressFill.style.width = '20%';
            progressText.textContent = 'アンケート回答を解析中...';
            const participants = parseSurveyResponses();
            console.log('=== 参加者データ ===');
            console.log('参加者数:', participants.length);
            if (participants.length > 0) {
                console.log('参加者サンプル:', participants[0]);
            }
            await sleep(300);

            // Step 2: Prepare shift slots
            progressFill.style.width = '40%';
            progressText.textContent = 'シフト枠を準備中...';
            const slots = prepareShiftSlots();
            console.log('=== シフト枠 ===');
            console.log('総枠数:', slots.length);
            console.log('空き枠数:', slots.filter(s => !s.isPreAssigned).length);
            console.log('既割り当て枠数:', slots.filter(s => s.isPreAssigned).length);
            if (slots.length > 0) {
                console.log('枠サンプル:', slots[0]);
            }
            await sleep(300);

            // Step 3: Run assignment
            progressFill.style.width = '60%';
            progressText.textContent = 'シフトを割当て中...';

            // ロジック選択
            const logicType = document.getElementById('logicSelect').value;
            let result;

            if (logicType === 'minimum_guarantee') {
                console.log('選択されたロジック: 全員最低1回保証 (v2.2 - 免除希望以外は必ず1回以上)');
                result = assignShiftsMinimumGuarantee(participants, slots);
            } else if (logicType === 'participant_priority') {
                console.log('選択されたロジック: 参加者優先 (v2.1 - 入れる日数が少ない順)');
                result = assignShiftsParticipantPriority(participants, slots);
            } else if (logicType === 'date_order') {
                console.log('選択されたロジック: 日付順 (script_date_order)');
                result = assignShiftsDateOrder(participants, slots);
            } else {
                console.log('選択されたロジック: 標準 (埋まりにくい順)');
                result = assignShifts(participants, slots);
            }

            console.log('=== 割り当て結果 ===');
            console.log('割り当て数:', result.assignedCount);
            console.log('未割り当て数:', result.unassignedSlots.length);
            await sleep(300);

            // Step 4: Display results
            progressFill.style.width = '80%';
            progressText.textContent = '結果を表示中...';
            state.assignmentResult = result;
            displayResults(result);
            await sleep(200);

            progressFill.style.width = '100%';
            progressText.textContent = '完了！';

            resultSection.classList.remove('hidden');
            resultSection.scrollIntoView({ behavior: 'smooth' });

        } catch (error) {
            console.error('割り当てエラー:', error);
            progressText.textContent = `エラー: ${error.message}`;
        } finally {
            state.isAssigning = false;
            updateAssignmentButtonState();
        }
    }

    function parseSurveyResponses() {
        const headers = state.surveyData[0];
        const participants = [];
        const colIndices = getColumnIndices(headers);

        console.log('=== アンケート解析開始 ===');
        console.log('解析用カラムインデックス:', colIndices);
        console.log('総データ行数:', state.surveyData.length - 1);

        let skippedCount = 0;
        let exemptCount = 0;
        let zeroCountSkipped = 0;

        for (let i = 1; i < state.surveyData.length; i++) {
            const row = state.surveyData[i];
            const emailValue = colIndices.email !== null && row[colIndices.email]
                ? row[colIndices.email].toString().trim()
                : '';

            // 希望回数を取得
            const countValue = colIndices.count !== null && row[colIndices.count]
                ? row[colIndices.count].toString().trim()
                : '';

            // デバッグ: 各行のキー情報を出力（最初の10行のみ）
            if (i <= 10) {
                console.log(`行${i}: email="${emailValue}", count="${countValue}"`);
            }

            // 第1希望の場所を確認（参加意思の有無を判定）
            const pref1Value = colIndices.pref1 !== null && row[colIndices.pref1]
                ? row[colIndices.pref1].toString().trim()
                : '';

            // 空行をスキップ（メールアドレスも第1希望もない場合）
            if (!emailValue && !pref1Value) {
                skippedCount++;
                continue;
            }

            // 希望回数カラムの値を解析
            let maxAssignments = 0;

            // 希望回数に「免除」が含まれている場合、免除希望者としてスキップ
            if (countValue !== '' && countValue.includes('免除')) {
                exemptCount++;
                if (exemptCount <= 5) {
                    console.log(`行${i}: 希望回数に「${countValue}」と記入 → 免除希望者として除外`);
                }
                continue;
            }

            if (countValue !== '') {
                // 全角数字を半角に変換
                const normalizedValue = countValue.replace(/[０-９]/g, s =>
                    String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
                );

                // 「何回でも」「無制限」「週◯回」などは無制限（999）として扱う
                // 実際の間隔制限は前後NG期間で制御
                if (normalizedValue.includes('何回でも') ||
                    normalizedValue.includes('無制限') ||
                    normalizedValue.includes('週') ||
                    normalizedValue.includes('上限なし')) {
                    maxAssignments = 999;
                    if (i <= 5) {
                        console.log(`行${i}: 希望回数「${countValue}」→ 無制限（999）と解析`);
                    }
                }
                // 「3回以上」などの「以上」も無制限として扱う
                else if (normalizedValue.includes('以上')) {
                    maxAssignments = 999;
                    if (i <= 5) {
                        console.log(`行${i}: 希望回数「${countValue}」→ 無制限（999）と解析`);
                    }
                }
                // 通常の数値解析
                else {
                    // 正規表現で数値を抽出（「期間中5回」「学期中5回」「月2回」「5」などに対応）
                    const numberMatch = normalizedValue.match(/(\d+)/);

                    if (numberMatch) {
                        const countVal = parseInt(numberMatch[1]);
                        if (!isNaN(countVal) && countVal > 0) {
                            maxAssignments = countVal;

                            // デバッグ: 元の値と抽出した数値をログ出力（最初の5件のみ）
                            if (i <= 5) {
                                console.log(`行${i}: 希望回数「${countValue}」→ ${maxAssignments}回と解析`);
                            }
                        }
                    } else {
                        // 数値が見つからない場合、デバッグログ出力（最初の5件のみ）
                        if (i <= 5) {
                            console.log(`⚠️ 行${i}: 希望回数「${countValue}」から数値を抽出できませんでした`);
                        }
                    }
                }
            }

            // 希望回数が0または未入力の場合はスキップ（必須入力とする）
            if (maxAssignments === 0) {
                zeroCountSkipped++;
                if (zeroCountSkipped <= 5) {
                    console.log(`⚠️ 行${i}をスキップ: 希望回数が未入力または0です (count="${countValue}")。データ検証ツールで修正してください。`);
                }
                continue;
            }




            // 追加対応情報の解析
            let canSupportAdditional = false;
            let maxAdditionalAssignments = 0;

            if (colIndices.additionalSupport !== null && row[colIndices.additionalSupport]) {
                const val = row[colIndices.additionalSupport].toString().trim();
                // 肯定的な回答か？
                if (val && (val.includes('1') || val.includes('はい') || val.includes('可能') || val.toLowerCase().includes('yes'))) {
                    canSupportAdditional = true;
                    const numMatch = val.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).match(/(\d+)/);
                    if (numMatch) maxAdditionalAssignments = parseInt(numMatch[1]);
                }
            }

        if (colIndices.additionalCount !== null && row[colIndices.additionalCount]) {
            const val = row[colIndices.additionalCount].toString().trim();
            const numMatch = val.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).match(/(\d+)/);
            if (numMatch) {
                const count = parseInt(numMatch[1]);
                if (count > 0) {
                    canSupportAdditional = true;
                    maxAdditionalAssignments = count;
                }
            } else if (val && (val.includes('はい') || val.includes('可能') || val.toLowerCase().includes('yes'))) {
                // 誤って ADD_COUNT にマッピングされた可否設問を救済する。
                canSupportAdditional = true;
            }
        }

            // デフォルト値補完: 追加回数が明示されていない場合は参加可能回数と同じにする
            if (canSupportAdditional && maxAdditionalAssignments === 0) {
                maxAdditionalAssignments = maxAssignments;
            }

            const participant = {
                id: i,
                email: emailValue,
                maxAssignments: maxAssignments,
                currentAssignments: 0,
                preferredLocations: [
                    colIndices.pref1 !== null ? (row[colIndices.pref1] || '').toString().trim() : '',
                    colIndices.pref2 !== null ? (row[colIndices.pref2] || '').toString().trim() : '',
                    colIndices.pref3 !== null ? (row[colIndices.pref3] || '').toString().trim() : '',
                    colIndices.pref4 !== null ? (row[colIndices.pref4] || '').toString().trim() : '',
                    colIndices.pref5 !== null ? (row[colIndices.pref5] || '').toString().trim() : ''
                ].filter(loc => loc),
                canSupportAdditional: canSupportAdditional,
                maxAdditionalAssignments: maxAdditionalAssignments,
                grade: colIndices.grade !== null ? (row[colIndices.grade] || '').toString() : '',
                // 新形式: fullName, 旧形式: lastName + firstName
                fullName: colIndices.fullName !== null ? (row[colIndices.fullName] || '').toString() : '',
                lastName: colIndices.lastName !== null ? (row[colIndices.lastName] || '').toString() : '',
                firstName: colIndices.firstName !== null ? (row[colIndices.firstName] || '').toString() : '',
                classNum: colIndices.classNum !== null ? (row[colIndices.classNum] || '').toString() : '',
                preferredMonths: parsePreferredMonths(colIndices.preferredMonth !== null ? row[colIndices.preferredMonth] : ''),
                preferredDays: parsePreferredDays(colIndices.preferredDay !== null ? row[colIndices.preferredDay] : ''),
                preferredDates: parsePreferredDates(colIndices.preferredDates !== null ? row[colIndices.preferredDates] : ''),
                ngDates: parsePreferredDates(colIndices.ngDates !== null ? row[colIndices.ngDates] : ''),
                ngDays: parsePreferredDays(colIndices.ngDays !== null ? row[colIndices.ngDays] : ''),
                ngMonths: parsePreferredMonths(colIndices.ngMonths !== null ? row[colIndices.ngMonths] : ''),
                freeText: colIndices.freeText !== null ? (row[colIndices.freeText] || '').toString() : ''
            };

            // 表示名を作成
            const gradeNum = participant.grade.match(/(\d+)/);
            const gradeStr = gradeNum ? `${gradeNum[1]}年` : participant.grade;
            // 氏名の取得（新形式: fullName, 旧形式: lastName + firstName）
            const nameStr = participant.fullName
                || (participant.lastName || participant.firstName ? `${participant.lastName}${participant.firstName}` : '')
                || (participant.email.split('@')[0] || `ID:${participant.id}`);

            const baseName = `${nameStr}_${gradeStr}`;

            // 重複チェック
            const existingCount = participants.filter(p => p.displayName.startsWith(baseName)).length;
            participant.displayName = existingCount > 0 ? `${baseName}(${existingCount + 1})` : baseName;
            // 既存シフトの再読み込み時に使う照合キー（氏名単一列運用を優先）
            attachParticipantMatchKeys(participant);

            participants.push(participant);
        }

        // デバッグ: 処理結果のサマリー
        console.log('=== アンケート解析結果 ===');
        console.log('有効参加者数:', participants.length);
        console.log('空行/非回答でスキップ:', skippedCount);
        console.log('免除希望者:', exemptCount);
        console.log('回数0でスキップ:', zeroCountSkipped);

        // 参加者の詳細（最初の5人）
        console.log('参加者サンプル（最初の5人）:');
        participants.slice(0, 5).forEach(p => {
            console.log(`  ${p.displayName}: 希望${p.maxAssignments}回, 場所: ${p.preferredLocations.join(',')}`);
        });

        // デバッグ: maxAssignmentsの分布をログ出力
        const maxDistribution = {};
        participants.forEach(p => {
            const key = p.maxAssignments >= 999 ? '無制限' : p.maxAssignments.toString();
            maxDistribution[key] = (maxDistribution[key] || 0) + 1;
        });
        console.log('=== maxAssignments分布 ===', maxDistribution);

        return participants;
    }

    function normalizeTextForMatch(str) {
        if (str === null || str === undefined) return '';
        return str.toString()
            .trim()
            .replace(/[Ａ-Ｚａ-ｚ０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
            .replace(/[ \t\r\n　]+/g, '');
    }

    function stripTrailingDuplicateMarker(str) {
        // displayNameの重複サフィックス "(2)" / "（2）" を除去
        return str.replace(/[（(]\d+[）)]$/, '');
    }

    function stripGradeSuffix(str) {
        // "_2年", "_2年(2)", "_2年（2）" 形式を除去
        return str.replace(/[_＿]\d+年(?:[（(]\d+[）)])?$/, '');
    }

    function buildNameMatchKeys(name) {
        const normalized = normalizeTextForMatch(name);
        if (!normalized) return [];

        const keys = new Set();
        const add = (val) => {
            if (val) keys.add(val);
        };

        add(normalized);
        const noDup = stripTrailingDuplicateMarker(normalized);
        add(noDup);
        add(stripGradeSuffix(noDup));

        return [...keys];
    }

    function attachParticipantMatchKeys(participant) {
        const allKeys = new Set();
        const displayNameKeys = buildNameMatchKeys(participant.displayName);
        displayNameKeys.forEach(k => allKeys.add(k));

        // 氏名単一列（fullName）を最優先キーとして保持
        const fullName = participant.fullName ? participant.fullName.toString() : '';
        buildNameMatchKeys(fullName).forEach(k => allKeys.add(k));

        // 旧形式（姓・名分離）にも後方互換で対応
        const legacyName = `${participant.lastName || ''}${participant.firstName || ''}`;
        buildNameMatchKeys(legacyName).forEach(k => allKeys.add(k));

        // 氏名欠損ケースの補助キー（メールローカル部）
        const mailLocalPart = participant.email ? participant.email.toString().split('@')[0] : '';
        buildNameMatchKeys(mailLocalPart).forEach(k => allKeys.add(k));

        participant.displayNameKeySet = new Set(displayNameKeys);
        participant.matchKeySet = allKeys;
    }

    function findParticipantByAssignedName(assignedName, participants) {
        const assignedKeys = buildNameMatchKeys(assignedName);
        if (assignedKeys.length === 0) return null;

        const directMatches = [];
        const broadMatches = [];

        participants.forEach(p => {
            if (!p.matchKeySet || !p.displayNameKeySet) {
                attachParticipantMatchKeys(p);
            }

            const hasDirectMatch = assignedKeys.some(key => p.displayNameKeySet.has(key));
            if (hasDirectMatch) {
                directMatches.push(p);
                return;
            }

            const hasBroadMatch = assignedKeys.some(key => p.matchKeySet.has(key));
            if (hasBroadMatch) {
                broadMatches.push(p);
            }
        });

        if (directMatches.length === 1) return directMatches[0];
        if (directMatches.length > 1) return null;
        if (broadMatches.length === 1) return broadMatches[0];
        return null;
    }

    function initializePreAssignmentState(participants) {
        participants.forEach(p => {
            p.assignedDates = [];
            p.currentAssignments = 0;
            if (!p.matchKeySet || !p.displayNameKeySet) {
                attachParticipantMatchKeys(p);
            }
        });
    }

    function applyPreAssignedCounts(slots, participants, contextLabel) {
        let preAssignedCount = 0;

        slots.forEach(slot => {
            if (!slot.isPreAssigned || !slot.assignedTo) return;

            preAssignedCount++;
            const participant = findParticipantByAssignedName(slot.assignedTo, participants);

            if (!participant) {
                console.warn(`[${contextLabel}] 既存割り当ての照合失敗: "${slot.assignedTo}" (${slot.date} ${slot.location})`);
                return;
            }

            participant.currentAssignments++;
            participant.assignedDates.push(slot.date);
            console.log(`既存割り当て検出: ${participant.displayName} - ${slot.date} ${slot.location}`);
        });

        return preAssignedCount;
    }

    function parsePreferredMonths(str) {
        if (!str) return [];
        const months = [];
        const strVal = str.toString();
        const matches = strVal.match(/(\d+)月/g);
        if (matches) {
            matches.forEach(m => {
                const num = parseInt(m);
                if (num >= 1 && num <= 12) months.push(num);
            });
        }
        // 「1月」形式以外にも「1,2,3」形式に対応
        const numMatches = strVal.match(/\d+/g);
        if (numMatches && months.length === 0) {
            numMatches.forEach(n => {
                const num = parseInt(n);
                if (num >= 1 && num <= 12) months.push(num);
            });
        }
        return [...new Set(months)]; // 重複除去
    }

    function parsePreferredDays(str) {
        if (!str) return [];
        const days = [];
        const strVal = str.toString();
        const dayNames = ['月', '火', '水', '木', '金'];
        dayNames.forEach(day => {
            if (strVal.includes(day)) days.push(day);
        });
        return days;
    }

    // 特定の日付を解析
    function parsePreferredDates(str) {
        if (!str) return [];
        const dates = [];
        // 全角数字を半角に変換し、カンマや読点を統一的な区切り文字に置換
        const strVal = str.toString()
            .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
            .replace(/[、，\s]+/g, ',');

        // カンマで分割して個別に評価
        const segments = strVal.split(',');

        segments.forEach(segment => {
            // パターン1: YYYY/MM/DD (2026/01/15)
            let match = segment.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})/);
            if (match) {
                dates.push(`${match[1]}/${parseInt(match[2])}/${parseInt(match[3])}`);
                return;
            }

            // パターン2: MM/DD (01/15) -> 年は現在の翌年などを推測した方が安全だが、一旦そのまま
            match = segment.match(/(\d{1,2})\/(\d{1,2})/);
            if (match) {
                // 年が省略されている場合、入力データや現在の日時から補完すべきだが
                // シンプルに M/D 形式で保存し、マッチング時に部分一致させるか
                // あるいはシフト表の日付（YYYY/M/D）から年を取得して補完する
                // ここでは暫定的に「月/日」として保存
                dates.push(`${parseInt(match[1])}/${parseInt(match[2])}`);
                return;
            }

            // パターン3: M月D日
            match = segment.match(/(\d{1,2})月(\d{1,2})日/);
            if (match) {
                dates.push(`${parseInt(match[1])}/${parseInt(match[2])}`);
            }
        });

        return [...new Set(dates)]; // 重複除去
    }

    function prepareShiftSlots() {
        const slots = [];

        state.dates.forEach(dateInfo => {
            // 休日はスキップ
            if (dateInfo.isHoliday) return;

            state.locations.forEach(location => {
                const row = state.shiftData[dateInfo.rowIndex];
                const existingValue = row[location.index] ? row[location.index].toString().trim() : '';
                const capacity = location.capacity || 1;

                // 既存の名前を解析（カンマ、読点、改行で区切り）
                const existingNames = existingValue ?
                    existingValue.split(/[,、\n\r]+/).map(n => n.trim()).filter(n => n) : [];

                // 定員分のスロットを生成
                for (let slotIndex = 0; slotIndex < capacity; slotIndex++) {
                    const assignedName = existingNames[slotIndex] || null;
                    const isPreAssigned = !!assignedName;

                    slots.push({
                        date: dateInfo.date,
                        dayOfWeek: dateInfo.dayOfWeek,
                        month: extractMonth(dateInfo.date),
                        location: location.name,          // 定員表記含む元の名前
                        baseName: location.baseName || location.name, // マッチング用基本名
                        locationIndex: location.index,
                        rowIndex: dateInfo.rowIndex,
                        capacity: capacity,               // この地点の定員
                        slotIndex: slotIndex,             // このスロットのインデックス（0始まり）
                        assignedTo: isPreAssigned ? assignedName : null,
                        isPreAssigned: isPreAssigned,
                        originalValue: existingValue      // 元の値を保持
                    });
                }
            });
        });

        // デバッグログ
        const multiCapacitySlots = slots.filter(s => s.capacity > 1);
        if (multiCapacitySlots.length > 0) {
            console.log(`📊 複数人体制スロット数: ${multiCapacitySlots.length} (総スロット: ${slots.length})`);
        }

        return slots;
    }


    // 日付の正確な比較用ヘルパー関数
    // ゼロパディングに依存しない（2025/04/07 と 2025/4/7 は同一と判定）
    function isDateMatch(assignedDate, preferredDate) {
        const assignedParts = assignedDate.split('/').map(p => parseInt(p, 10));
        const preferredParts = preferredDate.split('/').map(p => parseInt(p, 10));

        if (preferredParts.length === 3) {
            // 年月日すべて指定 → 年月日を数値で比較
            return assignedParts.length >= 3 &&
                assignedParts[0] === preferredParts[0] &&
                assignedParts[1] === preferredParts[1] &&
                assignedParts[2] === preferredParts[2];
        } else if (preferredParts.length === 2) {
            // 月日のみ指定 → 月と日を数値で比較
            return assignedParts.length >= 2 &&
                assignedParts[assignedParts.length - 2] === preferredParts[0] &&
                assignedParts[assignedParts.length - 1] === preferredParts[1];
        }
        return false;
    }

    function extractMonth(dateStr) {
        // "2026/1/5" や "1/5" など様々な形式に対応
        const match = dateStr.match(/(\d{4})\/(\d+)\//) || dateStr.match(/^(\d+)\//);
        if (match) {
            // "2026/1/5" の場合は month は match[2]
            // "1/5" の場合は match[1]
            const monthIndex = match.length > 2 ? 2 : 1;
            return parseInt(match[monthIndex]);
        }
        return null;
    }

    // ===========================================
    // v3.0新機能: 日付間隔計算（前後NG期間制約用）
    // ===========================================
    /**
     * 2つの日付文字列の間隔（日数）を計算する
     * @param {string} dateStr1 - 日付文字列 (例: "2026/1/5")
     * @param {string} dateStr2 - 日付文字列 (例: "2026/1/8")
     * @returns {number} 日付間隔の絶対値（日数）、計算不能の場合は999
     */
    function getDaysDifference(dateStr1, dateStr2) {
        try {
            // 日付文字列をパース (YYYY/M/D または M/D 形式)
            const parseDate = (str) => {
                const parts = str.split('/').map(p => parseInt(p, 10));
                if (parts.length === 3) {
                    // YYYY/M/D
                    return new Date(parts[0], parts[1] - 1, parts[2]);
                } else if (parts.length === 2) {
                    // M/D → 現在の年を使用
                    const now = new Date();
                    return new Date(now.getFullYear(), parts[0] - 1, parts[1]);
                }
                return null;
            };

            const date1 = parseDate(dateStr1);
            const date2 = parseDate(dateStr2);

            if (!date1 || !date2) return 999; // パース失敗時は制約なしとして扱う

            const diffTime = Math.abs(date2 - date1);
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
            return diffDays;
        } catch (e) {
            console.warn('日付差分計算エラー:', e, dateStr1, dateStr2);
            return 999; // エラー時は制約なしとして扱う
        }
    }

    /**
     * 参加者の既割り当て日と候補スロットの日付が前後NG期間内かをチェック
     * @param {object} participant - 参加者オブジェクト（assignedDatesを持つ）
     * @param {string} slotDate - チェックするスロットの日付
     * @returns {boolean} NG期間内ならtrue（割り当て不可）
     */
    function isWithinNgPeriod(participant, slotDate) {
        const ngPeriod = getUserNgPeriodDays();

        // NG期間が0の場合は制約無効
        if (ngPeriod === 0) {
            return false;
        }

        if (!participant.assignedDates || participant.assignedDates.length === 0) {
            return false; // 割り当て済み日がなければOK
        }

        for (const assignedDate of participant.assignedDates) {
            const daysDiff = getDaysDifference(assignedDate, slotDate);
            if (daysDiff < ngPeriod && daysDiff > 0) {
                // 同日（daysDiff === 0）は別のチェックで除外されるため、ここでは1日以上の差のみチェック
                console.log(`[NG期間] ${participant.displayName}: ${assignedDate}から${slotDate}まで${daysDiff}日 → NG（${ngPeriod}日未満）`);
                return true; // NG期間内
            }
        }
        return false; // NG期間外（割り当て可能）
    }

    // 参加者が入れる日数を計算（v2.1新機能）
    // ===========================================
    function calculateAvailableDays(participant, slots) {
        // 参加者の希望条件に合致するスロット（日付）の数を計算
        // これにより「選択肢が少ない人」を定量的に評価できる

        const availableSlots = new Set(); // 重複を避けるためSetを使用

        slots.forEach(slot => {
            // 既に埋まっているスロットはカウントしない
            if (slot.isPreAssigned) return;

            // 場所のマッチング（必須条件）
            if (!participant.preferredLocations || participant.preferredLocations.length === 0) return;
            const locationMatch = participant.preferredLocations.includes(slot.location);
            if (!locationMatch) return;

            // 【最優先】特定の日付が指定されている場合
            if (participant.preferredDates && participant.preferredDates.length > 0) {
                const dateMatch = participant.preferredDates.some(prefDate => isDateMatch(slot.date, prefDate));
                if (dateMatch) {
                    availableSlots.add(slot.date);
                }
                return; // 特定日指定がある場合は、月・曜日の条件は無視
            }

            // 月の条件チェック
            const monthMatch = participant.preferredMonths.length === 0 ||
                participant.preferredMonths.includes(slot.month);

            // 曜日の条件チェック
            const dayMatch = participant.preferredDays.length === 0 ||
                participant.preferredDays.some(d => slot.dayOfWeek.includes(d));

            if (monthMatch && dayMatch) {
                availableSlots.add(slot.date);
            }
        });

        return availableSlots.size;
    }

    // ===========================================
    // 参加者の制約レベルを評価（v2.1新機能）
    // ===========================================
    function calculateConstraintLevel(participant) {
        // 制約の厳しさを数値化
        // 値が大きいほど制約が厳しい（優先度が高い）

        let level = 0;

        // 特定日指定（最も制約が厳しい）
        if (participant.preferredDates && participant.preferredDates.length > 0) {
            level += 1000; // 特定日指定は最優先
            level += (10 - participant.preferredDates.length) * 100; // 日数が少ないほど厳しい
            return level;
        }

        // 曜日指定
        if (participant.preferredDays && participant.preferredDays.length > 0) {
            level += 100;
            level += (5 - participant.preferredDays.length) * 10; // 曜日の選択肢が少ないほど厳しい
        }

        // 月指定
        if (participant.preferredMonths && participant.preferredMonths.length > 0) {
            level += 50;
            level += (12 - participant.preferredMonths.length) * 2; // 月の選択肢が少ないほど厳しい
        }

        // 場所の選択肢が少ないほど厳しい
        if (participant.preferredLocations && participant.preferredLocations.length > 0) {
            level += (3 - participant.preferredLocations.length) * 5;
        }

        return level;
    }

    function assignShifts(participants, slots) {
        let assignedCount = 0;
        const unassignedSlots = [];

        // 参加者がいない場合
        if (participants.length === 0) {
            console.warn('参加者が0人です');
            return {
                slots: slots,
                participants: participants,
                assignedCount: 0,
                unassignedSlots: slots,
                totalSlots: slots.length,
                preAssignedCount: 0
            };
        }

        initializePreAssignmentState(participants);
        const preAssignedCount = applyPreAssignedCounts(slots, participants, '標準ロジック');

        console.log(`既存の割り当て数: ${preAssignedCount}件`);
        console.log('参加者の既割り当て状況:');
        participants.filter(p => p.currentAssignments > 0).forEach(p => {
            console.log(`  ${p.displayName}: ${p.currentAssignments}/${p.maxAssignments}回`);
        });

        // 参加者をシャッフル（公平性のため初期順序をランダム化）
        const shuffledParticipants = [...participants].sort(() => Math.random() - 0.5);
        // 希望回数が少ない順 → 高学年順にソート
        shuffledParticipants.sort((a, b) => {
            // 第1優先: 希望回数が少ない順
            if (a.maxAssignments !== b.maxAssignments) {
                return a.maxAssignments - b.maxAssignments;
            }
            // 第2優先: 高学年優先
            const gradeA = parseInt(a.grade.match(/(\d+)/)?.[1] || '0');
            const gradeB = parseInt(b.grade.match(/(\d+)/)?.[1] || '0');
            return gradeB - gradeA; // 降順（高学年優先）
        });

        // ===========================================
        // 改善点: 競合倍率の高いスロットから順に処理する
        // ===========================================

        // 1. 全スロットに対して候補者数を計算（既に埋まっているスロットは除外）
        const unassignedOnlySlots = slots.filter(s => !s.assignedTo);

        unassignedOnlySlots.forEach(slot => {
            // この時点では割り当て上限などを厳密に見る必要はなく、
            // 「このスロットに入れる可能性がある人」が何人いるかを数える
            const potentialCandidates = findCandidates(slot, shuffledParticipants, true);
            slot.demandScore = potentialCandidates.length;
        });

        // 2. 候補者が少ない（＝人気がない、または条件が厳しい）スロット順にソート
        // これにより、なり手が少ないスロットを確実に埋めにいく
        // 逆に「人気スロット」は候補者が多いので後回しでも埋まる可能性が高い
        const sortedSlots = [...unassignedOnlySlots].sort((a, b) => {
            // まず候補者数で比較（少ない順）
            const demandDiff = a.demandScore - b.demandScore;
            if (demandDiff !== 0) return demandDiff;
            // 候補者数が同じ場合は、日付順で処理（早い日付を優先）
            return a.rowIndex - b.rowIndex;
        });

        console.log(`割り当て対象スロット: ${sortedSlots.length}件`);
        if (sortedSlots.length > 0) {
            console.log(`候補者数分布: 最小=${sortedSlots[0].demandScore}, 最大=${sortedSlots[sortedSlots.length - 1].demandScore}`);
        }

        // 3. ソートされた順序で割り当て実行
        sortedSlots.forEach(slot => {
            const candidates = findCandidates(slot, shuffledParticipants);

            if (candidates.length > 0) {
                const selected = selectBestCandidate(candidates, slot);
                slot.assignedTo = selected.displayName;
                selected.currentAssignments++;
                selected.assignedDates.push(slot.date); // 割り当て日を記録
                assignedCount++;
            } else {
                unassignedSlots.push(slot);
            }
        });

        // ===========================================
        // 第2パス: 柔軟な条件で未割り当てを減らす
        // ===========================================
        // 第1パスで埋まらなかったスロットに対して、
        // 希望回数に余裕がある人や「追加対応可」の人を対象に再度割り当てを試みる

        const stillUnassignedSlots = slots.filter(s => !s.assignedTo);

        if (stillUnassignedSlots.length > 0) {
            console.log(`\n=== 第2パス開始: 残り${stillUnassignedSlots.length}スロットに柔軟な条件で割り当て試行 ===`);

            // 希望回数にまだ余裕がある人を抽出
            const flexibleParticipants = shuffledParticipants.filter(p => {
                return p.currentAssignments < p.maxAssignments || p.canSupportAdditional;
            });

            console.log(`柔軟な割り当て候補者: ${flexibleParticipants.length}人`);

            let secondPassCount = 0;

            // 候補者が少ない順に再処理
            const resortedSlots = [...stillUnassignedSlots].sort((a, b) => {
                const aCandidates = findCandidatesFlexible(a, flexibleParticipants, true);
                const bCandidates = findCandidatesFlexible(b, flexibleParticipants, true);
                return aCandidates.length - bCandidates.length;
            });

            resortedSlots.forEach(slot => {
                // 候補者を再フィルタリング（希望回数上限をリアルタイムでチェック）
                const candidates = findCandidatesFlexible(slot, flexibleParticipants, false).filter(p => {
                    if (p.canSupportAdditional) {
                        const additionalLimit = p.maxAdditionalAssignments || 1;
                        return p.currentAssignments < p.maxAssignments + additionalLimit;
                    }
                    return p.currentAssignments < p.maxAssignments;
                });

                if (candidates.length > 0) {
                    const selected = selectBestCandidate(candidates, slot);
                    slot.assignedTo = selected.displayName;
                    selected.currentAssignments++;
                    selected.assignedDates.push(slot.date);
                    assignedCount++;
                    secondPassCount++;
                }
            });

            console.log(`第2パスで追加割り当て: ${secondPassCount}件`);
        }

        const remainingUnassigned = slots.filter(s => !s.assignedTo);

        console.log(`\n=== 最終結果 ===`);
        console.log(`総割り当て数: ${assignedCount}件`);
        console.log(`未割り当て数: ${remainingUnassigned.length}件`);

        return {
            slots: slots, // 元の配列を返す（表示順序維持のため）
            participants: shuffledParticipants,
            assignedCount: assignedCount,
            unassignedSlots: remainingUnassigned,
            totalSlots: slots.length,
            preAssignedCount: preAssignedCount
        };
    }

    // ===========================================
    // 柔軟な条件での候補者検索（第2パス用）
    // ===========================================
    function findCandidatesFlexible(slot, participants, isDryRun = false) {
        return participants.filter(p => {
            // 割り当て上限チェック（追加対応可の場合は+1まで許可）
            if (!isDryRun) {
                if (p.canSupportAdditional) {
                    // 追加対応可の人は、指定された追加回数（なければ無制限）まで許可
                    const additionalLimit = p.maxAdditionalAssignments || 99;
                    if (p.currentAssignments >= p.maxAssignments + additionalLimit) return false;
                } else {
                    if (p.currentAssignments >= p.maxAssignments) return false;
                }
            }

            // 同一日の重複チェック
            if (!isDryRun && p.assignedDates && p.assignedDates.includes(slot.date)) return false;

            // 【v3.0新機能】前後NG期間チェック（ドライラン時は無視）
            if (!isDryRun && isWithinNgPeriod(p, slot.date)) return false;

            // 【最優先】NG条件チェック（絶対条件）
            if (p.ngDates && p.ngDates.length > 0) {
                const isNgDate = p.ngDates.some(ngDate => isDateMatch(slot.date, ngDate));
                if (isNgDate) return false;
            }
            if (p.ngDays && p.ngDays.length > 0) {
                const isNgDay = p.ngDays.some(d => slot.dayOfWeek.includes(d));
                if (isNgDay) return false;
            }
            if (p.ngMonths && p.ngMonths.length > 0) {
                const isNgMonth = p.ngMonths.includes(slot.month);
                if (isNgMonth) return false;
            }

            // 場所マッチング（完全一致）
            // 定員表記を除いた基本名で比較（baseNameがあればそれを使用）
            if (!p.preferredLocations || p.preferredLocations.length === 0) return false;
            const slotBaseName = slot.baseName || slot.location;
            const locationMatch = p.preferredLocations.includes(slotBaseName);
            if (!locationMatch) return false;

            // 特定日指定がある場合は、その日付のみ許可（厳格）
            if (p.preferredDates && p.preferredDates.length > 0) {
                const dateMatch = p.preferredDates.some(prefDate => isDateMatch(slot.date, prefDate));
                return dateMatch;
            }

            // 月・曜日の条件チェック
            // 参加可能月の指定がある場合のみチェック（空の場合は互換性のため全OKとして扱う）
            if (p.preferredMonths && p.preferredMonths.length > 0) {
                const monthMatch = p.preferredMonths.includes(slot.month);
                if (!monthMatch) return false;
            }

            // 参加可能曜日の指定がある場合のみチェック（空の場合は互換性のため全OKとして扱う）
            if (p.preferredDays && p.preferredDays.length > 0) {
                const dayMatch = p.preferredDays.some(d => slot.dayOfWeek.includes(d));
                if (!dayMatch) return false;
            }

            return true;
        });
    }


    // ===========================================
    // 参加者優先割り当てロジック（v2.1新機能）
    // ===========================================
    function assignShiftsParticipantPriority(participants, slots) {
        let assignedCount = 0;
        const unassignedSlots = [];

        // 参加者がいない場合
        if (participants.length === 0) {
            console.warn('参加者が0人です');
            return {
                slots: slots,
                participants: participants,
                assignedCount: 0,
                unassignedSlots: slots,
                totalSlots: slots.length,
                preAssignedCount: 0
            };
        }

        initializePreAssignmentState(participants);
        const preAssignedCount = applyPreAssignedCounts(slots, participants, '参加者優先ロジック');

        console.log(`[参加者優先ロジック] 既存の割り当て数: ${preAssignedCount}件`);
        console.log('参加者の既割り当て状況:');
        participants.filter(p => p.currentAssignments > 0).forEach(p => {
            console.log(`  ${p.displayName}: ${p.currentAssignments}/${p.maxAssignments}回`);
        });

        // ===========================================
        // v2.1の核心: 参加者の「入れる日数」を計算
        // ===========================================
        participants.forEach(p => {
            p.availableDaysCount = calculateAvailableDays(p, slots);
            p.constraintLevel = calculateConstraintLevel(p);
        });

        // デバッグ: 入れる日数と制約レベルの分布を確認
        console.log('\n=== 参加者の制約状況（上位10名） ===');
        const sortedForDebug = [...participants].sort((a, b) => {
            if (a.availableDaysCount !== b.availableDaysCount) {
                return a.availableDaysCount - b.availableDaysCount;
            }
            return b.constraintLevel - a.constraintLevel;
        });
        sortedForDebug.slice(0, 10).forEach(p => {
            console.log(`  ${p.displayName}: 入れる日数=${p.availableDaysCount}日, 制約Lv=${p.constraintLevel}, 希望=${p.maxAssignments}回`);
        });

        // ===========================================
        // 参加者を「入れる日数」でソート（少ない順）
        // 同数の場合は制約レベル、さらに希望回数、学年でソート
        // ===========================================
        const sortedParticipants = [...participants].sort((a, b) => {
            // 第1優先: 入れる日数が少ない順
            if (a.availableDaysCount !== b.availableDaysCount) {
                return a.availableDaysCount - b.availableDaysCount;
            }
            // 第2優先: 制約レベルが高い順
            if (a.constraintLevel !== b.constraintLevel) {
                return b.constraintLevel - a.constraintLevel;
            }
            // 第3優先: 希望回数が少ない順
            if (a.maxAssignments !== b.maxAssignments) {
                return a.maxAssignments - b.maxAssignments;
            }
            // 第4優先: 高学年優先
            const gradeA = parseInt(a.grade.match(/(\d+)/)?.[1] || '0');
            const gradeB = parseInt(b.grade.match(/(\d+)/)?.[1] || '0');
            if (gradeA !== gradeB) {
                return gradeB - gradeA; // 降順（高学年優先）
            }
            // 同条件の場合はランダム
            return Math.random() - 0.5;
        });

        console.log(`\n=== 割り当て開始（参加者優先モード） ===`);
        console.log(`処理順序: 入れる日数が少ない順 → 制約レベルが高い順 → 希望回数が少ない順 → 高学年優先`);

        // ===========================================
        // 参加者順に処理（制約が厳しい人から）
        // ===========================================
        let participantProcessed = 0;

        sortedParticipants.forEach(participant => {
            // 既に希望回数に達している場合はスキップ
            if (participant.currentAssignments >= participant.maxAssignments) {
                return;
            }

            // この参加者が入れるスロットを探す
            const availableSlots = slots.filter(slot => {
                // 既に埋まっているスロットはスキップ
                if (slot.assignedTo) return false;

                // 同日に既に割り当てられていないかチェック
                if (participant.assignedDates && participant.assignedDates.includes(slot.date)) return false;

                // 【v3.0新機能】前後NG期間チェック
                if (isWithinNgPeriod(participant, slot.date)) return false;

                // 【最優先】NG条件チェック（絶対条件）
                if (participant.ngDates && participant.ngDates.length > 0) {
                    const isNgDate = participant.ngDates.some(ngDate => isDateMatch(slot.date, ngDate));
                    if (isNgDate) return false;
                }
                if (participant.ngDays && participant.ngDays.length > 0) {
                    const isNgDay = participant.ngDays.some(d => slot.dayOfWeek.includes(d));
                    if (isNgDay) return false;
                }
                if (participant.ngMonths && participant.ngMonths.length > 0) {
                    const isNgMonth = participant.ngMonths.includes(slot.month);
                    if (isNgMonth) return false;
                }

                // 場所マッチング
                if (!participant.preferredLocations || participant.preferredLocations.length === 0) return false;
                if (!participant.preferredLocations.includes(slot.location)) return false;

                // 特定日指定
                if (participant.preferredDates && participant.preferredDates.length > 0) {
                    return participant.preferredDates.some(prefDate => {
                        return isDateMatch(slot.date, prefDate);
                    });
                }

                // 月・曜日の条件
                const monthMatch = participant.preferredMonths.length === 0 ||
                    participant.preferredMonths.includes(slot.month);
                const dayMatch = participant.preferredDays.length === 0 ||
                    participant.preferredDays.some(d => slot.dayOfWeek.includes(d));

                return monthMatch && dayMatch;
            });

            if (availableSlots.length === 0) {
                return;
            }

            // 希望回数まで割り当てを試みる
            const needAssignments = participant.maxAssignments - participant.currentAssignments;
            let assigned = 0;

            // スロットを日付順にソート
            const sortedSlots = availableSlots.sort((a, b) => a.rowIndex - b.rowIndex);

            for (const slot of sortedSlots) {
                if (assigned >= needAssignments) break;

                // 【v3.0修正】ループ内でNG期間を再チェック（同一ループ内での連続割り当て防止）
                if (slot.assignedTo) continue; // 既に埋まっている場合スキップ
                // ループ開始時点の候補抽出後に同日枠が残っている可能性があるため、再チェックする
                if (participant.assignedDates && participant.assignedDates.includes(slot.date)) continue;
                if (isWithinNgPeriod(participant, slot.date)) continue; // NG期間内ならスキップ

                // 割り当て実行
                slot.assignedTo = participant.displayName;
                participant.currentAssignments++;
                participant.assignedDates.push(slot.date);
                assignedCount++;
                assigned++;
            }

            if (assigned > 0) {
                participantProcessed++;
                if (participantProcessed <= 10) {
                    console.log(`  ${participant.displayName}: ${assigned}件割り当て（入れる日数=${participant.availableDaysCount}日）`);
                }
            }
        });

        console.log(`処理した参加者数: ${participantProcessed}人`);

        // ===========================================
        // 第2パス: 柔軟な条件で未割り当てを減らす
        // ===========================================
        const stillUnassignedSlots = slots.filter(s => !s.assignedTo);

        if (stillUnassignedSlots.length > 0) {
            console.log(`\n=== 第2パス開始: 残り${stillUnassignedSlots.length}スロットに柔軟な条件で割り当て試行 ===`);

            const flexibleParticipants = sortedParticipants.filter(p => {
                return p.currentAssignments < p.maxAssignments || p.canSupportAdditional;
            });

            console.log(`柔軟な割り当て候補者: ${flexibleParticipants.length}人`);

            let secondPassCount = 0;

            const resortedSlots = [...stillUnassignedSlots].sort((a, b) => {
                const aCandidates = findCandidatesFlexible(a, flexibleParticipants, true);
                const bCandidates = findCandidatesFlexible(b, flexibleParticipants, true);
                return aCandidates.length - bCandidates.length;
            });

            resortedSlots.forEach(slot => {
                // 候補者を再フィルタリング（希望回数上限をリアルタイムでチェック）
                const candidates = findCandidatesFlexible(slot, flexibleParticipants, false).filter(p => {
                    if (p.canSupportAdditional) {
                        const additionalLimit = p.maxAdditionalAssignments || 1;
                        return p.currentAssignments < p.maxAssignments + additionalLimit;
                    }
                    return p.currentAssignments < p.maxAssignments;
                });

                if (candidates.length > 0) {
                    const selected = selectBestCandidate(candidates, slot);
                    slot.assignedTo = selected.displayName;
                    selected.currentAssignments++;
                    selected.assignedDates.push(slot.date);
                    assignedCount++;
                    secondPassCount++;
                }
            });

            console.log(`第2パスで追加割り当て: ${secondPassCount}件`);
        }

        const remainingUnassigned = slots.filter(s => !s.assignedTo);

        console.log(`\n=== 最終結果（参加者優先モード） ===`);
        console.log(`総割り当て数: ${assignedCount}件`);
        console.log(`未割り当て数: ${remainingUnassigned.length}件`);

        return {
            slots: slots,
            participants: sortedParticipants,
            assignedCount: assignedCount,
            unassignedSlots: remainingUnassigned,
            totalSlots: slots.length,
            preAssignedCount: preAssignedCount
        };
    }


    // ===========================================
    // 全員最低1回保証割り当てロジック（v2.2新機能）
    // ===========================================
    function assignShiftsMinimumGuarantee(participants, slots) {
        let assignedCount = 0;
        const unassignedSlots = [];

        // 参加者がいない場合
        if (participants.length === 0) {
            console.warn('参加者が0人です');
            return {
                slots: slots,
                participants: participants,
                assignedCount: 0,
                unassignedSlots: slots,
                totalSlots: slots.length,
                preAssignedCount: 0
            };
        }

        initializePreAssignmentState(participants);
        const preAssignedCount = applyPreAssignedCounts(slots, participants, '全員最低1回保証ロジック');

        console.log(`[全員最低1回保証ロジック] 既存の割り当て数: ${preAssignedCount}件`);
        console.log('参加者の既割り当て状況:');
        participants.filter(p => p.currentAssignments > 0).forEach(p => {
            console.log(`  ${p.displayName}: ${p.currentAssignments}/${p.maxAssignments}回`);
        });

        // 参加者の「入れる日数」と「制約レベル」を計算
        participants.forEach(p => {
            p.availableDaysCount = calculateAvailableDays(p, slots);
            p.constraintLevel = calculateConstraintLevel(p);
        });

        // 参加者を制約が厳しい順にソート
        const sortedParticipants = [...participants].sort((a, b) => {
            // 第1優先: 入れる日数が少ない順
            if (a.availableDaysCount !== b.availableDaysCount) {
                return a.availableDaysCount - b.availableDaysCount;
            }
            // 第2優先: 制約レベルが高い順
            if (a.constraintLevel !== b.constraintLevel) {
                return b.constraintLevel - a.constraintLevel;
            }
            // 第3優先: 希望回数が少ない順
            if (a.maxAssignments !== b.maxAssignments) {
                return a.maxAssignments - b.maxAssignments;
            }
            // 第4優先: 高学年優先
            const gradeA = parseInt(a.grade.match(/(\d+)/)?.[1] || '0');
            const gradeB = parseInt(b.grade.match(/(\d+)/)?.[1] || '0');
            if (gradeA !== gradeB) {
                return gradeB - gradeA; // 降順（高学年優先）
            }
            // 同条件の場合はランダム
            return Math.random() - 0.5;
        });

        console.log(`\n=== 第1パス: 全員に最低1回保証 ===`);
        console.log(`対象参加者: ${sortedParticipants.length}人`);

        // ===========================================
        // 第1パス: 全員に最低1回を割り当て
        // ===========================================
        let guaranteedCount = 0;
        const needsGuarantee = sortedParticipants.filter(p => p.currentAssignments === 0);

        console.log(`最低1回の割り当てが必要な人数: ${needsGuarantee.length}人`);

        needsGuarantee.forEach(participant => {
            // この参加者が入れるスロットを探す
            const availableSlots = slots.filter(slot => {
                if (slot.assignedTo) return false;
                if (participant.assignedDates && participant.assignedDates.includes(slot.date)) return false;

                // 【v3.0新機能】前後NG期間チェック
                if (isWithinNgPeriod(participant, slot.date)) return false;

                // 【最優先】NG条件チェック（絶対条件）
                if (participant.ngDates && participant.ngDates.length > 0) {
                    const isNgDate = participant.ngDates.some(ngDate => isDateMatch(slot.date, ngDate));
                    if (isNgDate) return false;
                }
                if (participant.ngDays && participant.ngDays.length > 0) {
                    const isNgDay = participant.ngDays.some(d => slot.dayOfWeek.includes(d));
                    if (isNgDay) return false;
                }
                if (participant.ngMonths && participant.ngMonths.length > 0) {
                    const isNgMonth = participant.ngMonths.includes(slot.month);
                    if (isNgMonth) return false;
                }

                if (!participant.preferredLocations || participant.preferredLocations.length === 0) return false;
                if (!participant.preferredLocations.includes(slot.location)) return false;

                // 特定日指定
                if (participant.preferredDates && participant.preferredDates.length > 0) {
                    return participant.preferredDates.some(prefDate => {
                        return isDateMatch(slot.date, prefDate);
                    });
                }

                const monthMatch = participant.preferredMonths.length === 0 ||
                    participant.preferredMonths.includes(slot.month);
                const dayMatch = participant.preferredDays.length === 0 ||
                    participant.preferredDays.some(d => slot.dayOfWeek.includes(d));

                return monthMatch && dayMatch;
            });

            if (availableSlots.length > 0) {
                // 日付が早い順にソートして最初の1つを割り当て
                const sortedSlots = availableSlots.sort((a, b) => a.rowIndex - b.rowIndex);
                const selectedSlot = sortedSlots[0];

                selectedSlot.assignedTo = participant.displayName;
                participant.currentAssignments++;
                participant.assignedDates.push(selectedSlot.date);
                assignedCount++;
                guaranteedCount++;

                if (guaranteedCount <= 10) {
                    console.log(`  ${participant.displayName}: 1回保証割り当て完了（${selectedSlot.date} ${selectedSlot.location}）`);
                }
            } else {
                console.warn(`  ⚠️ ${participant.displayName}: 入れるスロットが見つかりませんでした（入れる日数=${participant.availableDaysCount}日）`);
            }
        });

        console.log(`第1パス完了: ${guaranteedCount}人に最低1回を保証`);

        // ===========================================
        // 第2パス: 希望回数まで追加割り当て
        // ===========================================
        console.log(`\n=== 第2パス: 希望回数まで追加割り当て ===`);

        let additionalCount = 0;
        sortedParticipants.forEach(participant => {
            if (participant.currentAssignments >= participant.maxAssignments) {
                return;
            }

            const needAssignments = participant.maxAssignments - participant.currentAssignments;

            const availableSlots = slots.filter(slot => {
                if (slot.assignedTo) return false;
                if (participant.assignedDates && participant.assignedDates.includes(slot.date)) return false;

                // 【v3.0新機能】前後NG期間チェック
                if (isWithinNgPeriod(participant, slot.date)) return false;

                // 【最優先】NG条件チェック（絶対条件）
                if (participant.ngDates && participant.ngDates.length > 0) {
                    const isNgDate = participant.ngDates.some(ngDate => isDateMatch(slot.date, ngDate));
                    if (isNgDate) return false;
                }
                if (participant.ngDays && participant.ngDays.length > 0) {
                    const isNgDay = participant.ngDays.some(d => slot.dayOfWeek.includes(d));
                    if (isNgDay) return false;
                }
                if (participant.ngMonths && participant.ngMonths.length > 0) {
                    const isNgMonth = participant.ngMonths.includes(slot.month);
                    if (isNgMonth) return false;
                }

                if (!participant.preferredLocations || participant.preferredLocations.length === 0) return false;
                if (!participant.preferredLocations.includes(slot.location)) return false;

                if (participant.preferredDates && participant.preferredDates.length > 0) {
                    return participant.preferredDates.some(prefDate => {
                        return isDateMatch(slot.date, prefDate);
                    });
                }

                const monthMatch = participant.preferredMonths.length === 0 ||
                    participant.preferredMonths.includes(slot.month);
                const dayMatch = participant.preferredDays.length === 0 ||
                    participant.preferredDays.some(d => slot.dayOfWeek.includes(d));

                return monthMatch && dayMatch;
            });

            const sortedSlots = availableSlots.sort((a, b) => a.rowIndex - b.rowIndex);
            let assigned = 0;

            for (const slot of sortedSlots) {
                if (assigned >= needAssignments) break;

                // 【v3.0修正】ループ内でNG期間を再チェック（同一ループ内での連続割り当て防止）
                if (slot.assignedTo) continue; // 既に埋まっている場合スキップ
                // ループ開始時点の候補抽出後に同日枠が残っている可能性があるため、再チェックする
                if (participant.assignedDates && participant.assignedDates.includes(slot.date)) continue;
                if (isWithinNgPeriod(participant, slot.date)) continue; // NG期間内ならスキップ

                slot.assignedTo = participant.displayName;
                participant.currentAssignments++;
                participant.assignedDates.push(slot.date);
                assignedCount++;
                additionalCount++;
                assigned++;
            }

            if (assigned > 0 && additionalCount <= 10) {
                console.log(`  ${participant.displayName}: ${assigned}件追加割り当て（合計${participant.currentAssignments}/${participant.maxAssignments}回）`);
            }
        });

        console.log(`第2パス完了: ${additionalCount}件の追加割り当て`);

        // ===========================================
        // 第3パス: 柔軟な条件で未割り当てを埋める
        // ===========================================
        const stillUnassignedSlots = slots.filter(s => !s.assignedTo);

        if (stillUnassignedSlots.length > 0) {
            console.log(`\n=== 第3パス: 残り${stillUnassignedSlots.length}スロットに柔軟な条件で割り当て試行 ===`);

            const flexibleParticipants = sortedParticipants.filter(p => {
                return p.currentAssignments < p.maxAssignments || p.canSupportAdditional;
            });

            console.log(`柔軟な割り当て候補者: ${flexibleParticipants.length}人`);

            let thirdPassCount = 0;

            const resortedSlots = [...stillUnassignedSlots].sort((a, b) => {
                const aCandidates = findCandidatesFlexible(a, flexibleParticipants, true);
                const bCandidates = findCandidatesFlexible(b, flexibleParticipants, true);
                return aCandidates.length - bCandidates.length;
            });

            resortedSlots.forEach(slot => {
                // 候補者を再フィルタリング（希望回数上限をリアルタイムでチェック）
                const candidates = findCandidatesFlexible(slot, flexibleParticipants, false).filter(p => {
                    // 希望回数上限を超えていないか確認
                    return p.currentAssignments < p.maxAssignments || p.canSupportAdditional;
                });

                if (candidates.length > 0) {
                    const selected = selectBestCandidate(candidates, slot);
                    slot.assignedTo = selected.displayName;
                    selected.currentAssignments++;
                    selected.assignedDates.push(slot.date);
                    assignedCount++;
                    thirdPassCount++;
                }
            });

            console.log(`第3パス完了: ${thirdPassCount}件の追加割り当て`);
        }



        const remainingUnassigned = slots.filter(s => !s.assignedTo);

        // 最終統計
        const zeroAssignments = sortedParticipants.filter(p => p.currentAssignments === 0);

        console.log(`\n=== 最終結果（全員最低1回保証モード（条件厳守）） ===`);
        console.log(`総割り当て数: ${assignedCount}件`);
        console.log(`未割り当て数: ${remainingUnassigned.length}件`);
        console.log(`参加者総数: ${sortedParticipants.length}人`);
        console.log(`1回以上割り当て: ${sortedParticipants.length - zeroAssignments.length}人`);
        console.log(`0回の人: ${zeroAssignments.length}人`);

        if (zeroAssignments.length > 0) {
            console.warn('⚠️ 以下の参加者は希望条件に合うスロットがなく、0回になりました:');
            zeroAssignments.forEach(p => {
                console.warn(`  - ${p.displayName} (入れる日数=${p.availableDaysCount}日)`);
            });
            console.warn('→ 希望条件が厳しすぎるか、該当するスロットが既に埋まっています');
        }

        return {
            slots: slots,
            participants: sortedParticipants,
            assignedCount: assignedCount,
            unassignedSlots: remainingUnassigned,
            totalSlots: slots.length,
            preAssignedCount: preAssignedCount
        };
    }

    // ===========================================
    // 日付順（単純）割り当てロジック
    // ===========================================
    function assignShiftsDateOrder(participants, slots) {
        let assignedCount = 0;
        const unassignedSlots = [];

        if (participants.length === 0) return assignShifts(participants, slots); // Fallback if 0

        initializePreAssignmentState(participants);
        const preAssignedCount = applyPreAssignedCounts(slots, participants, '日付順ロジック');

        console.log(`[日付順ロジック] 既存の割り当て数: ${preAssignedCount}件`);

        const shuffledParticipants = [...participants].sort(() => Math.random() - 0.5);
        shuffledParticipants.sort((a, b) => a.maxAssignments - b.maxAssignments);

        // ソートせずに元の順序（通常は日付順）で処理
        slots.forEach(slot => {
            if (slot.assignedTo) return; // 既に埋まっている（プレ割り当て等）場合はスキップ

            const candidates = findCandidates(slot, shuffledParticipants);

            if (candidates.length > 0) {
                const selected = selectBestCandidate(candidates, slot);
                slot.assignedTo = selected.displayName;
                selected.currentAssignments++;
                selected.assignedDates.push(slot.date);
                assignedCount++;
            } else {
                unassignedSlots.push(slot);
            }
        });

        const remainingUnassigned = slots.filter(s => !s.assignedTo);

        return {
            slots: slots,
            participants: shuffledParticipants,
            assignedCount: assignedCount,
            unassignedSlots: remainingUnassigned,
            totalSlots: slots.length,
            preAssignedCount: preAssignedCount
        };
    }

    function findCandidates(slot, participants, isDryRun = false) {
        return participants.filter(p => {
            // 割り当て上限チェック（ドライラン時は無視）
            if (!isDryRun && p.currentAssignments >= p.maxAssignments) return false;

            // 同一日に既に割り当てられていないかチェック（ドライラン時は無視：潜在的な候補者数を知りたいため）
            // ただし、もし「特定の日しかダメ」な人が既に他で埋まっている場合は除外すべきだが、
            // ここではシンプルに「条件に合う人」の総数を需要スコアとする
            if (!isDryRun && p.assignedDates && p.assignedDates.includes(slot.date)) return false;

            // 【v3.0新機能】前後NG期間チェック（ドライラン時は無視）
            // 既に割り当てられた日付から7日以内の場合は除外
            if (!isDryRun && isWithinNgPeriod(p, slot.date)) return false;

            // NG日チェック
            if (p.ngDates && p.ngDates.length > 0) {
                const isNgDate = p.ngDates.some(ngDate => isDateMatch(slot.date, ngDate));
                if (isNgDate) return false; // NG日は絶対に除外
            }
            // NG曜日チェック
            if (p.ngDays && p.ngDays.length > 0) {
                const isNgDay = p.ngDays.some(d => slot.dayOfWeek.includes(d));
                if (isNgDay) return false; // NG曜日は絶対に除外
            }
            // NG月チェック
            if (p.ngMonths && p.ngMonths.length > 0) {
                const isNgMonth = p.ngMonths.includes(slot.month);
                if (isNgMonth) return false; // NG月は絶対に除外
            }

            // 希望場所のマッチング（完全一致のみ）
            // 定員表記を除いた基本名で比較（baseNameがあればそれを使用）
            if (!p.preferredLocations || p.preferredLocations.length === 0) return false;
            const slotBaseName = slot.baseName || slot.location;
            const locationMatch = p.preferredLocations.includes(slotBaseName);
            if (!locationMatch) return false;

            // 特定の日付が指定されている場合、その日付のみ許可
            // ※「特定日のみ希望」という運用なので、ここに記述があれば月・曜日の条件は無視して
            //   その日付かどうだけで判定する
            if (p.preferredDates && p.preferredDates.length > 0) {
                // スロットの日付が希望日に含まれているかチェック
                const dateMatch = p.preferredDates.some(prefDate => {
                    return isDateMatch(slot.date, prefDate);
                });
                return dateMatch;
            }

            // 希望月のマッチング（特定日が指定されていない場合のみ適用）
            const monthMatch = p.preferredMonths.length === 0 ||
                p.preferredMonths.includes(slot.month);

            // 希望曜日のマッチング（特定日が指定されていない場合のみ適用）
            const dayMatch = p.preferredDays.length === 0 ||
                p.preferredDays.some(d => slot.dayOfWeek.includes(d));

            return monthMatch && dayMatch;
        });
    }

    function selectBestCandidate(candidates, slot) {
        const scored = candidates.map(c => {
            let score = 0;

            // 【最優先】割り当て回数が少ない人を優先（平準化）
            // currentAssignmentsが少ないほど高スコア
            // 0回の人: 100点, 1回の人: 80点, 2回: 60点...
            score += (100 - c.currentAssignments * 20);

            // 【重要】希望回数に対する充足率を考慮
            // まだ希望回数に達していない人を優先する
            const fulfillmentRate = c.currentAssignments / c.maxAssignments;
            if (fulfillmentRate < 0.5) {
                score += 30; // まだ半分以下しか割り当てられていない
            } else if (fulfillmentRate < 0.8) {
                score += 15; // 80%未満
            }

            // 【戦略的調整】選択肢が少ない（＝融通が利かない）人を優先
            // 例：この場所しか希望していない人を、他も希望している人より優先
            if (c.preferredLocations && c.preferredLocations.length > 0) {
                // 1箇所のみ: +25, 2箇所: +12, 3箇所: +5
                const flexibilityBonus = [25, 12, 5][c.preferredLocations.length - 1] || 0;
                score += flexibilityBonus;
            }

            // 【日時制約】特定日や曜日・月を指定している人を優先
            // これらの条件がある人は他のスロットで割り当てにくいため優先
            let constraintCount = 0;
            if (c.preferredDates && c.preferredDates.length > 0) {
                // ★★★ 最重要: このスロットが特定参加可能日として指定されている場合 ★★★
                // 「この日にしか来られない」という人を最優先で割り当てる
                const isExactDateMatch = c.preferredDates.some(prefDate => isDateMatch(slot.date, prefDate));
                if (isExactDateMatch) {
                    score += 200; // 特定参加可能日に完全一致 → 最優先
                }
                constraintCount += 3; // 特定日指定は最も制約が強い
            }
            if (c.preferredMonths && c.preferredMonths.length > 0 && c.preferredMonths.length < 6) {
                constraintCount += 2; // 月指定（全月の半分未満）
            }
            if (c.preferredDays && c.preferredDays.length > 0 && c.preferredDays.length < 3) {
                constraintCount += 1; // 曜日指定（平日の半分未満）
            }
            score += constraintCount * 8;

            // 【希望順位】場所の優先度（第1希望 > 第2希望 > 第3希望）
            const locIdx = c.preferredLocations.findIndex(loc => {
                if (!loc) return false;
                return loc === slot.location;
            });

            if (locIdx === 0) score += 10;      // 第1希望
            else if (locIdx === 1) score += 6;  // 第2希望
            else if (locIdx === 2) score += 3;  // 第3希望

            return { candidate: c, score: score };
        });

        // スコア降順でソート、同点の場合は参加可能回数が少ない順、さらに高学年優先、最後にランダム
        scored.sort((a, b) => {
            const scoreDiff = b.score - a.score;
            if (scoreDiff !== 0) return scoreDiff;

            // 同点の場合は参加可能回数が少ない順（負担が軽い人を優先）
            if (a.candidate.maxAssignments !== b.candidate.maxAssignments) {
                return a.candidate.maxAssignments - b.candidate.maxAssignments; // 昇順（少ない順）
            }

            // 参加可能回数も同じ場合は学年で比較（高学年優先）
            const gradeA = parseInt(a.candidate.grade.match(/(\d+)/)?.[1] || '0');
            const gradeB = parseInt(b.candidate.grade.match(/(\d+)/)?.[1] || '0');
            if (gradeA !== gradeB) {
                return gradeB - gradeA; // 降順（高学年優先）
            }

            // 学年も同じ場合はランダム
            return Math.random() - 0.5;
        });

        return scored[0].candidate;
    }


    // ========================================================================
    // Results Display
    // ========================================================================
    function displayResults(result) {
        const resultSummary = document.getElementById('resultSummary');
        const resultTable = document.getElementById('resultTable');
        const unassignedList = document.getElementById('unassignedList');

        // 統計情報の計算
        const stats = {
            totalParticipants: result.participants.length,
            totalAssigned: result.slots.filter(s => s.assignedTo).length, // プレ割り当て含めた総数
            totalSlots: result.slots.length,
            unassignedSlots: result.unassignedSlots.length,
            requestedTotal: 0
        };

        // 希望回数の合計を計算 (無制限の人は除く)
        stats.requestedTotal = result.participants.reduce((sum, p) => {
            return p.maxAssignments < 999 ? sum + p.maxAssignments : sum;
        }, 0);

        const filledSlots = stats.totalSlots - stats.unassignedSlots;
        const coverageRate = Math.round((filledSlots / stats.totalSlots) * 100);

        // 結果サマリーの表示
        // Row 1: 供給リソース情報
        const summaryHtml = `
        <div style="display:flex; flex-wrap:wrap; gap:16px; width:100%; margin-bottom:16px; justify-content: center;">
            <div class="summary-card">
                <div class="summary-number">${stats.totalParticipants}</div>
                <div class="summary-label">
                    有効回答者数<br>
                    <span style="font-size:0.7em; color:#888;">(免除者など除外)</span>
                </div>
            </div>
            <div class="summary-card">
                <div class="summary-number">${stats.requestedTotal}</div>
                <div class="summary-label">
                    参加可能回数合計<br>
                    <span style="font-size:0.7em; color:#888;">(数値指定のみ)</span>
                </div>
            </div>
        </div>

        <!-- Row 2: 枠数計算 (総枠 - 実割当て = 不足) -->
        <div style="display:flex; flex-wrap:wrap; gap:16px; width:100%; align-items: center; justify-content: center; background: #f9fafb; padding: 16px; border-radius: 8px;">
            
            <div class="summary-card" style="border-left: 4px solid #3b82f6;">
                <div class="summary-number">${stats.totalSlots}</div>
                <div class="summary-label">
                    総枠数 (必要シフト)<br>
                    <span style="font-size:0.7em; color:#888;">(場所数 × 日数)</span>
                </div>
            </div>

            <div style="font-weight:bold; color:#666; font-size:1.5rem;">－</div>

            <div class="summary-card ${coverageRate < 100 ? 'warning' : ''}">
                <div class="summary-number">${stats.totalAssigned}</div>
                <div class="summary-label">
                    実割当て合計<br>
                    <span style="font-size:0.85em; color:${coverageRate < 100 ? '#d97706' : '#059669'}; font-weight:bold;">(充足率 ${coverageRate}%)</span>
                </div>
            </div>

            <div style="font-weight:bold; color:#666; font-size:1.5rem;">＝</div>

            <div class="summary-card ${stats.unassignedSlots > 0 ? 'danger' : ''}" style="border-left: 4px solid #ef4444;">
                <div class="summary-number">${stats.unassignedSlots}</div>
                <div class="summary-label">
                    未割当て (不足枠)<br>
                    <span style="font-size:0.7em; color:#888;">残り必要な人手</span>
                </div>
            </div>
        </div>
        `;
        resultSummary.innerHTML = summaryHtml;

        // 結果テーブル
        let html = '<thead><tr><th>日付</th><th>曜日</th>';
        state.locations.forEach(loc => {
            html += `<th>${loc.name}</th>`;
        });
        html += '</tr>';

        // 追加: 集計行 (充足/総枠, 未割当て率)
        const locationStats = state.locations.map(loc => {
            const total = result.slots.filter(s => s.location === loc.name).length;
            const unassigned = result.slots.filter(s => s.location === loc.name && !s.assignedTo).length;
            const filled = total - unassigned;
            const rate = total > 0 ? ((unassigned / total) * 100).toFixed(1) : '0.0';
            return { total, unassigned, filled, rate };
        });

        // 1行目: 充足 / 総枠
        html += '<tr style="background-color: #f8fafc;">';
        html += '<th colspan="2" style="text-align: right; padding: 8px; font-weight: normal; font-size: 0.9em; color: #64748b;">充足 / 総枠</th>';
        locationStats.forEach(stat => {
            const isFull = stat.unassigned === 0;
            const color = isFull ? '#059669' : '#d97706'; // 満杯なら緑、未達ならオレンジ
            html += `<th style="text-align: center; font-weight: normal; font-size: 0.9em; color: ${color};">${stat.filled} / ${stat.total}</th>`;
        });
        html += '</tr>';

        // 2行目: 未割当て率
        html += '<tr style="background-color: #f8fafc; border-bottom: 2px solid #e2e8f0;">';
        html += '<th colspan="2" style="text-align: right; padding: 8px; font-weight: normal; font-size: 0.9em; color: #64748b;">未割当て率</th>';
        locationStats.forEach(stat => {
            const rateVal = parseFloat(stat.rate);
            const color = rateVal > 20 ? '#dc2626' : (rateVal > 0 ? '#d97706' : '#059669');
            html += `<th style="text-align: center; font-weight: bold; font-size: 0.9em; color: ${color};">${stat.rate}%</th>`;
        });
        html += '</tr>';

        html += '</thead><tbody>';

        // 複数人体制対応: 同じ日付・場所の複数スロットを集約
        const slotsByDate = {};
        result.slots.forEach(slot => {
            if (!slotsByDate[slot.date]) {
                slotsByDate[slot.date] = {
                    dayOfWeek: slot.dayOfWeek,
                    locations: {}  // 各地点に対して名前の配列を保持
                };
            }
            // 地点ごとに名前を配列で保持
            if (!slotsByDate[slot.date].locations[slot.location]) {
                slotsByDate[slot.date].locations[slot.location] = [];
            }
            if (slot.assignedTo) {
                slotsByDate[slot.date].locations[slot.location].push(slot.assignedTo);
            }
        });

        // 日付でソート
        const sortedDates = Object.keys(slotsByDate).sort((a, b) => {
            const dateA = new Date(a.replace(/\//g, '-'));
            const dateB = new Date(b.replace(/\//g, '-'));
            return dateA - dateB;
        });

        sortedDates.forEach(date => {
            const data = slotsByDate[date];
            const isHoliday = ['土', '日', '祝'].some(d => data.dayOfWeek.includes(d));
            const rowStyle = isHoliday ? 'style="background: #f5f5f5; color: #999;"' : '';

            html += `<tr ${rowStyle}>`;
            html += `<td>${date}</td>`;
            html += `<td>${data.dayOfWeek}</td>`;
            state.locations.forEach(loc => {
                const names = data.locations[loc.name] || [];
                const capacity = loc.capacity || 1;
                const assignedCount = names.length;

                // 定員と割当て数を比較してスタイルを決定
                let cellClass = '';
                if (assignedCount === 0 && !isHoliday) {
                    cellClass = 'cell-empty';
                } else if (assignedCount < capacity && !isHoliday) {
                    cellClass = 'cell-partial'; // 部分的に埋まっている
                } else if (assignedCount > 0) {
                    cellClass = 'cell-assigned';
                }

                html += `<td class="${cellClass}">${names.join('、')}</td>`;
            });
            html += '</tr>';
        });

        html += '</tbody>';
        resultTable.innerHTML = html;

        // 未割り当てリストは集計行へ移動したため非表示
        unassignedSection.classList.add('hidden');
        unassignedList.innerHTML = '';

        // 割当て集計表示
        displayAssignmentSummary(result);

        // 特定参加可能日集中度表示
        displayPreferredDatesDensity(result);

        // 詳細分析表示
        displayDetailAnalysis(result);
    }

    // ========================================================================
    // 特定参加可能日集中度テーブル
    // ========================================================================
    function displayPreferredDatesDensity(result) {
        const densityTable = document.getElementById('preferredDatesDensityTable');
        if (!densityTable) return;

        // 各日・各場所への特定参加可能日の集計
        // キー: "日付_場所" → 値: { count: 人数, names: [名前リスト] }
        const densityMap = {};

        // 参加者の特定参加可能日を集計
        result.participants.forEach(p => {
            if (p.preferredDates && p.preferredDates.length > 0) {
                p.preferredDates.forEach(prefDate => {
                    // この日に対応する場所（希望場所）をすべてカウント
                    p.preferredLocations.forEach(location => {
                        // シフト表に存在する日付かチェック
                        const matchingSlot = result.slots.find(slot =>
                            isDateMatch(slot.date, prefDate) && slot.location === location
                        );
                        if (matchingSlot) {
                            const key = `${matchingSlot.date}_${location}`;
                            if (!densityMap[key]) {
                                densityMap[key] = { count: 0, names: [], date: matchingSlot.date, location: location };
                            }
                            densityMap[key].count++;
                            densityMap[key].names.push(p.displayName);
                        }
                    });
                });
            }
        });

        // 場所リスト
        const locations = state.locations.map(loc => loc.name);

        // 日付リスト（土日祝を除く）
        const dates = state.dates.filter(d => !d.isHoliday);

        // ヘッダー行
        let html = '<thead><tr><th>日付</th><th>曜日</th>';
        locations.forEach(loc => {
            html += `<th style="text-align:center;">${loc}</th>`;
        });
        html += '</tr></thead><tbody>';

        // 各日付の行
        dates.forEach(dateInfo => {
            html += `<tr><td>${dateInfo.date}</td><td>${dateInfo.dayOfWeek}</td>`;

            locations.forEach(location => {
                const key = `${dateInfo.date}_${location}`;
                const data = densityMap[key];

                if (data && data.count > 0) {
                    // 集中度に応じた背景色
                    let bgColor = '';
                    let fontWeight = '';
                    if (data.count >= 3) {
                        bgColor = 'background: #ffcdd2;'; // 赤（高集中）
                        fontWeight = 'font-weight: bold;';
                    } else if (data.count >= 2) {
                        bgColor = 'background: #ffe0b2;'; // オレンジ（中集中）
                    } else {
                        bgColor = 'background: #e8f5e9;'; // 緑（低集中）
                    }

                    // ツールチップに名前リストを表示
                    const tooltip = data.names.join(', ');
                    html += `<td style="${bgColor}${fontWeight} text-align:center; cursor:help;" title="${tooltip}">${data.count}人</td>`;
                } else {
                    html += '<td style="text-align:center; color:#ccc;">-</td>';
                }
            });

            html += '</tr>';
        });

        html += '</tbody>';
        densityTable.innerHTML = html;
    }

    function displayAssignmentSummary(result) {
        const summaryTable = document.getElementById('assignmentSummaryTable');

        // 参加者のmaxAssignmentsをマッピング
        const participantMaxMap = {};
        result.participants.forEach(p => {
            participantMaxMap[p.displayName] = p.maxAssignments;
        });

        // 全参加者をベースにして割当て回数を集計（0回の人も含む）
        const assignmentCounts = {};

        // まず、全参加者を登録（割当て0回でも表示）
        result.participants.forEach(p => {
            assignmentCounts[p.displayName] = {
                name: p.displayName,
                existingCount: 0,
                newCount: 0,
                dates: [],
                maxAssignments: p.maxAssignments,
                canSupportAdditional: p.canSupportAdditional || false
            };
        });

        // 次に、実際の割当てをカウント
        result.slots.forEach(slot => {
            if (slot.assignedTo && assignmentCounts[slot.assignedTo]) {
                if (slot.isPreAssigned) {
                    assignmentCounts[slot.assignedTo].existingCount++;
                } else {
                    assignmentCounts[slot.assignedTo].newCount++;
                }
                assignmentCounts[slot.assignedTo].dates.push(`${slot.date}(${slot.location})`);
            }
        });

        const totalExisting = Object.values(assignmentCounts).reduce((sum, p) => sum + p.existingCount, 0);
        const totalNew = Object.values(assignmentCounts).reduce((sum, p) => sum + p.newCount, 0);

        console.log('=== 割当て集計（全体） ===');
        console.log('集計対象者数:', Object.keys(assignmentCounts).length);
        console.log('既存割当て総数:', totalExisting);
        console.log('新規割当て総数:', totalNew);
        console.log('合計割当て総数:', totalExisting + totalNew);

        // 割当て回数でソート（合計の昇順）
        const sortedPeople = Object.values(assignmentCounts).sort((a, b) => {
            const totalA = a.existingCount + a.newCount;
            const totalB = b.existingCount + b.newCount;
            return totalA - totalB;
        });

        // テーブル生成
        let html = `
            <thead>
                <tr>
                    <th>No.</th>
                    <th style="min-width: 200px;">名前</th>
                    <th>参加可能回数</th>
                    <th>既存</th>
                    <th>新規</th>
                    <th>合計</th>
                    <th>割当て日</th>
                </tr>
            </thead>
            <tbody>
        `;

        sortedPeople.forEach((person, idx) => {
            const totalCount = person.existingCount + person.newCount;
            const isOverLimit = totalCount > person.maxAssignments && person.maxAssignments !== '-';

            // 追加対応可で参加可能回数を超えているか判定
            const isAdditionalSupport = isOverLimit && person.canSupportAdditional;
            const isRealOverLimit = isOverLimit && !person.canSupportAdditional;

            // 行スタイル: 追加対応可は薄い緑、問題ある超過は赤
            let rowStyle = '';
            if (isRealOverLimit) {
                rowStyle = 'style="background: #ffebee; color: #c62828;"';
            } else if (isAdditionalSupport) {
                rowStyle = 'style="background: #e8f5e9;"';
            }
            const maxDisplay = person.maxAssignments >= 999 ? '無制限' : person.maxAssignments;

            // 既存のみ、新規のみ、両方あり、で色分け
            let nameStyle = '';
            if (person.existingCount > 0 && person.newCount > 0) {
                nameStyle = 'style="background: #e3f2fd; font-weight: bold;"'; // 両方（青）
            } else if (person.newCount > 0) {
                nameStyle = 'style="background: #f1f8e9; font-weight: bold;"'; // 新規のみ（緑）
            } else {
                nameStyle = 'style="background: #fafafa;"'; // 既存のみ（灰色）
            }

            html += `
                <tr ${rowStyle}>
                    <td>${idx + 1}</td>
                    <td ${nameStyle}>${person.name}</td>
                    <td style="text-align: center;">${maxDisplay}</td>
                    <td style="text-align: center; color: #666;">${person.existingCount}</td>
                    <td style="text-align: center; font-weight: bold; color: #10b981;">${person.newCount}</td>
                    <td style="text-align: center; font-weight: bold; ${isRealOverLimit ? 'color: #c62828;' : (isAdditionalSupport ? 'color: #2e7d32;' : 'color: #1976d2;')}">${totalCount}${isRealOverLimit ? ' ⚠️' : (isAdditionalSupport ? ' 📌追加' : '')}</td>
                    <td style="font-size: 0.85em;">${person.dates.join(', ')}</td>
                </tr>
            `;
        });

        html += '</tbody>';
        summaryTable.innerHTML = html;
    }

    // ========================================================================
    // 詳細分析テーブル（アンケート回答と割り当て結果の照合）
    // ========================================================================
    function displayDetailAnalysis(result) {
        const analysisTable = document.getElementById('detailAnalysisTable');
        if (!analysisTable) return;

        // 各参加者の割り当て情報を集計
        const assignmentsByPerson = {};
        result.slots.forEach(slot => {
            if (slot.assignedTo) {
                if (!assignmentsByPerson[slot.assignedTo]) {
                    assignmentsByPerson[slot.assignedTo] = [];
                }
                assignmentsByPerson[slot.assignedTo].push({
                    date: slot.date,
                    location: slot.location,
                    baseName: slot.baseName || slot.location,  // マッチング用基本名
                    month: slot.month,
                    dayOfWeek: slot.dayOfWeek
                });
            }
        });

        // ヘッダー部分
        let html = '<thead><tr>';
        html += '<th>No.</th>';
        html += '<th style="min-width:200px;">名前</th>';
        html += '<th style="background:#e8f5e9;">希望場所</th>';
        html += '<th style="background:#e8f5e9;">希望月</th>';
        html += '<th style="background:#e8f5e9;">希望曜日</th>';
        html += '<th style="background:#e8f5e9;">特定参加可能日</th>';
        html += '<th style="background:#ffebee;">特定NG日</th>';
        html += '<th style="background:#e8f5e9;">参加可能回数</th>';
        html += '<th style="background:#e3f2fd;">割当て回数</th>';
        html += '<th style="background:#e3f2fd;min-width:300px;">割当て詳細</th>';
        html += '<th style="background:#e3f2fd;">合致</th>';
        html += '<th>判定</th>';
        html += '</tr></thead><tbody>';

        const sortedParticipants = [...result.participants].sort((a, b) => {
            const aCount = assignmentsByPerson[a.displayName]?.length || 0;
            const bCount = assignmentsByPerson[b.displayName]?.length || 0;
            return aCount - bCount;
        });

        sortedParticipants.forEach((p, idx) => {
            const assignments = assignmentsByPerson[p.displayName] || [];
            const assignmentCount = assignments.length;

            let allMatch = true;
            const assignmentDetails = assignments.map(a => {
                let status = '✅';
                // 定員表記を除いた基本名でマッチング
                const locationBaseName = a.baseName || a.location;
                const locationMatch = p.preferredLocations && p.preferredLocations.includes(locationBaseName);

                // 特定日が指定されている場合
                let dateMatch = true;
                if (p.preferredDates && p.preferredDates.length > 0) {
                    dateMatch = p.preferredDates.some(prefDate => {
                        const assignedParts = a.date.split('/');
                        const preferredParts = prefDate.split('/');
                        if (preferredParts.length === 3) {
                            return a.date === prefDate;
                        } else if (preferredParts.length === 2) {
                            return assignedParts.length >= 2 && assignedParts[assignedParts.length - 2] === preferredParts[0] && assignedParts[assignedParts.length - 1] === preferredParts[1];
                        }
                        return false;
                    });
                }

                // 月・曜日のチェック（特定日が指定されていない場合のみ）
                let monthMatch = true;
                let dayMatch = true;
                if (!p.preferredDates || p.preferredDates.length === 0) {
                    monthMatch = p.preferredMonths.length === 0 || p.preferredMonths.includes(a.month);
                    dayMatch = p.preferredDays.length === 0 || p.preferredDays.some(d => a.dayOfWeek.includes(d));
                }

                if (!locationMatch) { status = '❌場'; allMatch = false; }
                else if (!dateMatch && p.preferredDates && p.preferredDates.length > 0) { status = '❌日'; allMatch = false; }
                else if (!monthMatch && (!p.preferredDates || p.preferredDates.length === 0)) { status = '❌月'; allMatch = false; }
                else if (!dayMatch && (!p.preferredDates || p.preferredDates.length === 0)) { status = '❌曜'; allMatch = false; }

                const bg = status === '✅' ? '#e8f5e9' : '#ffebee';
                return '<span style="display:inline-block;margin:1px;padding:2px 4px;background:' + bg + ';border-radius:3px;font-size:0.8em;">' + a.date + '(' + a.location + ')' + status + '</span>';
            }).join('');

            let rowBg = '';
            let rowJudgment = '<span style="color:#2e7d32;">✅OK</span>';
            if (assignmentCount === 0) {
                rowBg = 'background:#fff3e0;';
                rowJudgment = '<span style="color:#e65100;">⚠️0回</span>';
            } else if (!allMatch) {
                rowBg = 'background:#ffebee;';
                rowJudgment = '<span style="color:#c62828;">❌問題</span>';
            }

            const locationDisplay = p.preferredLocations && p.preferredLocations.length > 0 ? p.preferredLocations.join(', ') : '-';
            const monthDisplay = p.preferredMonths && p.preferredMonths.length > 0 ? p.preferredMonths.map(m => m + '月').join(', ') : '-';
            const dayDisplay = p.preferredDays && p.preferredDays.length > 0 ? p.preferredDays.join(', ') : '-';
            const dateDisplay = p.preferredDates && p.preferredDates.length > 0 ? p.preferredDates.join(', ') : '-';
            const ngDateDisplay = p.ngDates && p.ngDates.length > 0 ? p.ngDates.join(', ') : '-';
            const maxDisplay = p.maxAssignments >= 999 ? '無制限' : p.maxAssignments;
            const matchCount = assignments.filter(a => {
                const locationBaseName = a.baseName || a.location;
                return p.preferredLocations && p.preferredLocations.includes(locationBaseName);
            }).length;
            const matchRate = assignmentCount > 0 ? Math.round(matchCount / assignmentCount * 100) + '%' : '-';
            const nameDisplay = p.displayName + (p.canSupportAdditional ? ' 📌' : '');

            html += '<tr style="' + rowBg + '">';
            html += '<td>' + (idx + 1) + '</td>';
            html += '<td>' + nameDisplay + '</td>';
            html += '<td>' + locationDisplay + '</td>';
            html += '<td>' + monthDisplay + '</td>';
            html += '<td>' + dayDisplay + '</td>';
            html += '<td>' + dateDisplay + '</td>';
            html += '<td>' + ngDateDisplay + '</td>';
            html += '<td>' + maxDisplay + '</td>';
            html += '<td style="font-weight:bold;color:' + (assignmentCount === 0 ? '#c62828' : '#1565c0') + ';">' + assignmentCount + '</td>';
            html += '<td>' + (assignmentDetails || '-') + '</td>';
            html += '<td>' + matchRate + '</td>';
            html += '<td>' + rowJudgment + '</td>';
            html += '</tr>';
        });

        html += '</tbody>';
        analysisTable.innerHTML = html;
    }

    // ========================================================================
    // Excel Export
    // ========================================================================
    exportBtn.addEventListener('click', exportToExcel);

    function exportToExcel() {
        if (!state.assignmentResult) {
            alert('先にシフトの割当てを実行してください');
            return;
        }

        const wb = XLSX.utils.book_new();
        const shiftExportData = [];

        // ヘッダー行
        const headerRow = ['日付', '曜日'];
        state.locations.forEach(loc => headerRow.push(loc.name));
        shiftExportData.push(headerRow);

        // データ行（複数人体制対応: 同じ場所の複数スロットを集約）
        const slotsByDate = {};
        state.assignmentResult.slots.forEach(slot => {
            if (!slotsByDate[slot.date]) {
                slotsByDate[slot.date] = {
                    dayOfWeek: slot.dayOfWeek,
                    locations: {}  // 各地点に対して名前の配列を保持
                };
            }
            // 地点ごとに名前を配列で保持
            if (!slotsByDate[slot.date].locations[slot.location]) {
                slotsByDate[slot.date].locations[slot.location] = [];
            }
            if (slot.assignedTo) {
                slotsByDate[slot.date].locations[slot.location].push(slot.assignedTo);
            }
        });

        const sortedDates = Object.keys(slotsByDate).sort((a, b) => {
            const dateA = new Date(a.replace(/\//g, '-'));
            const dateB = new Date(b.replace(/\//g, '-'));
            return dateA - dateB;
        });

        sortedDates.forEach(date => {
            const data = slotsByDate[date];
            const row = [date, data.dayOfWeek];
            state.locations.forEach(loc => {
                // 複数名をカンマ区切りで結合
                const names = data.locations[loc.name] || [];
                row.push(names.join('、'));
            });
            shiftExportData.push(row);
        });

        const ws = XLSX.utils.aoa_to_sheet(shiftExportData);

        // 列幅設定
        const colWidths = [{ wch: 12 }, { wch: 6 }];
        state.locations.forEach(() => colWidths.push({ wch: 18 }));
        ws['!cols'] = colWidths;

        XLSX.utils.book_append_sheet(wb, ws, 'シフト表');

        // === 割当て集計シート ===
        const summaryData = [];

        // ヘッダー行
        summaryData.push(['No.', '名前', '参加可能回数', '割当て回数', '割当て日']);

        // 参加者のmaxAssignmentsをマッピング
        const participantMaxMap = {};
        if (state.assignmentResult.participants) {
            state.assignmentResult.participants.forEach(p => {
                participantMaxMap[p.displayName] = p.maxAssignments;
            });
        }

        // 全参加者をベースにして割当て回数を集計（0回の人も含む）
        const assignmentCounts = {};

        // まず、全参加者を登録（割当て0回でも表示）
        if (state.assignmentResult.participants) {
            state.assignmentResult.participants.forEach(p => {
                const maxDisplay = p.maxAssignments >= 999 ? '無制限' : p.maxAssignments;
                assignmentCounts[p.displayName] = {
                    name: p.displayName,
                    max: maxDisplay,
                    count: 0,
                    dates: []
                };
            });
        }

        // 次に、実際の割当てをカウント
        state.assignmentResult.slots.forEach(slot => {
            if (slot.assignedTo && assignmentCounts[slot.assignedTo]) {
                assignmentCounts[slot.assignedTo].count++;
                assignmentCounts[slot.assignedTo].dates.push(`${slot.date} (${slot.location})`);
            }
        });

        // 割当て回数でソート（昇順）
        const sortedPeople = Object.values(assignmentCounts).sort((a, b) => a.count - b.count);

        // データ行追加
        sortedPeople.forEach((person, idx) => {
            summaryData.push([
                idx + 1,
                person.name,
                person.max,     // Export max
                person.count,
                person.dates.join(', ')
            ]);
        });

        const ws2 = XLSX.utils.aoa_to_sheet(summaryData);
        ws2['!cols'] = [{ wch: 5 }, { wch: 40 }, { wch: 10 }, { wch: 12 }, { wch: 80 }];
        XLSX.utils.book_append_sheet(wb, ws2, '割当て集計');

        // === 詳細分析シート ===
        const analysisData = [];

        // ヘッダー行
        analysisData.push([
            'No.', '名前', '追加対応可',
            '希望場所', '参加可能月', '参加可能曜日', '特定参加可能日', '特定NG日', '参加可能回数',
            '割当て回数', '割当て詳細', '条件合致', '判定'
        ]);

        // 各参加者の割当て情報を集計
        const assignmentsByPersonExcel = {};
        state.assignmentResult.slots.forEach(slot => {
            if (slot.assignedTo) {
                if (!assignmentsByPersonExcel[slot.assignedTo]) {
                    assignmentsByPersonExcel[slot.assignedTo] = [];
                }
                assignmentsByPersonExcel[slot.assignedTo].push({
                    date: slot.date,
                    location: slot.location,
                    month: slot.month,
                    dayOfWeek: slot.dayOfWeek
                });
            }
        });

        // 参加者をソート
        const sortedParticipantsExcel = [...state.assignmentResult.participants].sort((a, b) => {
            const aCount = assignmentsByPersonExcel[a.displayName]?.length || 0;
            const bCount = assignmentsByPersonExcel[b.displayName]?.length || 0;
            return aCount - bCount;
        });

        sortedParticipantsExcel.forEach((p, idx) => {
            const assignments = assignmentsByPersonExcel[p.displayName] || [];
            const assignmentCount = assignments.length;

            // 条件合致チェック
            let allMatch = true;
            const assignmentDetails = assignments.map(a => {
                let status = '✅';
                const locationMatch = p.preferredLocations && p.preferredLocations.includes(a.location);

                // 特定日が指定されている場合
                let dateMatch = true;
                if (p.preferredDates && p.preferredDates.length > 0) {
                    dateMatch = p.preferredDates.some(prefDate => {
                        const assignedParts = a.date.split('/');
                        const preferredParts = prefDate.split('/');
                        if (preferredParts.length === 3) {
                            return a.date === prefDate;
                        } else if (preferredParts.length === 2) {
                            return assignedParts.length >= 2 && assignedParts[assignedParts.length - 2] === preferredParts[0] && assignedParts[assignedParts.length - 1] === preferredParts[1];
                        }
                        return false;
                    });
                }

                // 月・曜日のチェック（特定日が指定されていない場合のみ）
                let monthMatch = true;
                let dayMatch = true;
                if (!p.preferredDates || p.preferredDates.length === 0) {
                    monthMatch = p.preferredMonths.length === 0 || p.preferredMonths.includes(a.month);
                    dayMatch = p.preferredDays.length === 0 || p.preferredDays.some(d => a.dayOfWeek.includes(d));
                }

                if (!locationMatch) { status = '❌場所'; allMatch = false; }
                else if (!dateMatch && p.preferredDates && p.preferredDates.length > 0) { status = '❌日付'; allMatch = false; }
                else if (!monthMatch && (!p.preferredDates || p.preferredDates.length === 0)) { status = '❌月'; allMatch = false; }
                else if (!dayMatch && (!p.preferredDates || p.preferredDates.length === 0)) { status = '❌曜日'; allMatch = false; }

                return a.date + '(' + a.location + ')' + status;
            }).join(', ');

            // 判定
            let judgment = 'OK';
            if (assignmentCount === 0) {
                judgment = '⚠️ 0回';
            } else if (!allMatch) {
                judgment = '❌ 問題あり';
            }

            // 合致率
            const matchCount = assignments.filter(a => p.preferredLocations && p.preferredLocations.includes(a.location)).length;
            const matchRate = assignmentCount > 0 ? Math.round(matchCount / assignmentCount * 100) + '%' : '-';

            const locationDisplay = p.preferredLocations && p.preferredLocations.length > 0 ? p.preferredLocations.join(', ') : '-';
            const monthDisplay = p.preferredMonths && p.preferredMonths.length > 0 ? p.preferredMonths.map(m => m + '月').join(', ') : '-';
            const dayDisplay = p.preferredDays && p.preferredDays.length > 0 ? p.preferredDays.join(', ') : '-';
            const dateDisplay = p.preferredDates && p.preferredDates.length > 0 ? p.preferredDates.join(', ') : '-';
            const ngDateDisplay = p.ngDates && p.ngDates.length > 0 ? p.ngDates.join(', ') : '-';
            const maxDisplay = p.maxAssignments >= 999 ? '無制限' : p.maxAssignments;

            analysisData.push([
                idx + 1,
                p.displayName,
                p.canSupportAdditional ? '○' : '',
                locationDisplay,
                monthDisplay,
                dayDisplay,
                dateDisplay,
                ngDateDisplay,
                maxDisplay,
                assignmentCount,
                assignmentDetails || '-',
                matchRate,
                judgment
            ]);
        });

        const ws3 = XLSX.utils.aoa_to_sheet(analysisData);
        ws3['!cols'] = [
            { wch: 5 },   // No.
            { wch: 25 },  // 名前
            { wch: 8 },   // 追加対応可
            { wch: 30 },  // 希望場所
            { wch: 15 },  // 希望月
            { wch: 15 },  // 希望曜日
            { wch: 15 },  // 特定希望日
            { wch: 15 },  // 特定NG日
            { wch: 10 },  // 希望回数
            { wch: 10 },  // 割当回数
            { wch: 60 },  // 割り当て詳細
            { wch: 10 },  // 条件合致
            { wch: 12 }   // 判定
        ];
        XLSX.utils.book_append_sheet(wb, ws3, '詳細分析');

        const now = new Date();
        const dateStr = `${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}`;
        XLSX.writeFile(wb, `旗振りシフト表_${dateStr}.xlsx`);
    }

    // ========================================================================
    // Utilities
    // ========================================================================
    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    function escapeHtml(text) {
        return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    // ========================================================================
    // Initialize
    // ========================================================================
    updateAssignmentButtonState();
    console.log('🚩 PTA旗振りシフト割り当てツール initialized!');
});












