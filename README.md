<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>史代新批次報價單</title>
    <!-- 引入 Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- 引入 PDF.js -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
    <!-- 引入 FontAwesome -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Google Fonts -->
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&family=Roboto+Mono:wght@400;500;700&display=swap" rel="stylesheet">
    
    <script>
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        
        // Tailwind 配置
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: {
                        sans: ['"Noto Sans TC"', 'sans-serif'],
                        mono: ['"Roboto Mono"', 'monospace'],
                        arial: ['Arial', 'sans-serif'],
                    }
                }
            }
        }
    </script>

    <style>
        body { background-color: #f3f4f6; }
        .drag-active { border-color: #3b82f6 !important; background-color: #eff6ff !important; transform: scale(1.01); }
        .table-row-anim { animation: fadeIn 0.3s ease-in-out; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(5px); } to { opacity: 1; transform: translateY(0); } }
        
        /* 自定義捲軸 */
        .custom-scrollbar::-webkit-scrollbar { width: 12px; height: 12px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #f1f1f1; border-radius: 6px; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 6px; border: 3px solid #f1f1f1; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
        
        /* Tab Active State */
        .tab-active { border-bottom: 2px solid #2563eb; color: #1d4ed8; background-color: #eff6ff; }

        /* Diff Status Colors */
        .diff-match { background-color: #f0fdf4; color: #15803d; }
        .diff-ok-money { background-color: #ecfccb; color: #3f6212; }
        .diff-error { background-color: #fef2f2; color: #991b1b; }
        .diff-warn { background-color: #fff7ed; color: #9a3412; }
        .diff-info { background-color: #eff6ff; color: #1e40af; }
        .diff-fuzzy { background-color: #fffbeb; color: #b45309; }
        .diff-smart { background-color: #f3e8ff; color: #6b21a8; }
    </style>
</head>
<body class="text-gray-700 h-screen flex flex-col overflow-hidden relative">

    <!-- 頂部導航 -->
    <header class="bg-white border-b border-gray-200 shadow-sm z-20 flex-none">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
            <div class="flex items-center gap-3">
                <div class="h-10 w-10 rounded-lg shadow-sm border border-red-100 bg-red-50 flex items-center justify-center">
                     <i class="fa-solid fa-file-pdf text-2xl text-red-600"></i>
                </div>
                <div>
                    <h1 class="text-xl font-bold text-gray-800 tracking-tight">史代新批次報價單<span class="text-blue-600 text-sm font-medium px-2 py-0.5 bg-blue-50 rounded-full border border-blue-100 ml-1">Pro</span></h1>
                </div>
            </div>
            <div class="flex items-center gap-3">
                <button onclick="resetApp()" class="hidden group px-4 py-2 text-sm font-medium text-gray-500 hover:text-red-600 transition-colors" id="clearBtn">
                    <i class="fa-solid fa-trash-can mr-1 group-hover:animate-bounce"></i> 清除全部
                </button>
            </div>
        </div>
    </header>

    <!-- 主要內容區 -->
    <main class="flex-1 overflow-y-auto p-4 sm:p-6 lg:p-8 scroll-smooth" id="mainContainer">
        <div class="max-w-[95%] mx-auto space-y-6">

            <!-- 上傳區域 -->
            <div id="dropZone" class="bg-white rounded-2xl border-2 border-dashed border-gray-300 p-10 flex flex-col items-center justify-center text-center cursor-pointer transition-all duration-300 hover:border-blue-400 hover:shadow-lg group min-h-[200px]">
                <input type="file" id="fileInput" accept="application/pdf" multiple class="hidden" />
                <div id="uploadPrompt" class="space-y-4 pointer-events-none">
                    <div class="w-16 h-16 bg-blue-50 rounded-full flex items-center justify-center mx-auto mb-4 group-hover:scale-110 transition-transform duration-300">
                        <i class="fa-solid fa-cloud-arrow-up text-3xl text-blue-500"></i>
                    </div>
                    <h3 class="text-lg font-bold text-gray-700">點擊或拖放 PDF 檔案</h3>
                    <p class="text-gray-400 text-sm max-w-sm mx-auto">支援批次解析。系統將自動識別單號、計算代購費。</p>
                </div>
                <div id="loadingPrompt" class="hidden space-y-4">
                    <div class="relative w-16 h-16 mx-auto">
                        <div class="absolute inset-0 border-4 border-gray-200 rounded-full"></div>
                        <div class="absolute inset-0 border-4 border-blue-500 rounded-full border-t-transparent animate-spin"></div>
                    </div>
                    <p class="text-base font-medium text-gray-700 animate-pulse" id="loadingText">正在解析檔案...</p>
                </div>
            </div>

            <!-- 錯誤訊息 -->
            <div id="errorMsg" class="hidden transform transition-all duration-300">
                <div class="bg-red-50 border-l-4 border-red-500 p-4 rounded-r-lg shadow-sm flex items-start gap-3">
                    <i class="fa-solid fa-circle-exclamation text-red-500 mt-1"></i>
                    <div>
                        <h3 class="font-bold text-red-800">部分檔案解析失敗</h3>
                        <ul id="errorList" class="mt-2 text-sm text-red-700 list-disc list-inside space-y-1"></ul>
                    </div>
                </div>
            </div>

            <!-- 結果顯示區 -->
            <div id="resultSection" class="hidden space-y-4">
                
                <!-- Dashboard -->
                <div id="pdfDashboard" class="bg-white rounded-xl shadow-sm border border-gray-200 p-4 flex flex-col md:flex-row justify-between items-center gap-4 sticky top-0 z-10 backdrop-blur-sm bg-white/95">
                    <div class="flex items-center gap-2">
                        <span class="bg-blue-100 text-blue-700 text-xs font-bold px-2 py-1 rounded-full border border-blue-200" id="fileCountBadge">0 份</span>
                        <div class="flex items-center gap-1 text-sm text-gray-600">
                             原始: <span id="sumOriginal" class="font-bold font-arial">$0</span>
                             + 代購: <span id="sumFee" class="font-bold font-arial text-blue-600">$0</span>
                             = 總計: <span id="sumTotal" class="font-bold font-arial text-green-700">$0</span>
                        </div>
                    </div>
                    <div class="flex gap-2">
                        <button onclick="copyCurrentView()" id="copyBtn" class="bg-blue-600 hover:bg-blue-700 text-white px-4 py-1.5 rounded-lg text-sm font-medium transition-colors shadow-sm flex items-center gap-2">
                            <i class="fa-regular fa-copy"></i> 複製表格
                        </button>
                    </div>
                </div>

                <!-- Tabs -->
                <div class="bg-white rounded-t-xl border-b border-gray-200 flex overflow-x-auto">
                    <button onclick="switchTab('detail')" id="tabDetail" class="flex-1 py-3 px-4 text-sm font-medium text-gray-500 hover:text-gray-700 hover:bg-gray-50 transition-colors tab-active flex items-center justify-center gap-2 whitespace-nowrap">
                        <i class="fa-solid fa-file-invoice"></i> 報價單解析 (PDF)
                    </button>
                    <button onclick="switchTab('sheet')" id="tabSheet" class="flex-1 py-3 px-4 text-sm font-medium text-gray-500 hover:text-gray-700 hover:bg-gray-50 transition-colors flex items-center justify-center gap-2 whitespace-nowrap">
                        <i class="fa-solid fa-pen-to-square"></i> 合約及非合約 填寫區
                    </button>
                    <button onclick="switchTab('diff')" id="tabDiff" class="flex-1 py-3 px-4 text-sm font-medium text-gray-500 hover:text-gray-700 hover:bg-gray-50 transition-colors flex items-center justify-center gap-2 whitespace-nowrap border-l border-gray-100">
                        <i class="fa-solid fa-scale-balanced"></i> 差異比對分析
                    </button>
                </div>

                <!-- Content -->
                <div class="bg-white rounded-b-xl shadow-sm border border-t-0 border-gray-200 min-h-[400px]">
                    
                    <!-- Detail View -->
                    <div id="viewDetail" class="p-0">
                        <div class="overflow-x-auto">
                            <table class="w-full text-left border-collapse">
                                <thead class="bg-gray-50 text-gray-600 text-xs uppercase font-semibold tracking-wider">
                                    <tr>
                                        <th class="p-3 border-b w-12 text-center">#</th>
                                        <th class="p-3 border-b w-24">編號</th>
                                        <th class="p-3 border-b w-28">產品編號</th>
                                        <th class="p-3 border-b min-w-[200px]">產品名稱</th> 
                                        <th class="p-3 border-b text-right w-16">數量</th>
                                        <th class="p-3 border-b text-center w-16">單位</th>
                                        <th class="p-3 border-b text-right w-24">單價</th>
                                        <th class="p-3 border-b text-right w-24">原始金額</th>
                                        <th class="p-3 border-b text-right w-24 text-blue-600">代購費</th>
                                        <th class="p-3 border-b text-right w-24 text-green-700">總金額</th>
                                        <th class="p-3 border-b min-w-[100px]">備註</th> 
                                    </tr>
                                </thead>
                                <tbody id="tableBodyDetail" class="divide-y divide-gray-100 text-sm"></tbody>
                            </table>
                        </div>
                    </div>

                    <!-- Sheet View -->
                    <div id="viewSheet" class="hidden p-0 flex flex-col h-full">
                        <div class="bg-blue-50 p-3 text-sm text-blue-800 border-b border-blue-100 flex flex-wrap justify-between items-center gap-2">
                            <span><i class="fa-solid fa-paste mr-1"></i> 請將 Excel 或 Sheets 的資料複製並貼上，系統自動辨識欄位。</span>
                            <button onclick="clearSheetInput()" class="text-xs bg-white border border-blue-200 hover:bg-blue-100 px-2 py-1 rounded text-blue-700 transition">
                                清空重填
                            </button>
                        </div>
                        <div class="p-4 space-y-4">
                            <textarea id="sheetPasteInput" placeholder="在此處貼上 合約及非合約 資料..." class="w-full h-32 p-3 text-sm font-mono border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all placeholder-gray-400 bg-gray-50 hover:bg-white"></textarea>
                            <div class="overflow-x-auto border rounded-lg border-gray-200 shadow-sm bg-white">
                                <table class="w-full text-left border-collapse" id="sheetTable">
                                    <thead class="bg-gray-100 text-gray-700 text-sm font-bold border-b-2 border-gray-300">
                                        <tr>
                                            <th class="p-2 border border-gray-300 w-16 bg-gray-50">Item</th>
                                            <th class="p-2 border border-gray-300 w-32 bg-gray-50">產品編號</th>
                                            <th class="p-2 border border-gray-300 text-right w-20 bg-gray-50">訂購量</th>
                                            <th class="p-2 border border-gray-300 text-right w-24 bg-gray-50">金額</th>
                                            <th class="p-2 border border-gray-300 text-right w-24 bg-gray-50">小計</th>
                                            <th class="p-2 border border-gray-300 text-right w-20 bg-gray-50">代購費</th>
                                            <th class="p-2 border border-gray-300 text-right w-24 bg-gray-50">總計</th>
                                        </tr>
                                    </thead>
                                    <tbody id="tableBodySheet" class="text-sm text-gray-800 font-arial bg-white">
                                        <tr><td colspan="7" class="p-8 text-center text-gray-400 italic">等待資料貼上...</td></tr>
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>

                    <!-- Diff View -->
                    <div id="viewDiff" class="hidden p-0 flex flex-col h-full">
                        <div class="bg-gray-50 p-4 border-b border-gray-200 flex-none flex flex-col md:flex-row justify-between items-center gap-4">
                            <div class="flex flex-col gap-1 w-full md:w-auto">
                                <div class="text-sm text-gray-600 flex items-center gap-2">
                                    <h3 class="font-bold text-gray-800 mb-1"><i class="fa-solid fa-magnifying-glass-chart mr-2"></i>比對結果概覽</h3>
                                    <p class="text-xs">系統自動依照 <span class="font-arial bg-gray-200 px-1 rounded mx-1">報價單</span> 進行分組比對。</p>
                                </div>
                                <div class="flex flex-wrap gap-2 text-xs font-bold">
                                    <span class="flex items-center gap-1 px-2 py-1 rounded bg-f0fdf4 text-green-700 bg-green-50"><i class="fa-solid fa-check"></i> 完全一致</span>
                                    <span class="flex items-center gap-1 px-2 py-1 rounded bg-ecfccb text-lime-800 bg-lime-100"><i class="fa-solid fa-circle-check"></i> 金額一致</span>
                                    <span class="flex items-center gap-1 px-2 py-1 rounded bg-f3e8ff text-purple-700 bg-purple-50"><i class="fa-solid fa-wand-magic-sparkles"></i> 智慧推測</span>
                                    <span class="flex items-center gap-1 px-2 py-1 rounded bg-fef2f2 text-red-700 bg-red-50"><i class="fa-solid fa-xmark"></i> 數值不符</span>
                                    <span class="flex items-center gap-1 px-2 py-1 rounded bg-eff6ff text-blue-700 bg-blue-50"><i class="fa-solid fa-arrow-left"></i> 僅填寫區</span>
                                    <span class="flex items-center gap-1 px-2 py-1 rounded bg-fff7ed text-orange-700 bg-orange-50"><i class="fa-solid fa-arrow-right"></i> 僅 PDF</span>
                                </div>
                            </div>
                            <!-- 新增複製按鈕 -->
                            <button onclick="copyCurrentView()" class="flex-none bg-white hover:bg-gray-50 text-gray-700 border border-gray-300 px-4 py-2 rounded-lg text-sm font-medium transition-colors shadow-sm flex items-center gap-2 whitespace-nowrap">
                                <i class="fa-regular fa-copy"></i> 複製比對表
                            </button>
                        </div>
                        <div class="overflow-auto max-h-[65vh] w-full relative custom-scrollbar border-b border-gray-200">
                            <table class="w-full text-left border-collapse" id="diffTable">
                                <thead class="bg-gray-100 text-gray-600 text-xs uppercase font-semibold sticky top-0 z-20 shadow-sm">
                                    <tr>
                                        <!-- Merged Cells for Status, Code, Note -->
                                        <th rowspan="2" class="p-3 border-r w-24 sticky left-0 top-0 z-30 bg-gray-100 border-b border-gray-300" style="vertical-align: middle; text-align: center;">狀態</th>
                                        <th rowspan="2" class="p-3 w-32 border-r sticky top-0 z-20 bg-gray-100 border-b border-gray-300" style="vertical-align: middle; text-align: center;">產品編號</th>
                                        
                                        <!-- SWAPPED: Sheet First (Blue, A) -->
                                        <th class="p-3 border-b border-gray-300 bg-blue-50/90 text-blue-800 sticky top-0 z-20" colspan="6" style="vertical-align: middle; text-align: center;">合約及非合約 (A)</th>
                                        
                                        <!-- SWAPPED: PDF Second (Orange, B) -->
                                        <th class="p-3 border-b border-gray-300 bg-orange-50/90 text-orange-800 sticky top-0 z-20" colspan="6" style="vertical-align: middle; text-align: center;">PDF 報價單 (B)</th>
                                        
                                        <th rowspan="2" class="p-3 border-l min-w-[250px] sticky top-0 z-20 bg-gray-100 border-b border-gray-300" style="vertical-align: middle; text-align: center;">差異說明</th>
                                    </tr>
                                    <tr>
                                        <!-- Sheet 子欄位 (原 B 現在是 A) -->
                                        <th class="p-2 text-center bg-blue-50/50 text-xs w-32 sticky top-[45px] z-10 border-b border-gray-200" style="vertical-align: middle; text-align: center;">Item</th>
                                        <th class="p-2 text-center bg-blue-50/50 text-xs w-20 sticky top-[45px] z-10 border-b border-gray-200" style="vertical-align: middle; text-align: center;">訂購量</th>
                                        <th class="p-2 text-center bg-blue-50/50 text-xs w-24 sticky top-[45px] z-10 border-b border-gray-200" style="vertical-align: middle; text-align: center;">金額</th>
                                        <th class="p-2 text-center bg-blue-50/50 text-xs w-24 sticky top-[45px] z-10 border-b border-gray-200" style="vertical-align: middle; text-align: center;">小計</th>
                                        <th class="p-2 text-center bg-blue-50/50 text-xs w-20 sticky top-[45px] z-10 border-b border-gray-200" style="vertical-align: middle; text-align: center;">代購費</th>
                                        <th class="p-2 text-center bg-blue-50/50 text-xs w-24 font-bold border-r sticky top-[45px] z-10 border-b border-gray-200" style="vertical-align: middle; text-align: center;">總計</th>
                                        
                                        <!-- PDF 子欄位 (原 A 現在是 B) -->
                                        <th class="p-2 text-center bg-orange-50/50 text-xs w-48 sticky top-[45px] z-10 border-b border-gray-200" style="vertical-align: middle; text-align: center;">產品名稱</th>
                                        <th class="p-2 text-center bg-orange-50/50 text-xs w-20 sticky top-[45px] z-10 border-b border-gray-200" style="vertical-align: middle; text-align: center;">數量</th>
                                        <th class="p-2 text-center bg-orange-50/50 text-xs w-24 sticky top-[45px] z-10 border-b border-gray-200" style="vertical-align: middle; text-align: center;">單價</th>
                                        <th class="p-2 text-center bg-orange-50/50 text-xs w-24 sticky top-[45px] z-10 border-b border-gray-200" style="vertical-align: middle; text-align: center;">原始金額</th>
                                        <th class="p-2 text-center bg-orange-50/50 text-xs w-20 sticky top-[45px] z-10 border-b border-gray-200" style="vertical-align: middle; text-align: center;">代購費</th>
                                        <th class="p-2 text-center bg-orange-50/50 text-xs w-24 font-bold sticky top-[45px] z-10 border-b border-gray-200" style="vertical-align: middle; text-align: center;">總金額</th>
                                        
                                    </tr>
                                </thead>
                                <tbody id="tableBodyDiff" class="text-sm divide-y divide-gray-100">
                                    <tr><td colspan="15" class="p-8 text-center text-gray-400">請先上傳 PDF 並在填寫區貼上資料以進行比對</td></tr>
                                </tbody>
                            </table>
                        </div>
                    </div>

                </div>
            </div>
        </div>
        <div class="h-10"></div>
    </main>

    <!-- Custom Modal for Delete Confirmation -->
    <div id="confirmModal" class="hidden fixed inset-0 z-50 flex items-center justify-center bg-gray-900 bg-opacity-50 backdrop-blur-sm transition-opacity">
        <div class="bg-white rounded-lg shadow-xl p-6 w-full max-w-sm transform scale-100 transition-transform">
            <h3 class="text-lg font-bold text-gray-900 mb-2"><i class="fa-solid fa-triangle-exclamation text-red-500 mr-2"></i>確認移除?</h3>
            <p class="text-gray-600 mb-6">您確定要移除此報價單嗎？此動作無法復原。</p>
            <div class="flex justify-end gap-3">
                <button onclick="closeConfirmModal()" class="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-lg text-sm font-medium transition-colors">取消</button>
                <button onclick="confirmDelete()" class="px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg text-sm font-medium transition-colors shadow-sm">移除</button>
            </div>
        </div>
    </div>

    <!-- Toast Notification -->
    <div id="toast" class="fixed bottom-5 right-5 transform translate-y-20 opacity-0 transition-all duration-300 z-50">
        <div class="bg-gray-800 text-white px-6 py-3 rounded-lg shadow-lg flex items-center gap-3">
            <i class="fa-solid fa-circle-check text-green-400"></i>
            <span id="toastMsg">操作成功</span>
        </div>
    </div>

    <script>
        // --- 全域變數 ---
        let allFilesData = []; 
        let sheetDataItems = []; 
        let currentTab = 'detail'; 
        let fileIdToDelete = null;

        // UI 元素
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const uploadPrompt = document.getElementById('uploadPrompt');
        const loadingPrompt = document.getElementById('loadingPrompt');
        const loadingText = document.getElementById('loadingText');
        const resultSection = document.getElementById('resultSection');
        const tableBodyDetail = document.getElementById('tableBodyDetail');
        const tableBodySheet = document.getElementById('tableBodySheet');
        const tableBodyDiff = document.getElementById('tableBodyDiff');
        const errorMsg = document.getElementById('errorMsg');
        const errorList = document.getElementById('errorList');
        const fileCountBadge = document.getElementById('fileCountBadge');
        const clearBtn = document.getElementById('clearBtn');
        const pdfDashboard = document.getElementById('pdfDashboard');
        const confirmModal = document.getElementById('confirmModal');
        
        const tabDetail = document.getElementById('tabDetail');
        const tabSheet = document.getElementById('tabSheet');
        const tabDiff = document.getElementById('tabDiff');
        const viewDetail = document.getElementById('viewDetail');
        const viewSheet = document.getElementById('viewSheet');
        const viewDiff = document.getElementById('viewDiff');
        
        const sheetInputArea = document.getElementById('sheetPasteInput');

        // --- 事件監聽 ---
        dropZone.addEventListener('click', () => fileInput.click());
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('drag-active');
        });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-active'));
        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('drag-active');
            if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
        });
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length) handleFiles(e.target.files);
        });
        
        // Sheet Paste Logic
        sheetInputArea.addEventListener('input', (e) => {
            parseAndRenderSheetData(e.target.value);
        });

        // --- Tab 切換 ---
        function switchTab(tab) {
            currentTab = tab;
            resultSection.classList.remove('hidden');
            [tabDetail, tabSheet, tabDiff].forEach(t => t.classList.remove('tab-active'));
            [viewDetail, viewSheet, viewDiff].forEach(v => v.classList.add('hidden'));

            if (tab === 'detail') {
                tabDetail.classList.add('tab-active');
                viewDetail.classList.remove('hidden');
                updateDashboardVisibility(true);
            } else if (tab === 'sheet') {
                tabSheet.classList.add('tab-active');
                viewSheet.classList.remove('hidden');
                updateDashboardVisibility(false);
            } else if (tab === 'diff') {
                tabDiff.classList.add('tab-active');
                viewDiff.classList.remove('hidden');
                updateDashboardVisibility(false);
                renderDiffView();
            }
        }

        function updateDashboardVisibility(show) {
            if (show && allFilesData.length > 0) {
                dropZone.classList.add('hidden');
                pdfDashboard.classList.remove('hidden');
            } else {
                dropZone.classList.add('hidden'); 
                pdfDashboard.classList.add('hidden');
                if (currentTab === 'detail' && allFilesData.length === 0) {
                    dropZone.classList.remove('hidden');
                }
            }
        }

        // --- Remove File ---
        function removeFile(id, event) {
            if(event) event.stopPropagation();
            fileIdToDelete = id;
            confirmModal.classList.remove('hidden');
        }

        function closeConfirmModal() {
            confirmModal.classList.add('hidden');
            fileIdToDelete = null;
        }

        function confirmDelete() {
            if (!fileIdToDelete) return;
            allFilesData = allFilesData.filter(f => f.id !== fileIdToDelete);
            if (allFilesData.length === 0) {
                resetApp();
            } else {
                renderAllData();
                updateGlobalSummary();
                showToast('已移除報價單');
            }
            closeConfirmModal();
        }

        // --- Sheet 資料處理 ---
        function parseAndRenderSheetData(text) {
            sheetDataItems = []; 
            if (!text.trim()) {
                tableBodySheet.innerHTML = '<tr><td colspan="7" class="p-8 text-center text-gray-400 italic">等待資料貼上...</td></tr>';
                return;
            }
            const rows = text.trim().split('\n');
            let html = '';
            
            const fmt = (n) => (n !== undefined && !isNaN(n)) ? n.toLocaleString() : '';

            rows.forEach((row) => {
                let cols = row.split('\t');
                let subtotal = parseNumber(cols[4]);
                let fee = parseNumber(cols[5]);
                let total = parseNumber(cols[6]);
                if (isNaN(total)) { fee = 0; total = subtotal; }
                const item = {
                    itemNo: (cols[0] || '').trim(),
                    code: (cols[1] || '').trim(),
                    qty: parseNumber(cols[2]),
                    price: parseNumber(cols[3]),
                    subtotal: subtotal,
                    fee: fee,
                    total: total
                };
                if (!item.code && isNaN(item.qty) && isNaN(item.price)) {} 
                else { if(item.code || item.itemNo) sheetDataItems.push(item); }
                
                html += `<tr class="hover:bg-blue-50 transition-colors group">`;
                html += `<td class="p-2 border border-gray-300 font-arial text-gray-700 whitespace-nowrap overflow-hidden text-ellipsis">${item.itemNo}</td>`;
                html += `<td class="p-2 border border-gray-300 font-arial text-gray-700 whitespace-nowrap overflow-hidden text-ellipsis">${item.code}</td>`;
                html += `<td class="p-2 border border-gray-300 font-arial text-gray-700 whitespace-nowrap overflow-hidden text-ellipsis text-right">${fmt(item.qty)}</td>`;
                html += `<td class="p-2 border border-gray-300 font-arial text-gray-700 whitespace-nowrap overflow-hidden text-ellipsis text-right">${fmt(item.price)}</td>`;
                html += `<td class="p-2 border border-gray-300 font-arial text-gray-700 whitespace-nowrap overflow-hidden text-ellipsis text-right">${fmt(item.subtotal)}</td>`;
                html += `<td class="p-2 border border-gray-300 font-arial text-gray-700 whitespace-nowrap overflow-hidden text-ellipsis text-right">${fmt(item.fee)}</td>`;
                html += `<td class="p-2 border border-gray-300 font-arial text-gray-700 whitespace-nowrap overflow-hidden text-ellipsis text-right font-bold">${fmt(item.total)}</td>`;
                html += `</tr>`;
            });
            tableBodySheet.innerHTML = html;
        }

        function clearSheetInput() {
            if(sheetInputArea) {
                sheetInputArea.value = '';
                parseAndRenderSheetData('');
                showToast('已清空填寫區');
            }
        }

        // --- 核心：差異比對邏輯 ---
        function getCharacterSet(str) {
            const s = (str || '').toString().toLowerCase().replace(/[()（）\s\-_\[\]×*xX]/g, '');
            return { raw: s, set: new Set(Array.from(s)) };
        }

        function calculateHybridScore(str1, str2) {
            const data1 = getCharacterSet(str1);
            const data2 = getCharacterSet(str2);
            if (data1.set.size === 0 || data2.set.size === 0) return 0;
            let intersection = 0;
            data1.set.forEach(char => { if (data2.set.has(char)) intersection++; });
            const minLen = Math.min(data1.set.size, data2.set.size);
            const baseScore = minLen === 0 ? 0 : intersection / minLen;
            let maxConsecutive = 0;
            const s1 = data1.raw; const s2 = data2.raw;
            for(let i=0; i<s1.length; i++) {
                for(let j=0; j<s2.length; j++) {
                    let k = 0;
                    while(i+k < s1.length && j+k < s2.length && s1[i+k] === s2[j+k]) { k++; }
                    if(k > maxConsecutive) maxConsecutive = k;
                }
            }
            const bonus = maxConsecutive >= 2 ? Math.min(maxConsecutive * 0.1, 0.5) : 0;
            return baseScore + bonus;
        }

        function renderDiffView() {
            tableBodyDiff.innerHTML = '';
            if (allFilesData.length === 0 && sheetDataItems.length === 0) {
                tableBodyDiff.innerHTML = '<tr><td colspan="15" class="p-12 text-center text-gray-400"><i class="fa-solid fa-inbox text-4xl mb-4 block"></i>無資料可比對，請先匯入 PDF 並貼上填寫區資料</td></tr>';
                return;
            }

            // 1. Prepare Sheet Items with index for stable sorting and usage tracking
            const sheetItemsPool = sheetDataItems.map((item, index) => ({
                ...item,
                originalIndex: index,
                isUsed: false
            }));

            let html = '';
            const fmt = (n) => (n !== undefined && !isNaN(n)) ? n.toLocaleString() : '-';

            // 2. Iterate each PDF File (Quote) and perform matching
            allFilesData.forEach(file => {
                const pdfItems = file.items.map(i => ({...i})); // shallow copy
                const matchedPairs = [];
                const pdfUsedIndices = new Set(); // indices within this file's items array

                // --- Matching Logic (Scoped to this file and unused sheet items) ---
                
                // Phase 1: Exact Code Match
                pdfItems.forEach((pdf, pIdx) => {
                    const sIdx = sheetItemsPool.findIndex(sheet => !sheet.isUsed && sheet.code === pdf.code && sheet.code !== '');
                    if (sIdx !== -1) {
                        matchedPairs.push({ pdf, sheet: sheetItemsPool[sIdx], matchType: 'code' });
                        pdfUsedIndices.add(pIdx);
                        sheetItemsPool[sIdx].isUsed = true;
                    }
                });

                // Phase 2: Fuzzy Name Match
                pdfItems.forEach((pdf, pIdx) => {
                    if (pdfUsedIndices.has(pIdx)) return;
                    let bestMatchIdx = -1; let bestScore = 0;
                    sheetItemsPool.forEach((sheet, sIdx) => {
                        if (sheet.isUsed) return;
                        const score = calculateHybridScore(pdf.name, sheet.itemNo);
                        if (score > 0.8 && score > bestScore) { bestScore = score; bestMatchIdx = sIdx; }
                    });
                    if (bestMatchIdx !== -1) {
                        matchedPairs.push({ pdf, sheet: sheetItemsPool[bestMatchIdx], matchType: 'fuzzy' });
                        pdfUsedIndices.add(pIdx);
                        sheetItemsPool[bestMatchIdx].isUsed = true;
                    }
                });

                // Phase 3: Smart Match (Lower threshold)
                const potentialMatches = [];
                pdfItems.forEach((pdf, pIdx) => {
                    if (pdfUsedIndices.has(pIdx)) return;
                    sheetItemsPool.forEach((sheet, sIdx) => {
                        if (sheet.isUsed) return;
                        const score = calculateHybridScore(pdf.name, sheet.itemNo);
                        if (score >= 0.15) { potentialMatches.push({ pIdx, sIdx, score }); }
                    });
                });
                potentialMatches.sort((a, b) => b.score - a.score);
                potentialMatches.forEach(match => {
                    if (!pdfUsedIndices.has(match.pIdx) && !sheetItemsPool[match.sIdx].isUsed) {
                        matchedPairs.push({ pdf: pdfItems[match.pIdx], sheet: sheetItemsPool[match.sIdx], matchType: 'smart', score: match.score });
                        pdfUsedIndices.add(match.pIdx);
                        sheetItemsPool[match.sIdx].isUsed = true;
                    }
                });

                // Phase 4: Unmatched PDF items (in this file)
                pdfItems.forEach((pdf, idx) => {
                    if (!pdfUsedIndices.has(idx)) matchedPairs.push({ pdf, sheet: null, matchType: 'pdf-only' });
                });

                // Sort: By Sheet Original Index (A), then PDF code (B)
                matchedPairs.sort((a, b) => {
                    const indexA = (a.sheet) ? a.sheet.originalIndex : Infinity;
                    const indexB = (b.sheet) ? b.sheet.originalIndex : Infinity;
                    
                    // If one or both have sheet index, sort by that
                    if (indexA !== Infinity || indexB !== Infinity) {
                        return indexA - indexB;
                    }
                    
                    // Fallback for PDF-only items
                    return (a.pdf?.code || '').localeCompare(b.pdf?.code || '');
                });

                // --- Calculate Statistics for Header ---
                const fileTotal = file.items.reduce((sum, item) => sum + item.total, 0); // PDF Total
                const matchedSheetTotal = matchedPairs.reduce((sum, pair) => sum + (pair.sheet ? pair.sheet.total : 0), 0);
                
                // --- Generate HTML for Group Header ---
                // MODIFIED: Simple text header aligned above totals
                html += `
                    <tr class="bg-gray-50 border-t border-b border-gray-200 font-bold text-sm">
                        <td colspan="2" class="p-2 border-r text-left pl-3 text-gray-800">
                            報價單號: ${file.quoteNumber}
                        </td>
                        
                        <!-- Section A Header: Aligned Right to sit over Total -->
                        <td colspan="6" class="p-2 border-r text-right pr-3 text-blue-700 bg-blue-50/30">
                            總額 (A) $${fmt(matchedSheetTotal)}
                        </td>
                        
                        <!-- Section B Header: Aligned Right to sit over Total -->
                        <td colspan="6" class="p-2 border-r text-right pr-3 text-gray-700 bg-orange-50/30">
                            總額 (B) $${fmt(fileTotal)}
                        </td>
                        
                        <td class="p-2 border-l"></td>
                    </tr>
                `;

                // --- Generate HTML for Rows ---
                matchedPairs.forEach(pair => {
                    html += generateRowHtml(pair, fmt);
                });
            });

            // 3. Unmatched Sheet Items (Global Leftovers)
            const leftovers = sheetItemsPool.filter(item => !item.isUsed);
            if (leftovers.length > 0) {
                 // Sort leftovers by original index
                 leftovers.sort((a,b) => a.originalIndex - b.originalIndex);
                 
                 const leftoverTotal = leftovers.reduce((sum, item) => sum + item.total, 0);

                 // MODIFIED: Simple header for leftovers
                 html += `
                    <tr class="bg-blue-50 border-t border-b border-blue-100 font-bold text-sm">
                        <td colspan="2" class="p-2 border-r text-left pl-3 text-blue-800">
                            未匹配的合約/非合約項目 (${leftovers.length} 筆)
                        </td>
                         <td colspan="6" class="p-2 border-r text-right pr-3 text-blue-700">
                            總計: $${fmt(leftoverTotal)}
                        </td>
                        <td colspan="7" class="p-2"></td>
                    </tr>
                 `;
                 
                 leftovers.forEach(sheetItem => {
                     // Create a pair with null PDF
                     const pair = { pdf: null, sheet: sheetItem, matchType: 'sheet-only' };
                     html += generateRowHtml(pair, fmt);
                 });
            }

            tableBodyDiff.innerHTML = html;
        }

        // Helper for generating row HTML
        function generateRowHtml(pair, fmt) {
             const { pdf, sheet, matchType } = pair;
             let status = '', statusClass = '', diffNote = [];
             // Use Sheet code/item if available as primary display code
             const displayCode = sheet?.code || pdf?.code || '?';
             
             // 預設文字顏色
             let sheetColor = "text-blue-700";
             let pdfColor = "text-gray-600";
             
             // 差異高亮樣式 (紅字)
             let hlQty = "", hlPrice = "", hlSub = "", hlFee = "", hlTotal = "";

             if (pdf && sheet) {
                 // Ensure we are comparing numbers for price
                 // PDF price often comes as string "4,351.00", remove comma and parse
                 const pdfPriceNum = typeof pdf.price === 'string' ? parseNumber(pdf.price) : pdf.price;
                 const sheetPriceNum = typeof sheet.price === 'string' ? parseNumber(sheet.price) : sheet.price;

                 const qtyMatch = pdf.qty === sheet.qty;
                 // Use small epsilon for float comparison just in case
                 const priceMatch = Math.abs(pdfPriceNum - sheetPriceNum) < 0.01;
                 const amountMatch = Math.abs(pdf.amount - sheet.subtotal) < 1;
                 const feeMatch = Math.abs(pdf.fee - sheet.fee) < 1; 
                 const totalMatch = Math.abs(pdf.total - sheet.total) < 1;

                 // 若不一致，設定紅色 (移除 font-bold)
                 if (!qtyMatch) hlQty = "text-red-600 bg-red-50";
                 if (!priceMatch) hlPrice = "text-red-600 bg-red-50";
                 if (!amountMatch) hlSub = "text-red-600 bg-red-50";
                 if (!feeMatch) hlFee = "text-red-600 bg-red-50";
                 if (!totalMatch) hlTotal = "text-red-600 bg-red-50";

                 if (totalMatch) {
                     if (qtyMatch && amountMatch && priceMatch && feeMatch) { 
                         status = '<i class="fa-solid fa-check mr-1"></i>完全一致'; 
                         statusClass = 'diff-match font-bold'; 
                     } else { 
                         status = '<i class="fa-solid fa-circle-check mr-1"></i>金額一致'; 
                         statusClass = 'diff-ok-money font-bold'; 
                         if (!qtyMatch) { diffNote.push(`數量差異 (${sheet.qty} vs ${pdf.qty})`); diffNote.push(`<span class="text-xs text-gray-500 block">⚠️ 可能為單位不同</span>`); }
                         if (!priceMatch) diffNote.push(`單價差異`); 
                         if (!amountMatch) diffNote.push(`小計差異`);
                         if (!feeMatch) diffNote.push(`代購費差異`);
                     }
                 } else {
                     status = '<i class="fa-solid fa-triangle-exclamation mr-1"></i>數值不符'; 
                     statusClass = 'diff-error font-bold';
                     if (!qtyMatch) diffNote.push(`數量差異`); 
                     if (!priceMatch) diffNote.push(`單價差異`); 
                     if (!amountMatch) diffNote.push(`小計差異`); 
                     if (!feeMatch) diffNote.push(`代購費差異`);
                     if (!totalMatch) diffNote.push(`總計差異`);
                 }
                 if (matchType === 'fuzzy') { if(!status.includes('數值不符')) { statusClass = statusClass.replace('diff-match', 'diff-match').replace('diff-ok-money', 'diff-ok-money'); } diffNote.unshift(`<span class="text-amber-700 font-bold text-xs">名稱相似</span>`); } 
                 else if (matchType === 'smart') { diffNote.unshift(`<span class="text-purple-700 font-bold text-xs">智慧推測</span>`); }
             } else if (pdf) { status = '<i class="fa-solid fa-arrow-right mr-1"></i>僅 PDF'; statusClass = 'diff-warn font-bold'; diffNote.push('填寫區缺少'); } 
             else if (sheet) { status = '<i class="fa-solid fa-arrow-left mr-1"></i>僅填寫區'; statusClass = 'diff-info font-bold'; diffNote.push('PDF 缺少'); }

             return `<tr class="hover:bg-gray-50 border-b border-gray-100 group">
                 <td class="p-3 border-r sticky left-0 bg-white group-hover:bg-gray-50 z-10 ${statusClass} text-xs whitespace-nowrap" style="vertical-align: middle; text-align: left;">${status}</td>
                 <td class="p-3 border-r font-arial text-gray-700 font-medium" style="vertical-align: middle; text-align: left;">${displayCode}</td>
                 
                 <!-- Sheet (A) Data First -->
                 <td class="p-3 font-arial ${sheetColor} ${sheet?'':'bg-gray-50 text-gray-300'} text-xs break-words min-w-[150px]" title="${sheet?.itemNo||''}" style="vertical-align: middle; text-align: left;">${sheet ? sheet.itemNo : '-'}</td>
                 <td class="p-3 font-arial ${sheetColor} ${hlQty} ${sheet?'':'bg-gray-50 text-gray-300'}" style="vertical-align: middle; text-align: right;">${sheet ? fmt(sheet.qty) : '-'}</td>
                 <td class="p-3 font-arial ${sheetColor} ${hlPrice} ${sheet?'':'bg-gray-50 text-gray-300'} text-xs" style="vertical-align: middle; text-align: right;">${sheet ? fmt(sheet.price) : '-'}</td>
                 <td class="p-3 font-arial ${sheetColor} ${hlSub} ${sheet?'':'bg-gray-50 text-gray-300'} text-xs" style="vertical-align: middle; text-align: right;">${sheet ? fmt(sheet.subtotal) : '-'}</td>
                 <td class="p-3 font-arial ${sheetColor} ${hlFee} ${sheet?'':'bg-gray-50 text-gray-300'} text-xs" style="vertical-align: middle; text-align: right;">${sheet ? fmt(sheet.fee) : '-'}</td>
                 <td class="p-3 font-arial ${sheet?'text-blue-800':'bg-gray-50 text-gray-300'} ${hlTotal} border-r font-medium" style="vertical-align: middle; text-align: right;">${sheet ? fmt(sheet.total) : '-'}</td>
                 
                 <!-- PDF (B) Data Second -->
                 <td class="p-3 font-arial ${pdfColor} ${pdf?'':'bg-gray-50 text-gray-300'} text-xs break-words min-w-[150px]" title="${pdf?.name||''}" style="vertical-align: middle; text-align: left;">${pdf ? pdf.name : '-'}</td>
                 <td class="p-3 font-arial ${pdfColor} ${hlQty} ${pdf?'':'bg-gray-50 text-gray-300'}" style="vertical-align: middle; text-align: right;">${pdf ? fmt(pdf.qty) : '-'}</td>
                 <td class="p-3 font-arial ${pdfColor} ${hlPrice} ${pdf?'':'bg-gray-50 text-gray-300'} text-xs" style="vertical-align: middle; text-align: right;">${pdf ? pdf.price : '-'}</td>
                 <td class="p-3 font-arial ${pdfColor} ${hlSub} ${pdf?'':'bg-gray-50 text-gray-300'} text-xs" style="vertical-align: middle; text-align: right;">${pdf ? fmt(pdf.amount) : '-'}</td>
                 <td class="p-3 font-arial ${pdfColor} ${hlFee} ${pdf?'':'bg-gray-50 text-gray-300'} text-xs" style="vertical-align: middle; text-align: right;">${pdf ? fmt(pdf.fee) : '-'}</td>
                 <td class="p-3 font-arial ${pdf?'text-gray-800':'bg-gray-50 text-gray-300'} ${hlTotal} font-medium" style="vertical-align: middle; text-align: right;">${pdf ? fmt(pdf.total) : '-'}</td>
                 
                 <td class="p-3 border-l text-xs text-red-600 leading-relaxed" style="vertical-align: middle; text-align: left;">${diffNote.join('<br>')}</td>
             </tr>`;
        }

        async function handleFiles(fileList) {
            const files = Array.from(fileList).filter(f => f.type === 'application/pdf');
            if (files.length === 0) { showToast('請上傳 PDF 格式的檔案', 'error'); return; }
            if (files.length > 20) { showToast('一次最多支援 20 個檔案，已自動截斷', 'warning'); files.length = 20; }
            uploadPrompt.classList.add('hidden'); loadingPrompt.classList.remove('hidden'); dropZone.classList.add('pointer-events-none', 'bg-gray-50'); loadingText.innerText = `正在解析 ${files.length} 個檔案...`;
            const processingPromises = files.map(file => processSingleFile(file));
            try {
                const results = await Promise.allSettled(processingPromises);
                const errors = []; let newFilesCount = 0;
                results.forEach((result, index) => {
                    if (result.status === 'fulfilled') {
                        const exists = allFilesData.some(f => f.fileName === result.value.fileName && f.quoteNumber === result.value.quoteNumber);
                        if (!exists) { allFilesData.push(result.value); newFilesCount++; }
                    } else { errors.push(`檔案 ${files[index].name}: ${result.reason.message}`); }
                });
                if (errors.length > 0) showErrorList(errors); else errorMsg.classList.add('hidden');
                if (newFilesCount > 0) {
                    allFilesData.sort((a, b) => { const dateDiff = b.quoteDate.localeCompare(a.quoteDate); if (dateDiff !== 0) return dateDiff; return a.quoteNumber.localeCompare(b.quoteNumber); });
                    renderAllData(); updateGlobalSummary(); switchTab('detail'); clearBtn.classList.remove('hidden'); showToast(`成功解析 ${newFilesCount} 個檔案`);
                } else if (allFilesData.length === 0) { showToast('未能解析出有效資料', 'error'); resetUIState(); }
            } catch (e) { console.error(e); showToast('系統發生錯誤', 'error'); resetUIState(); }
        }

        async function processSingleFile(file) {
            try {
                const arrayBuffer = await file.arrayBuffer(); const textContent = await extractTextFromPDF(arrayBuffer); const quoteInfo = extractQuoteInfo(textContent); const rawRows = parseTextToRows(textContent); const processedItems = processLogic(rawRows);
                if (processedItems.length === 0) throw new Error("無法識別表格");
                return { id: 'f-' + Math.random().toString(36).substr(2, 9), fileName: file.name, quoteNumber: quoteInfo.number || '未知單號', quoteDate: quoteInfo.date, items: processedItems, isCollapsed: false };
            } catch (err) { throw err; }
        }

        async function extractTextFromPDF(arrayBuffer) {
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            let fullTextLines = [];
            for (let i = 1; i <= pdf.numPages; i++) {
                try {
                    const page = await pdf.getPage(i); const textContent = await page.getTextContent();
                    const items = textContent.items.map(item => ({ str: item.str, x: item.transform[4], y: item.transform[5], w: item.width }));
                    items.sort((a, b) => b.y - a.y || a.x - b.x);
                    let currentLine = [], lastY = -1;
                    items.forEach(item => {
                        if (lastY === -1 || Math.abs(item.y - lastY) < 5) { currentLine.push(item); lastY = item.y; }
                        else { fullTextLines.push(sortAndJoinLine(currentLine)); currentLine = [item]; lastY = item.y; }
                    });
                    if (currentLine.length > 0) fullTextLines.push(sortAndJoinLine(currentLine));
                } catch (e) {}
            }
            return fullTextLines;
        }

        function sortAndJoinLine(lineItems) {
            lineItems.sort((a, b) => a.x - b.x);
            return lineItems.map((item, idx) => { if (idx > 0 && (item.x - (lineItems[idx-1].x + lineItems[idx-1].w)) > 5) return " " + item.str; return item.str; }).join('').trim();
        }

        function parseNumber(str) { if (!str) return NaN; const num = parseFloat(String(str).replace(/,/g, '').trim()); return isNaN(num) ? NaN : num; }
        function isNumericAndNonZero(str) { const num = parseNumber(str); return !isNaN(num) && num !== 0; }

        function extractQuoteInfo(lines) {
            let number = '', date = ''; const d = new Date(); const todayStr = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
            for (const line of lines) {
                let dateMatch = line.match(/報價日期\s*[:：]\s*([\d\/]+)/); if (dateMatch && dateMatch[1]) date = dateMatch[1].replace(/\//g, '');
                let numberMatch = line.match(/報價單[\s\n]*號\s*[:：]\s*([A-Za-z0-9-]+)/); if (numberMatch && numberMatch[1]) number = numberMatch[1];
                if (date && number) break;
            }
            return { number, date: date || todayStr };
        }

        function parseTextToRows(lines) {
            const extractedRows = [], feeRegex = /(代購管理費|ZZ14)/i;
            lines.forEach(line => {
                const cleanLine = line.replace(/\s+/g, ' ').trim(); if (!cleanLine) return;
                if (feeRegex.test(cleanLine)) { const matches = cleanLine.match(/([\d,]+\.?\d*)\s*$/); if (matches) extractedRows.push({ type: 'fee', amount: parseNumber(matches[1]) }); return; }
                const parts = cleanLine.split(' '); if (parts.length < 5) return; if (!/^\d{3}$/.test(parts[0])) return;
                try {
                    let amountIndex = -1;
                    for (let i = parts.length - 1; i >= 4; i--) { if (isNumericAndNonZero(parts[i])) { if (isNumericAndNonZero(parts[i-1]) && isNumericAndNonZero(parts[i-3])) { amountIndex = i; break; } } }
                    if (amountIndex === -1) { if (isNumericAndNonZero(parts[parts.length-1]) && isNumericAndNonZero(parts[parts.length-2]) && isNumericAndNonZero(parts[parts.length-4])) { amountIndex = parts.length - 1; } }
                    if (amountIndex === -1) return;
                    const amount = parseNumber(parts[amountIndex]); const price = parts[amountIndex - 1]; const unit = parts[amountIndex - 2]; const qty = parts[amountIndex - 3];
                    const nameParts = parts.slice(2, amountIndex - 3); const remarksParts = parts.slice(amountIndex + 1);
                    if (nameParts.length >= 1 && !isNaN(amount)) { extractedRows.push({ type: 'product', id: parts[0], code: parts[1], name: nameParts.join(' '), qty, unit, price, amount, remarks: remarksParts.join(' ') }); }
                } catch (e) {}
            });
            return extractedRows;
        }

        function processLogic(rows) {
            const finalData = [];
            rows.forEach(row => {
                if (row.type === 'product') { const qty = Math.round(row.qty); finalData.push({ ...row, qty: qty, fee: 0, total: row.amount }); }
                else if (row.type === 'fee') { if (finalData.length > 0) { const lastItem = finalData[finalData.length - 1]; lastItem.fee = (lastItem.fee || 0) + row.amount; lastItem.total = lastItem.amount + lastItem.fee; } }
            });
            return finalData;
        }

        function renderAllData() {
            tableBodyDetail.innerHTML = '';
            allFilesData.forEach(file => {
                const fileTotalOriginal = file.items.reduce((s,i)=>s+i.amount,0); const fileTotalFee = file.items.reduce((s,i)=>s+i.fee,0); const fileTotalSum = file.items.reduce((s,i)=>s+i.total,0);
                const detailHeader = document.createElement('tr');
                detailHeader.innerHTML = `
                    <td colspan="11" class="bg-gray-100 p-0 border-b border-gray-200">
                        <div class="flex justify-between items-center p-4 hover:bg-gray-200 cursor-pointer select-none" onclick="toggleFileDetails('${file.id}')">
                            <div class="flex items-center gap-4">
                                <span class="bg-white border text-gray-400 rounded-full w-6 h-6 flex items-center justify-center text-xs"><i class="fa-solid fa-chevron-${file.isCollapsed?'right':'down'}"></i></span>
                                <div><span class="font-bold text-gray-800">${file.quoteNumber}</span><span class="text-xs text-gray-500 ml-2">${file.quoteDate}</span></div>
                            </div>
                            <div class="flex items-center gap-4">
                                <span class="font-bold text-green-700 font-arial">$${fileTotalSum.toLocaleString()}</span>
                                <button type="button" onclick="event.stopPropagation(); removeFile('${file.id}', event)" class="text-gray-400 hover:text-red-500 p-2 rounded-full hover:bg-red-50 transition-colors z-10 relative"><i class="fa-solid fa-xmark pointer-events-none"></i></button>
                            </div>
                        </div>
                    </td>`;
                tableBodyDetail.appendChild(detailHeader);
                if (!file.isCollapsed) {
                    file.items.forEach((item, idx) => {
                        const tr = document.createElement('tr');
                        tr.className = `border-b border-gray-100 hover:bg-blue-50/30 ${file.id}-row`;
                        tr.innerHTML = `
                            <td class="p-3 text-center text-gray-400 text-xs">${idx+1}</td>
                            <td class="p-3 font-arial text-gray-600 text-xs">${item.id}</td>
                            <td class="p-3 font-arial text-gray-600 text-xs">${item.code}</td>
                            <td class="p-3 text-sm">${item.name}</td> 
                            <td class="p-3 text-right font-arial">${item.qty}</td>
                            <td class="p-3 text-center text-xs">${item.unit}</td>
                            <td class="p-3 text-right font-arial text-xs">${item.price}</td>
                            <td class="p-3 text-right font-arial text-gray-600">${item.amount.toLocaleString()}</td>
                            <td class="p-3 text-right font-arial ${item.fee>0?'text-blue-600 font-bold':'text-gray-300 text-xs'}">${item.fee>0?item.fee.toLocaleString():'-'}</td>
                            <td class="p-3 text-right font-arial font-bold text-green-700 bg-green-50/50">${item.total.toLocaleString()}</td>
                            <td class="p-3 text-xs text-red-500">${item.remarks||''}</td>`;
                        tableBodyDetail.appendChild(tr);
                    });
                }
            });
            updateFileCountBadge();
        }

        function toggleFileDetails(fileId) { const file = allFilesData.find(f => f.id === fileId); if (file) { file.isCollapsed = !file.isCollapsed; renderAllData(); } }
        function updateGlobalSummary() { let tO=0, tF=0, tT=0; allFilesData.forEach(f => f.items.forEach(i => { tO+=i.amount; tF+=i.fee; tT+=i.total; })); document.getElementById('sumOriginal').innerText = '$'+tO.toLocaleString(); document.getElementById('sumFee').innerText = '$'+tF.toLocaleString(); document.getElementById('sumTotal').innerText = '$'+tT.toLocaleString(); }
        function updateFileCountBadge() { fileCountBadge.innerText = `${allFilesData.length} 份`; }
        function showErrorList(errors) { errorMsg.classList.remove('hidden'); errorList.innerHTML = errors.map(e=>`<li>${e}</li>`).join(''); }
        function showToast(msg, type='success') { const t = document.getElementById('toast'), tm = document.getElementById('toastMsg'); tm.innerText = msg; t.classList.remove('translate-y-20', 'opacity-0'); setTimeout(()=>t.classList.add('translate-y-20', 'opacity-0'), 3000); }
        function resetUIState() { loadingPrompt.classList.add('hidden'); uploadPrompt.classList.remove('hidden'); dropZone.classList.remove('pointer-events-none', 'bg-gray-50', 'hidden'); fileInput.value = ''; }
        function resetApp() { allFilesData=[]; tableBodyDetail.innerHTML=''; resultSection.classList.add('hidden'); errorMsg.classList.add('hidden'); clearBtn.classList.add('hidden'); resetUIState(); if(currentTab === 'sheet') { resultSection.classList.remove('hidden'); } document.getElementById('sumOriginal').innerText='$0'; document.getElementById('sumFee').innerText='$0'; document.getElementById('sumTotal').innerText='$0'; }
        function copyCurrentView() {
            let targetId = ''; if (currentTab === 'detail') targetId = 'tableBodyDetail'; else if (currentTab === 'sheet') targetId = 'sheetTable'; else if (currentTab === 'diff') targetId = 'diffTable';
            const el = document.getElementById(targetId); if(!el || (currentTab === 'detail' && allFilesData.length === 0)) return showToast('無內容可複製', 'error');
            const range = document.createRange(); if(currentTab === 'detail') range.selectNode(el.parentNode); else range.selectNode(el);
            const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);
            try { document.execCommand('copy'); showToast('已複製表格！'); } catch(e) { showToast('複製失敗', 'error'); } sel.removeAllRanges();
        }
    </script>
</body>
</html>
