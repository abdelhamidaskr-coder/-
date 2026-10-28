<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>نظام إدارة طلبيات المدارس والمخبوزات التفاعلي</title>

    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    
    <!-- Font Arabic Cairo -->
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap" rel="stylesheet">

    <style>
        body {
            font-family: 'Cairo', sans-serif;
            background-color: #f1f5f9;
            color: #0f172a;
        }

        @media print {
            .no-print {
                display: none !important;
            }
            body {
                background-color: #ffffff;
                padding: 0;
            }
            .container {
                max-width: 100% !important;
                width: 100% !important;
                margin: 0 !important;
                padding: 0 !important;
            }
            .card-box {
                box-shadow: none !important;
                border: 1px solid #cbd5e1 !important;
            }
            input {
                border: none !important;
                background: transparent !important;
            }
        }

        .qty-input:focus {
            background-color: #fef08a !important;
            outline: 2px solid #ca8a04 !important;
        }
    </style>
</head>
<body class="p-3 md:p-6">

    <div class="max-w-[1600px] mx-auto space-y-6 container">
        
        <!-- Header & Top Actions -->
        <div class="bg-gradient-to-r from-emerald-900 to-teal-800 text-white rounded-2xl p-6 shadow-xl flex flex-col md:flex-row justify-between items-center gap-4 card-box">
            <div>
                <h1 class="text-2xl md:text-3xl font-extrabold flex items-center gap-3">
                    <span>🥐</span> نظام إدارة طلبيات المدارس المطور
                </h1>
                <p class="text-emerald-200 text-xs md:text-sm mt-1">تعديل المدارس، الأصناف، الأسعار، والأعداد مباشرة مع التحديث التلقائي والطباعة والتصدير</p>
            </div>
            
            <div class="flex flex-wrap gap-2 md:gap-3 no-print">
                <button id="btnOpenSchoolModal" class="bg-blue-600 hover:bg-blue-700 text-white font-bold px-4 py-2.5 rounded-xl shadow transition text-sm flex items-center gap-1.5">
                    ➕ إضافة مدرسة جديدة
                </button>
                <button id="btnOpenItemModal" class="bg-purple-600 hover:bg-purple-700 text-white font-bold px-4 py-2.5 rounded-xl shadow transition text-sm flex items-center gap-1.5">
                    🥯 إضافة صنف جديد
                </button>
                <button id="btnPrint" class="bg-amber-500 hover:bg-amber-600 text-slate-900 font-bold px-4 py-2.5 rounded-xl shadow transition text-sm flex items-center gap-1.5">
                    🖨️ طباعة التقرير
                </button>
                <button id="btnExport" class="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-4 py-2.5 rounded-xl shadow transition text-sm flex items-center gap-1.5">
                    📥 تصدير Excel
                </button>
            </div>
        </div>

        <!-- Dashboard Cards -->
        <div class="grid grid-cols-1 md:grid-cols-4 gap-4">
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 card-box border-r-8 border-r-emerald-700">
                <span class="text-slate-500 font-semibold text-xs">إجمالي القطع المطلوبة</span>
                <div class="text-2xl font-extrabold text-emerald-900 mt-1" id="dashTotalQty">0</div>
                <span class="text-[10px] text-slate-400">حبة كلياً لجميع المدارس</span>
            </div>
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 card-box border-r-8 border-r-amber-500">
                <span class="text-slate-500 font-semibold text-xs">إجمالي المبلغ المطلوب</span>
                <div class="text-2xl font-extrabold text-amber-600 mt-1" id="dashTotalAmount">0.00 ر.س</div>
                <span class="text-[10px] text-slate-400">مجموع التكاليف الكلي</span>
            </div>
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 card-box border-r-8 border-r-blue-600">
                <span class="text-slate-500 font-semibold text-xs">عدد المدارس / الأقسام</span>
                <div class="text-2xl font-extrabold text-blue-900 mt-1" id="dashSchoolsCount">0</div>
                <span class="text-[10px] text-slate-400">مدرسة مسجلة بالدليل</span>
            </div>
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 card-box border-r-8 border-r-purple-600">
                <span class="text-slate-500 font-semibold text-xs">عدد الأصناف بالمظروف</span>
                <div class="text-2xl font-extrabold text-purple-900 mt-1" id="dashItemsCount">0</div>
                <span class="text-[10px] text-slate-400">صنف مخبوزات ومعجنات</span>
            </div>
        </div>

        <!-- TABLE 1: Interactive Matrix -->
        <div class="bg-white rounded-2xl shadow-md border border-slate-200 p-5 card-box overflow-hidden">
            <div class="flex justify-between items-center mb-3">
                <div>
                    <h2 class="text-lg md:text-xl font-bold text-emerald-900 flex items-center gap-2">
                        <span>📊</span> جدول الطلبيات والتعديل المباشر (الأصناف × المدارس)
                    </h2>
                    <p class="text-xs text-slate-500">يمكنك تعديل أي عدد حبات داخل الجدول مباشرة، أو تغيير اسم المدرسة بالنقر عليه!</p>
                </div>
            </div>

            <div class="overflow-x-auto max-h-[600px]">
                <table class="w-full text-xs text-center border-collapse" id="matrixTable">
                    <thead id="matrixThead" class="sticky top-0 bg-slate-900 text-white z-10"></thead>
                    <tbody id="matrixTbody" class="divide-y divide-slate-200"></tbody>
                    <tfoot id="matrixTfoot" class="sticky bottom-0 bg-