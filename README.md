

<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>جدول طلبيات المدارس والمخبوزات</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Cairo', sans-serif; background-color: #f8fafc; color: #0f172a; }
        @media print {
            .no-print { display: none !important; }
            body { background-color: #ffffff; padding: 0; }
            input { border: none !important; background: transparent !important; }
        }
        .qty-box {
            width: 100%;
            text-align: center;
            font-weight: 700;
            padding: 4px 2px;
            border-radius: 6px;
            border: 1px solid transparent;
            transition: all 0.2s;
        }
        .qty-box:hover {
            border-color: #94a3b8;
            background-color: #f1f5f9;
        }
        .qty-box:focus {
            background-color: #fef08a !important;
            border-color: #ca8a04 !important;
            outline: none;
        }
        .has-qty {
            background-color: #ecfdf5;
            color: #065f46;
            font-weight: 800;
        }
    </style>
</head>
<body class="p-3 md:p-6">

    <div class="max-w-[1600px] mx-auto space-y-6">
        
        <!-- الهيدر والأزرار -->
        <div class="bg-emerald-900 text-white rounded-2xl p-5 shadow-lg flex flex-col md:flex-row justify-between items-center gap-4">
            <div>
                <h1 class="text-2xl font-extrabold flex items-center gap-2">
                    <span>🥐</span> جدول طلبيات المدارس التفاعلي
                </h1>
                <p class="text-emerald-200 text-xs md:text-sm mt-1">المدارس والأصناف ثابتة - يمكنك تعديل أي عدد أو سعر مباشرة داخل الخانات وسيتم تحديث الحساب فوراً</p>
            </div>
            
            <div class="flex gap-3 no-print">
                <button onclick="window.print()" class="bg-amber-500 hover:bg-amber-600 text-slate-900 font-bold px-5 py-2.5 rounded-xl shadow transition text-sm flex items-center gap-1.5">
                    🖨️ طباعة
                </button>
                <button onclick="exportToExcel()" class="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-5 py-2.5 rounded-xl shadow transition text-sm flex items-center gap-1.5">
                    📥 تصدير Excel
                </button>
            </div>
        </div>

        <!-- بطاقات المجموع الإجمالي -->
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 border-r-8 border-r-emerald-700">
                <span class="text-slate-500 font-semibold text-xs">إجمالي عدد القطع المطلوبة (حبة)</span>
                <div class="text-3xl font-extrabold text-emerald-900 mt-1" id="dashTotalQty">0</div>
            </div>
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 border-r-8 border-r-amber-500">
                <span class="text-slate-500 font-semibold text-xs">المبلغ الإجمالي المطلوب (ر.س)</span>
                <div class="text-3xl font-extrabold text-amber-600 mt-1" id="dashTotalAmount">0.00 ر.س</div>
            </div>
        </div>

        <!-- 1. جدول المدارس والكميات (تعديل مباشر في الخانات) -->
        <div class="bg-white rounded-2xl shadow-md border border-slate-200 p-5 overflow-hidden">
            <div class="mb-3">
                <h2 class="text-lg font-bold text-emerald-900">📊 جدول كميات المدارس (انقر على أي رقم لتعديله)</h2>
            </div>
            <div class="overflow-x-auto max-h-[600px]">
                <table class="w-full text-xs text-center border-collapse" id="matrixTable">
                    <thead id="matrixThead" class="sticky top-0 bg-slate-100 text-slate-900 z-10"></thead>
                    <tbody id="matrixTbody" class="divide-y divide-slate-200"></tbody>
                    <tfoot id="matrixTfoot" class="sticky bottom-0 bg-slate-100 font-bold border-t-2 border-slate-400 z-10"></tfoot>
                </table>
            </div>
        </div>

        <!-- 2. جدول أسعار الأصناف والملخص المالي -->
        <div class="bg-white rounded-2xl shadow-md border border-slate-200 p-5">
            <div class="mb-3">
                <h2 class="text-lg font-bold text-emerald-900">🏷️ جدول أسعار الأصناف والملخص الإجمالي</h2>
            </div>
            <div class="overflow-x-auto">
                <table class="w-full text-xs text-right border-collapse">
                    <thead>
                        <tr class="bg-emerald-900 text-white font-bold">
                            <th class="p-2.5 border text-center w-12">#</th>
                            <th class="p-2.5 border">اسم الصنف</th>
                            <th class="p-2.5 border text-center w-36">سعر الوحدة (ر.س)</th>
                            <th class="p-2.5 border text-center w-36">إجمالي الكمية المطلوبة</th>
                            <th class="p-2.5 border text