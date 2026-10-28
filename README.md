<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>نظام إدارة طلبيات المدارس والمخبوزات</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Cairo', sans-serif; background-color: #f1f5f9; color: #0f172a; }
        @media print {
            .no-print { display: none !important; }
            body { background-color: #ffffff; padding: 0; }
        }
        .qty-input:focus { background-color: #fef08a !important; outline: 2px solid #ca8a04 !important; }
    </style>
</head>
<body class="p-3 md:p-6">

    <div class="max-w-[1600px] mx-auto space-y-6 container">
        
        <!-- Header & Top Actions -->
        <div class="bg-gradient-to-r from-emerald-900 to-teal-800 text-white rounded-2xl p-6 shadow-xl flex flex-col md:flex-row justify-between items-center gap-4">
            <div>
                <h1 class="text-2xl md:text-3xl font-extrabold flex items-center gap-3">
                    <span>🥐</span> نظام إدارة طلبيات المدارس المطور
                </h1>
                <p class="text-emerald-200 text-xs md:text-sm mt-1">تعديل المدارس، الأصناف، الأسعار، والأعداد مباشرة مع التحديث التلقائي والطباعة والتصدير</p>
            </div>
            
            <div class="flex flex-wrap gap-2 md:gap-3 no-print">
                <button onclick="openSchoolModal()" class="bg-blue-600 hover:bg-blue-700 text-white font-bold px-4 py-2.5 rounded-xl shadow transition text-sm flex items-center gap-1.5">
                    ➕ إضافة مدرسة جديدة
                </button>
                <button onclick="openItemModal()" class="bg-purple-600 hover:bg-purple-700 text-white font-bold px-4 py-2.5 rounded-xl shadow transition text-sm flex items-center gap-1.5">
                    🥯 إضافة صنف جديد
                </button>
                <button onclick="window.print()" class="bg-amber-500 hover:bg-amber-600 text-slate-900 font-bold px-4 py-2.5 rounded-xl shadow transition text-sm flex items-center gap-1.5">
                    🖨️ طباعة التقرير
                </button>
                <button onclick="exportToExcel()" class="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-4 py-2.5 rounded-xl shadow transition text-sm flex items-center gap-1.5">
                    📥 تصدير Excel
                </button>
            </div>
        </div>

        <!-- Dashboard Cards -->
        <div class="grid grid-cols-1 md:grid-cols-4 gap-4">
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 border-r-8 border-r-emerald-700">
                <span class="text-slate-500 font-semibold text-xs">إجمالي القطع المطلوبة</span>
                <div class="text-2xl font-extrabold text-emerald-900 mt-1" id="dashTotalQty">0</div>
            </div>
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 border-r-8 border-r-amber-500">
                <span class="text-slate-500 font-semibold text-xs">إجمالي المبلغ المطلوب</span>
                <div class="text-2xl font-extrabold text-amber-600 mt-1" id="dashTotalAmount">0.00 ر.س</div>
            </div>
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 border-r-8 border-r-blue-600">
                <span class="text-slate-500 font-semibold text-xs">عدد المدارس / الأقسام</span>
                <div class="text-2xl font-extrabold text-blue-900 mt-1" id="dashSchoolsCount">0</div>
            </div>
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 border-r-8 border-r-purple-600">
                <span class="text-slate-5