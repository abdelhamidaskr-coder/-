<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>نظام إدارة طلبيات المدارس والمخبوزات</title>
    <!-- Tailwind CSS & SheetJS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap" rel="stylesheet">
    <style>
        body { font-family: 'Cairo', sans-serif; background-color: #f8fafc; color: #1e293b; }
        @media print {
            .no-print { display: none !important; }
            body { background-color: #fff; padding: 0; }
            .container { max-width: 100% !important; width: 100% !important; margin: 0 !important; }
            .card-box { box-shadow: none !important; border: 1px solid #cbd5e1 !important; }
            table { font-size: 10px !important; }
        }
    </style>
</head>
<body class="p-4 md:p-8">
    <div class="max-w-[1500px] mx-auto space-y-8 container">
        <!-- Header & Action Buttons -->
        <div class="bg-emerald-900 text-white rounded-2xl p-6 shadow-xl flex flex-col md:flex-row justify-between items-center gap-4 card-box">
            <div>
                <h1 class="text-2xl md:text-3xl font-extrabold flex items-center gap-3">🥐 نظام إدارة طلبيات المدارس الموحد</h1>
                <p class="text-emerald-200 text-sm mt-1">حساب تلقائي للكميات والتكاليف مع إمكانية التعديل، الطباعة، والتصدير لإكسل</p>
            </div>
            <div class="flex flex-wrap gap-3 no-print">
                <button onclick="window.print()" class="bg-amber-500 hover:bg-amber-600 text-slate-900 font-bold px-5 py-2.5 rounded-xl shadow transition flex items-center gap-2">
                    🖨️ طباعة التقرير
                </button>
                <button onclick="exportToExcel()" class="bg-emerald-600 hover:bg-emerald-700 text-white font-bold px-5 py-2.5 rounded-xl shadow transition flex items-center gap-2">
                    📥 تصدير إلى Excel
                </button>
            </div>
        </div>

        <!-- Dashboard Summary -->
        <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
            <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 card-box border-r-8 border-r-emerald-700">
                <span class="text-slate-500 font-semibold text-sm">إجمالي عدد القطع المطلوبة</span>
                <div class="text-3xl font-extrabold text-emerald-900 mt-2" id="dashTotalQty">0</div>
            </div>
            <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 card-box border-r-8 border-r-amber-500">
                <span class="text-slate-500 font-semibold text-sm">المبلغ الإجمالي المطلـوب</span>
                <div class="text-3xl font-extrabold text-amber-600 mt-2" id="dashTotalAmount">0.00 ر.س</div>
            </div>
            <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 card-box border-r-8 border-r-blue-600">
                <span class="text-slate-500 font-semibold text-sm">عدد المدارس والأقسام</span>
                <div class="text-3xl font-extrabold text-blue-900 mt-2" id="dashSchoolsCount">19</div>
            </div>
        </div>

        <!-- Section 1: Prices Table -->
        <div class="bg-white rounded-2xl shadow-md border border-slate-200 p-6 card-box">
            <h2 class="text-xl font-bold text-emerald-900 mb-4">🏷️ قائمة أسعار الأصناف والملخص الإجمالي</h2>
            <div class="overflow-x-auto">
                <table class="w-full text-sm text-right border-collapse" id="pricesTable">
                    <thead>
                        <tr class="bg-emerald-900 text-white font-bold">
                            <th class="p-3 border border-emerald-800 text-center w-12">#</th>
                            <th class="p-3 border border-emerald-800">اسم الصنف</th>
                            <th class="p-3 border border-emerald-800 text-center w-36">سعر الوحدة (ر.س)</th>
                            <th class="p-3 border border-emerald-800 text-center w-36">إجمالي الكمية (حبة)</th>
                            <th class="p-3 border border-emerald-800 text-center w-40">المبلغ الإجمالي (ر.س)</th>
                        </tr>
                    </thead>
                    <tbody id="pricesTbody"></tbody>
                    <tfoot>
                        <tr class="bg-emerald-50 font-extrabold text-emerald-900 border-t-2 border-emerald-800">
                            <td colspan="3" class="p-3 border border-emerald-200 text-left pl-4">الإجمالي الكلي:</td>
                            <td id="summaryTotalQty" class="p-3 border border-emerald-200 text-center text-base">0</td>
                            <td id="summaryTotalAmt" class="p-3 border border-emerald-200 text-center text-lg text-emerald-800">0.00 ر.س</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>

        <!-- Section 2: Matrix Table -->
        <div class="bg-white rounded-2xl shadow-md border border-slate-200 p-6 card-box">
            <h2 class="text-xl font-bold text-emerald-900 mb-4">📊 جدول الطلبيات الشامل (الأصناف × المدارس)</h2>
            <div class="overflow-x-auto">
                <table class="w-full text-xs border-collapse" id="matrixTable">
                    <thead id="matrixThead" class="bg-slate-800 text-white"></thead>
                    <tbody id="matrixTbody" class="divide-y divide-slate-200 text-center"></tbody>
                    <tfoot id="matrixTfoot" class="bg-slate-100 font-bold border-t-2 border-slate-400"></tfoot>
                </table>
            </div>
        </div>
    </div>

    <script>
        const itemsPrices = {
            "زعتر": 1.25, "زعتر بالفرن": 2.50, "زعتر مدور": 2.50, "مناقيش زعتر": 2.50,
            "جبنة": 1.25, "جبن بالفرن": 5.75, "جبنة عكاوي": 7.00, "مناقيش جبن": 5.75,
            "لبنة": 1.25, "لبنة بالفرن": 5.50, "بيتزا": 1.25, "بيتزا مدور": 1.25,
            "سبانخ": 1.25, "سبانخ بالفرن": 4.60, "فلافل": 4.60, "كرواسون": 2.00,
            "كرواسون زبدة": 2.00, "كرواسون زعتر": 2.25, "دونات": 2.00, "دونات شوكولاته": 2.50,
            "كيك أرامكو": 1.50, "كيك أرامكو شوكولاته": 1.50, "سويت رول": 1.25, "سينابون": 3.00,
            "بوكس معجنات": 25.00, "ميني كب كيك": 15.00
        };

        const schoolsData = [
            { name: "مدرسة التعلم الأهلية (بنات)", items: {"زعتر": 5, "دونات": 5, "بيتزا": 5, "جبنة": 5, "لبنة": 5} },
            { name: "مدرسة رويال", items: {"جبنة": 5, "لبنة": 4, "زعتر": 5, "بيتزا": 4, "كرواسون زبدة": 3, "كيك أرامكو": 5} },
            { name: "التربية النموذجية (بنين)", items: {"لبنة بالفرن": 5, "جبن بالفرن": 7, "زعتر": 6, "بيتزا": 6, "سبانخ": 12, "جبنة": 10, "دونات شوكولاته": 4, "كيك أرامكو": 4} },
            { name: "مدارس التربية النموذجية (بنات)", items: {"زعتر بالفرن": 5, "جبن بالفرن": 8, "سبانخ": 2, "زعتر مدور": 4, "لبنة": 4, "بيتزا": 4, "سينابون": 4, "كيك أرامكو": 4, "جبنة": 4} },
            { name: "التعلم (بنين)", items: {"جبنة": 10, "لبنة": 10, "زعتر": 10, "بيتزا": 10, "دونات": 10, "سويت رول": 5} },
            { name: "الراكة", items: {"لبنة": 10, "بيتزا": 10, "جبنة": 10, "زعتر": 5, "بوكس معجنات": 1} },
            { name: "الظهران", items: {"كيك أرامكو": 10, "جبنة": 4, "لبنة": 3, "زعتر": 3} },
            { name: "ذوي الأوطان", items: {"جبن بالفرن": 5, "سبانخ بالفرن": 4, "زعتر بالفرن": 3, "فلافل": 6, "لبنة": 2, "جبنة": 3, "كرواسون": 4, "بيتزا": 3, "كيك أرامكو": 3, "كيك أرامكو شوكولاته": 3} },
            { name: "واحة الإبداع (بنات)", items: {"كرواسون زبدة": 4, "جبنة": 2, "زعتر": 2, "بيتزا": 2, "سبانخ": 1, "لبنة": 1} },
            { name: "واحة الإبداع (بنين)", items: {"كرواسون زبدة": 15, "بيتزا": 15, "جبنة": 10, "لبنة": 10} },
            { name: "مدارس السلام (بنات)", items: {"جبنة": 10, "لبنة": 10, "بيتزا": 8, "كرواسون زبدة": 40, "كيك أرامكو": 15, "دونات": 4, "سبانخ": 10, "سويت رول": 10} },
            { name: "السلام ابتدائي بنين (د1)", items: {"جبنة عكاوي": 6, "زعتر": 4, "لبنة": 4, "بيتزا": 3, "كرواسون زبدة": 4} },
            { name: "السلام متوسط (بنين)", items: {"جبنة عكاوي": 15, "لبنة": 5, "جبنة": 5, "كرواسون": 5, "زعتر": 5, "دونات": 2, "سويت رول": 3} },
            { name: "السلام ابتدائي (د2)", items: {"بيتزا": 4, "دونات": 3, "سويت رول": 3, "جبنة": 4, "لبنة": 3, "زعتر": 3, "جبن بالفرن": 1} },
            { name: "السلام ثانوية (بنين)", items: {"لبنة": 10, "زعتر": 10, "جبنة": 5, "دونات": 5, "كرواسون": 10, "جبنة عكاوي": 15} },
            { name: "الكفاح (بنين)", items: {"جبن بالفرن": 8, "زعتر": 2} },
            { name: "الكفاح (بنات)", items: {"مناقيش زعتر": 5, "مناقيش جبن": 4, "كرواسون زعتر": 4} },
            { name: "رواد الخليج - الزهور (بنات)", items: {"جبن بالفرن": 10, "زعتر بالفرن": 10, "سبانخ بالفرن": 2, "بيتزا مدور": 5, "بوكس معجنات": 1, "ميني كب كيك": 1} },
            { name: "الزهور (بنين)", items: {"جبن بالفرن": 20, "جبنة": 15, "لبنة": 15, "زعتر": 15, "بيتزا": 5, "كيك أرامكو": 5} }
        ];

        const itemNames = Object.keys(itemsPrices);

        function getItemTotalQty(item) {
            return schoolsData.reduce((sum, s) => sum + (s.items[item] || 0), 0);
        }

        function renderAll() {
            const pricesTbody = document.getElementById('pricesTbody');
            pricesTbody.innerHTML = '';
            let grandTotalQty = 0, grandTotalAmount = 0;

            itemNames.forEach((item, index) => {
                const qty = getItemTotalQty(item);
                const price = itemsPrices[item];
                const totalAmt = qty * price;
                grandTotalQty += qty;
                grandTotalAmount += totalAmt;

                pricesTbody.innerHTML += `
                    <tr class="${index % 2 === 0 ? 'bg-slate-50' : 'bg-white'} hover:bg-emerald-50">
                        <td class="p-3 border text-center font-bold text-slate-500">${index + 1}</td>
                        <td class="p-3 border font-bold text-slate-800">${item}</td>
                        <td class="p-3 border text-center">
                            <input type="number" step="0.25" value="${price}" onchange="updatePrice('${item}', this.value)" class="w-24 text-center border rounded p-1 font-bold text-emerald-800 no-print">
                            <span class="hidden print:inline font-bold">${price.toFixed(2)}</span>
                        </td>
                        <td class="p-3 border text-center font-bold text-slate-700">${qty}</td>
                        <td class="p-3 border text-center font-bold text-emerald-700">${totalAmt.toFixed(2)} ر.س</td>
                    </tr>`;
            });

            document.getElementById('summaryTotalQty').innerText = grandTotalQty.toLocaleString();
            document.getElementById('summaryTotalAmt').innerText = grandTotalAmount.toFixed(2) + ' ر.س';
            document.getElementById('dashTotalQty').innerText = grandTotalQty.toLocaleString();
            document.getElementById('dashTotalAmount').innerText = grandTotalAmount.toFixed(2) + ' ر.س';

            renderMatrix();
        }

        function updatePrice(item, newPrice) {
            const val = parseFloat(newPrice);
            if (!isNaN(val) && val >= 0) {
                itemsPrices[item] = val;
                renderAll();
            }
        }

        function renderMatrix() {
            const thead = document.getElementById('matrixThead');
            const tbody = document.getElementById('matrixTbody');
            const tfoot = document.getElementById('matrixTfoot');

            let headHTML = `<tr class="bg-emerald-900 text-white"><th class="p-2 border">#</th><th class="p-2 border text-right">الصنف / المدرسة</th>`;
            schoolsData.forEach(s => headHTML += `<th class="p-2 border text-center">${s.name}</th>`);
            headHTML += `<th class="p-2 border text-center bg-emerald-800">إجمالي الحبات</th><th class="p-2 border text-center bg-emerald-950">الإجمالي (ر.س)</th></tr>`;
            thead.innerHTML = headHTML;

            tbody.innerHTML = '';
            itemNames.forEach((item, idx) => {
                let rowQty = 0;
                let rowHTML = `<tr class="${idx % 2 === 0 ? 'bg-slate-50' : 'bg-white'}"><td class="p-2 border font-bold">${idx + 1}</td><td class="p-2 border font-bold text-right">${item}</td>`;
                schoolsData.forEach(s => {
                    const q = s.items[item] || 0;
                    rowQty += q;
                    rowHTML += `<td class="p-2 border ${q > 0 ? 'font-bold text-emerald-900 bg-emerald-50' : 'text-slate-300'}">${q > 0 ? q : '-'}</td>`;
                });
                const rowAmt = rowQty * itemsPrices[item];
                rowHTML += `<td class="p-2 border font-bold bg-slate-100">${rowQty}</td><td class="p-2 border font-bold text-emerald-800 bg-emerald-100">${rowAmt.toFixed(2)}</td></tr>`;
                tbody.innerHTML += rowHTML;
            });
        }

        function exportToExcel() {
            const wb = XLSX.utils.book_new();
            const pricesData = [["قائمة الأسعار والملخص العام"], [], ["#", "اسم الصنف", "سعر الوحدة (ر.س)", "إجمالي الكمية", "المبلغ الإجمالي"]];
            itemNames.forEach((item, i) => {
                const q = getItemTotalQty(item), p = itemsPrices[item];
                pricesData.push([i + 1, item, p, q, q * p]);
            });
            XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(pricesData), "الأسعار والملخص");
            XLSX.writeFile(wb, "طلبيات_المدارس_الشاملة.xlsx");
        }

        window.onload = renderAll;
    </script>
</body>
</html>