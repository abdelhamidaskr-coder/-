
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
                    <thead id="matrixThead" class="sticky top-0 bg-slate-900 text-white z-10"></thead>
                    <tbody id="matrixTbody" class="divide-y divide-slate-200"></tbody>
                    <tfoot id="matrixTfoot" class="sticky bottom-0 bg-slate-100 font-bold border-t-2 border-slate-400 z-10"></tfoot>
                </table>
            </div>
        </div>

        <!-- 2. جدول أسعار الأصناف والملخص المالي -->
        <div class="bg-white rounded-2xl shadow-md border border-slate-200 p-5">
            <div class="mb-3">
                <h2 class="text-lg font-bold text-emerald-900">🏷️ جدول أسعار الأصناف والملخص الإجمالي (حسب الصورة المرفقة)</h2>
            </div>
            <div class="overflow-x-auto">
                <table class="w-full text-xs text-right border-collapse">
                    <thead>
                        <tr class="bg-emerald-900 text-white font-bold">
                            <th class="p-2.5 border text-center w-12">#</th>
                            <th class="p-2.5 border">اسم الصنف</th>
                            <th class="p-2.5 border text-center w-36">سعر الوحدة (ر.س)</th>
                            <th class="p-2.5 border text-center w-36">إجمالي الكمية المطلوبة</th>
                            <th class="p-2.5 border text-center w-40">المبلغ الإجمالي (ر.س)</th>
                        </tr>
                    </thead>
                    <tbody id="pricesTbody" class="divide-y divide-slate-200"></tbody>
                    <tfoot>
                        <tr class="bg-emerald-100 font-extrabold text-emerald-950 text-sm">
                            <td colspan="3" class="p-3 border text-left pl-4">الإجمالي الكلي لجميع المدارس:</td>
                            <td id="summaryTotalQty" class="p-3 border text-center text-base text-slate-900">0</td>
                            <td id="summaryTotalAmt" class="p-3 border text-center text-lg text-emerald-900">0.00 ر.س</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>

    </div>

    <script>
        // الأصناف والأسعار الثابتة (مطابقة للصورة تماماً)
        const items = [
            { id: "i1", name: "زعتر", price: 1.25 },
            { id: "i2", name: "زعتر بالفرن", price: 2.50 },
            { id: "i3", name: "زعتر مدور", price: 2.50 },
            { id: "i4", name: "مناقيش زعتر", price: 2.50 },
            { id: "i5", name: "جبنة", price: 1.25 },
            { id: "i6", name: "جبن بالفرن", price: 5.75 },
            { id: "i7", name: "جبنة عكاوي", price: 7.00 },
            { id: "i8", name: "مناقيش جبن", price: 5.75 },
            { id: "i9", name: "لبنة", price: 1.25 },
            { id: "i10", name: "لبنة بالفرن", price: 5.50 },
            { id: "i11", name: "بيتزا", price: 1.25 },
            { id: "i12", name: "بيتزا مدور", price: 1.25 },
            { id: "i13", name: "سبانخ", price: 1.25 },
            { id: "i14", name: "سبانخ بالفرن", price: 4.60 },
            { id: "i15", name: "فلافل", price: 4.60 },
            { id: "i16", name: "كرواسون", price: 2.00 },
            { id: "i17", name: "كرواسون زبدة", price: 2.00 },
            { id: "i18", name: "كرواسون زعتر", price: 2.25 },
            { id: "i19", name: "دونات", price: 2.00 },
            { id: "i20", name: "دونات شوكولاته", price: 2.50 },
            { id: "i21", name: "كيك أرامكو", price: 1.50 },
            { id: "i22", name: "كيك أرامكو شوكولاته", price: 1.50 },
            { id: "i23", name: "سويت رول", price: 1.25 },
            { id: "i24", name: "سينابون", price: 3.00 },
            { id: "i25", name: "بوكس معجنات", price: 25.00 },
            { id: "i26", name: "ميني كب كيك", price: 15.00 }
        ];

        // المدارس الثابتة بطلبياتها الأولية
        const schools = [
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

        function getItemTotalQty(itemName) {
            return schools.reduce((sum, s) => sum + (s.items[itemName] || 0), 0);
        }

        function renderAll() {
            renderMatrix();
            renderPricesTable();
            updateDashboards();
        }

        function renderMatrix() {
            const thead = document.getElementById('matrixThead');
            const tbody = document.getElementById('matrixTbody');
            const tfoot = document.getElementById('matrixTfoot');

            let headHTML = `<tr><th class="p-2 border border-slate-700 w-8">#</th><th class="p-2 border border-slate-700 text-right min-w-[150px]">اسم الصنف / المدرسة</th>`;
            schools.forEach((s) => {
                headHTML += `<th class="p-2 border border-slate-700 min-w-[95px] text-xs font-bold">${s.name}</th>`;
            });
            headHTML += `<th class="p-2 border border-slate-700 w-24 bg-emerald-800 text-white">إجمالي الحبات</th><th class="p-2 border border-slate-700 w-28 bg-emerald-950 text-amber-300">الإجمالي (ر.س)</th></tr>`;
            thead.innerHTML = headHTML;

            tbody.innerHTML = '';
            items.forEach((item, idx) => {
                let rowQty = 0;
                let rowHTML = `<tr class="${idx % 2 === 0 ? 'bg-slate-50' : 'bg-white'}"><td class="p-2 border font-bold text-slate-400">${idx + 1}</td><td class="p-2 border font-bold text-right text-slate-800">${item.name}</td>`;
                
                schools.forEach((s, sIdx) => {
                    const q = s.items[item.name] || 0;
                    rowQty += q;
                    rowHTML += `<td class="p-1 border">
                        <input type="number" min="0" value="${q > 0 ? q : ''}" placeholder="-" 
                            onchange="updateQty(${sIdx}, '${item.name}', this.value)" 
                            class="qty-box ${q > 0 ? 'has-qty' : 'text-slate-300'}">
                    </td>`;
                });

                const rowAmt = rowQty * item.price;
                rowHTML += `<td class="p-2 border font-extrabold bg-slate-100">${rowQty}</td><td class="p-2 border font-extrabold text-emerald-800 bg-emerald-100">${rowAmt.toFixed(2)}</td></tr>`;
                tbody.innerHTML += rowHTML;
            });

            let footQtyHTML = `<tr class="bg-slate-200 font-bold text-slate-800"><td colspan="2" class="p-2.5 border text-right">مجموع حبات المدرسة:</td>`;
            let schoolTotalsAmt = new Array(schools.length).fill(0);

            schools.forEach((s, sIdx) => {
                let sqty = 0, samt = 0;
                items.forEach(i => {
                    const q = s.items[i.name] || 0;
                    sqty += q; samt += q * i.price;
                });
                schoolTotalsAmt[sIdx] = samt;
                footQtyHTML += `<td class="p-2 border text-center text-slate-900 font-extrabold">${sqty}</td>`;
            });

            let totalAllQty = 0, totalAllAmt = 0;
            schools.forEach(s => {
                items.forEach(i => {
                    const q = s.items[i.name] || 0;
                    totalAllQty += q;
                    totalAllAmt += q * i.price;
                });
            });

            footQtyHTML += `<td class="p-2 border text-center bg-slate-300 text-slate-950 font-extrabold">${totalAllQty}</td><td class="p-2 border text-center bg-emerald-200 text-emerald-950 font-extrabold">${totalAllAmt.toFixed(2)}</td></tr>`;
            
            let footAmtHTML = `<tr class="bg-emerald-950 text-white font-extrabold"><td colspan="2" class="p-2.5 border text-right">مجموع مبلغ المدرسة (ر.س):</td>`;
            schoolTotalsAmt.forEach(amt => { footAmtHTML += `<td class="p-2 border text-center text-amber-300">${amt.toFixed(2)}</td>`; });
            footAmtHTML += `<td class="p-2 border text-center bg-amber-500 text-slate-950">${totalAllQty}</td><td class="p-2 border text-center bg-amber-400 text-slate-950">${totalAllAmt.toFixed(2)}</td></tr>`;

            tfoot.innerHTML = footQtyHTML + footAmtHTML;
        }

        function renderPricesTable() {
            const tbody = document.getElementById('pricesTbody');
            tbody.innerHTML = '';
            let grandQty = 0, grandAmt = 0;

            items.forEach((item, idx) => {
                const qty = getItemTotalQty(item.name);
                const amt = qty * item.price;
                grandQty += qty; grandAmt += amt;

                tbody.innerHTML += `<tr class="${idx % 2 === 0 ? 'bg-slate-50' : 'bg-white'}">
                    <td class="p-2 border text-center font-bold text-slate-400">${idx + 1}</td>
                    <td class="p-2 border font-bold text-slate-800">${item.name}</td>
                    <td class="p-2 border text-center">
                        <input type="number" step="0.25" value="${item.price}" onchange="updatePrice('${item.id}', this.value)" class="w-24 text-center border rounded font-bold text-emerald-800 p-1 focus:bg-amber-100 focus:outline-none">
                    </td>
                    <td class="p-2 border text-center font-bold text-slate-700">${qty}</td>
                    <td class="p-2 border text-center font-bold text-emerald-700">${amt.toFixed(2)} ر.س</td>
                </tr>`;
            });

            document.getElementById('summaryTotalQty').innerText = grandQty.toLocaleString();
            document.getElementById('summaryTotalAmt').innerText = grandAmt.toFixed(2) + ' ر.س';
        }

        function updateDashboards() {
            let totalQty = 0, totalAmt = 0;
            items.forEach(i => {
                const q = getItemTotalQty(i.name);
                totalQty += q; totalAmt += q * i.price;
            });
            document.getElementById('dashTotalQty').innerText = totalQty.toLocaleString();
            document.getElementById('dashTotalAmount').innerText = totalAmt.toFixed(2) + ' ر.س';
        }

        function updateQty(schoolIndex, itemName, val) {
            const num = parseInt(val) || 0;
            if (num > 0) schools[schoolIndex].items[itemName] = num;
            else delete schools[schoolIndex].items[itemName];
            renderAll();
        }

        function updatePrice(itemId, val) {
            const item = items.find(i => i.id === itemId);
            const p = parseFloat(val);
            if (item && !isNaN(p) && p >= 0) {
                item.price = p;
                renderAll();
            }
        }

        function exportToExcel() {
            const wb = XLSX.utils.book_new();
            
            const pricesData = [["قائمة الأسعار والملخص العام"], [], ["#", "اسم الصنف", "سعر الوحدة (ر.س)", "إجمالي الكمية", "الإجمالي المالي (ر.س)"]];
            let gQty = 0, gAmt = 0;
            items.forEach((i, idx) => {
                const q = getItemTotalQty(i.name), a = q * i.price;
                gQty += q; gAmt += a;
                pricesData.push([idx + 1, i.name, i.price, q, a]);
            });
            pricesData.push([], ["", "الإجمالي الكلي", "", gQty, gAmt]);
            XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(pricesData), "الأسعار والملخص");

            const matrixHead = ["#", "اسم الصنف"];
            schools.forEach(s => matrixHead.push(s.name));
            matrixHead.push("إجمالي الحبات", "الإجمالي (ر.س)");

            const matrixData = [["جدول الطلبيات الشامل (المدارس × الأصناف)"], [], matrixHead];
            items.forEach((i, idx) => {
                const row = [idx + 1, i.name];
                let rq = 0;
                schools.forEach(s => {
                    const q = s.items[i.name] || 0;
                    rq += q; row.push(q > 0 ? q : "");
                });
                row.push(rq, rq * i.price);
                matrixData.push(row);
            });

            XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(matrixData), "جدول الطلبيات الشامل");
            XLSX.writeFile(wb, "طلبيات_المدارس.xlsx");
        }

        window.onload = renderAll;
    </script>
</body>
</html>