<!DOCTYPE html>
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