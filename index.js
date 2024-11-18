const TelegramBot = require('node-telegram-bot-api');
const ExcelJS = require('exceljs'); // استيراد مكتبة exceljs
require('dotenv').config(); // إذا كنت تستخدم متغيرات بيئية
const express = require('express'); // إضافة Express لتشغيل السيرفر

// إعداد سيرفر Express (لتشغيل التطبيق على Render أو في بيئة محلية)
const app = express();
const port = process.env.PORT || 4000; // المنفذ الافتراضي

// استبدل بالتوكن الخاص بك
const token = process.env.TELEGRAM_BOT_TOKEN || '7203035834:AAEaT5eaKIKYnbD7jtlEijifCr7z7t1ZBL0';

// إنشاء البوت
const bot = new TelegramBot(token, { polling: true });

// تخزين البيانات من Excel
let data = {};

// دالة لتحميل البيانات من Excel
async function loadDataFromExcel() {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('gas18-11-2024.xlsx'); // اسم الملف
        const worksheet = workbook.worksheets[0]; // أول ورقة عمل

        worksheet.eachRow((row, rowNumber) => {
            // قراءة القيم من الصفوف
            const idNumber = row.getCell(1).value?.toString().trim(); // رقم الهوية
            const name = row.getCell(2).value?.toString().trim(); // اسم الطالب
            const phoneNumber = row.getCell(3).value?.toString().trim(); // رقم الجوال
            const province = row.getCell(4).value?.toString().trim(); // المحافظة
            const city = row.getCell(5).value?.toString().trim(); // المدينة
            const area = row.getCell(6).value?.toString().trim(); // الحي/المنطقة
            const distributorId = row.getCell(7).value?.toString().trim(); // هوية الموزع
            const distributorName = row.getCell(8).value?.toString().trim(); // اسم الموزع
            const distributorPhone = row.getCell(9).value?.toString().trim(); // رقم الموزع
            const status = row.getCell(10).value?.toString().trim(); // الحالة
            const orderDate = row.getCell(11).value?.toString().trim(); // تاريخ الطلب
            const district = row.getCell(12).value?.toString().trim();

            if (idNumber && name) {
                data[idNumber] = {
                    name: name || "غير متوفر",
                    phoneNumber: phoneNumber || "غير متوفر",
                    province: province || "غير متوفر",
                    city: city || "غير متوفر",
                    area: area || "غير متوفر",
                    distributorId: distributorId || "غير متوفر",
                    distributorName: distributorName || "غير متوفر",
                    distributorPhone: distributorPhone || "غير متوفر",
                    status: status || "غير متوفر",
                    orderDate: orderDate || "غير متوفر",
                    district: district || "غير متوفر",
                };
            }
        });

        console.log('تم تحميل البيانات بنجاح.');
    } catch (error) {
        console.error('حدث خطأ أثناء قراءة ملف Excel:', error.message);
    }
}

// تحميل البيانات عند بدء التشغيل
loadDataFromExcel();

// الرد على أوامر البوت
bot.onText(/\/start/, (msg) => {
    bot.sendMessage(msg.chat.id, "مرحبًا! أدخل رقم الهوية للحصول على التفاصيل.");
});

bot.on('message', (msg) => {
    const chatId = msg.chat.id;
    const idNumber = msg.text.trim(); // رقم الهوية

    if (idNumber === '/start') return;

    const user = data[idNumber];
    if (user) {
        const response = `
الاسم: ${user.name}
المحافظة: ${user.province}
المدينة: ${user.city}
الحي / المنطقة: ${user.area}
هوية الموزع: ${user.distributorId}
اسم الموزع: ${user.distributorName}
رقم جوال الموزع: ${user.distributorPhone}
الحالة: ${user.status}
تاريخ الطلب: ${user.orderDate}
        `;
        bot.sendMessage(chatId, response);
    } else {
        bot.sendMessage(chatId, "عذرًا، لم أتمكن من العثور على بيانات لرقم الهوية المدخل.");
    }
});

// تشغيل السيرفر
app.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});
