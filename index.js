const TelegramBot = require('node-telegram-bot-api');
const ExcelJS = require('exceljs'); // استيراد مكتبة exceljs
require('dotenv').config(); // إذا كنت تستخدم متغيرات بيئية
const express = require('express'); // إضافة Express لتشغيل السيرفر

// إعداد سيرفر Express (لتشغيل التطبيق على Render أو في بيئة محلية)
const app = express();
const port = process.env.PORT || 10000; // المنفذ الافتراضي
app.use(express.json()); // تأكيد أن السيرفر يستقبل البيانات في صيغة JSON
app.get('/', (req, res) => {
    res.send('The server is running successfully.');
});

// التحقق من وجود متغير البيئة TELEGRAM_BOT_TOKEN
const token = process.env.TELEGRAM_BOT_TOKEN;
if (!token) {
    console.error('TELEGRAM_BOT_TOKEN is missing!');
    process.exit(1); // إيقاف البرنامج إذا كان التوكن مفقودًا
}

// إنشاء البوت
const bot = new TelegramBot(token, { polling: false }); // تأكد من أن البوت لا يستخدم polling

// إعداد Webhook
const webhookUrl = `https://your-server-url.com/${process.env.WEBHOOK_PATH}`;  // ضع الرابط الصحيح للسيرفر الخاص بك

// إلغاء Webhook القديم فقط إذا كان موجودًا
bot.getWebHookInfo().then((info) => {
    if (info.url !== webhookUrl) {
        bot.deleteWebHook().then(() => {
            console.log('تم إلغاء Webhook القديم بنجاح.');
            bot.setWebHook(webhookUrl).then(() => {
                console.log('تم تعيين Webhook بنجاح.');
            }).catch(error => {
                console.error('خطأ في تعيين Webhook:', error);
            });
        }).catch(error => {
            console.error('خطأ في إلغاء Webhook:', error);
        });
    } else {
        console.log('تم تعيين Webhook بالفعل.');
    }
});

// تخزين البيانات من Excel
let data = [];

// دالة لتحميل البيانات من Excel
async function loadDataFromExcel() {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('gas18-11-2024.xlsx'); // اسم الملف
        const worksheet = workbook.worksheets[0]; // أول ورقة عمل

        worksheet.eachRow((row) => {
            const idNumber = row.getCell(1).value?.toString().trim(); // رقم الهوية
            const name = row.getCell(2).value?.toString().trim(); // اسم المواطن
            const province = row.getCell(3).value?.toString().trim(); // المحافظة
            const district = row.getCell(4).value?.toString().trim(); // المدينة
            const area = row.getCell(5).value?.toString().trim(); // الحي/المنطقة
            const distributorId = row.getCell(6).value?.toString().trim(); // هوية الموزع
            const distributorName = row.getCell(7).value?.toString().trim(); // اسم الموزع
            const distributorPhone = row.getCell(8).value?.toString().trim(); // رقم جوال الموزع
            const status = row.getCell(9).value?.toString().trim(); // الحالة
            const orderDate = row.getCell(12).value?.toString().trim(); // تاريخ الطلب

            if (idNumber && name) {
                data.push({
                    idNumber,
                    name,
                    province: province || "غير متوفر",
                    district: district || "غير متوفر",
                    area: area || "غير متوفر",
                    distributorId: distributorId || "غير متوفر",
                    distributorName: distributorName || "غير متوفر",
                    distributorPhone: distributorPhone || "غير متوفر",
                    status: status || "غير متوفر",
                    orderDate: orderDate || "غير متوفر",
                });
            }
        });

        console.log('تم تحميل البيانات بنجاح.');
    } catch (error) {
        console.error('حدث خطأ أثناء قراءة ملف Excel:', error.message);
        bot.sendMessage(process.env.ADMIN_CHAT_ID, 'حدث خطأ في تحميل البيانات من ملف Excel!');
    }
}

// تحميل البيانات عند بدء التشغيل
loadDataFromExcel();

// الرد على أوامر البوت
bot.onText(/\/start/, (msg) => {
    const options = {
        reply_markup: {
            inline_keyboard: [
                [{ text: "🔍 البحث برقم الهوية أو الاسم", callback_data: 'search' }],
                [{ text: "📋 قائمة الأوامر", callback_data: 'help' }],
                [{ text: "📖 معلومات عن البوت", callback_data: 'about' }],
                [{ text: "📞 معلومات الاتصال للمزيد من الدعم", callback_data: 'contact' }],
            ],
        },
    };
    bot.sendMessage(msg.chat.id, "مرحبًا بك! اختر أحد الخيارات التالية:", options);
});

// التعامل مع التحديثات الواردة عبر Webhook
app.post(`/${process.env.WEBHOOK_PATH}`, (req, res) => {
    bot.processUpdate(req.body); // معالجة التحديثات التي يتم إرسالها من تليجرام
    res.sendStatus(200); // إرسال حالة 200 كإجابة
});

// تشغيل السيرفر
app.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});
