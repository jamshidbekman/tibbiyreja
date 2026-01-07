# Deployment Guide (Serverga O'rnatish Qo'llanmasi)

Ushbu dastur Node.js muhitida ishlashga mo'ljallangan. Serverga o'rnatish va IP orqali kirish uchun quyidagi qadamlarni bajaring.

## 1. Talablar
- **Node.js**: Serverda Node.js (v16 yoki yuqori) o'rnatilgan bo'lishi kerak.
- **Git**: Kodni yuklab olish uchun.

## 2. O'rnatish

Serveringiz terminalida quyidagi buyruqlarni bajaring:

```bash
# 1. Kodni yuklab olish
git clone https://github.com/jamshidbekman/tibbiyreja.git
cd tibbiyreja

# 2. Kutubxonalarni o'rnatish
npm install
```

## 3. Ishga Tushirish

### Oddiy usul (Sinov uchun)
```bash
npm start
```
Bu dasturni `3000` portda ishga tushiradi.
Brauzerda `http://SERVER_IP:3000` manziliga kirsangiz, dastur ochiladi.

### Doimiy ishlash uchun (PM2 orqali)
Serveringiz o'chib yonganda ham dastur avtomatik ishlashi uchun `pm2` dan foydalanish tavsiya etiladi.

```bash
# PM2 o'rnatish (agar yo'q bo'lsa)
npm install -g pm2

# Dasturni ishga tushirish
pm2 start npm --name "reja" -- start

# Ro'yxatni saqlash
pm2 save
```

## 4. Portni O'zgartirish (Ixtiyoriy)
Agar dasturni boshqa portda (masalan 8080) ishlatmoqchi bo'lsangiz:

```bash
PORT=8080 pm2 start npm --name "reja" -- start
```

## 5. Tekshirish
Brauzeringizni oching va serveringiz IP manzili va portini kiriting:
`http://SIZNING_SERVER_IP:3000`

Dastur ochilishi kerak.
