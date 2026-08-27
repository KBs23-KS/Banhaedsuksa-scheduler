# BHS School Scheduler — Phase 6 Deploy Package

แพ็กเกจนี้ใช้โค้ด Phase 6 เดิมเป็น `src/App.jsx` และคง Firebase เดิมไว้ในโค้ด

## Firebase

ไม่ต้องนำตัวเว็บไปลง Firebase Hosting ถ้าจะใช้ Vercel หรือ Netlify

เว็บที่ Deploy จะยังเชื่อม Firebase โปรเจกต์เดิม:

- Project ID: `school-scheduler-bhs`
- Firebase Authentication: Anonymous Auth ตามโค้ดเดิม
- Firestore: ใช้ฐานข้อมูลเดิม

ดังนั้นถ้าระบบเดิมเชื่อม Firebase ได้อยู่แล้ว ไม่ต้องเปลี่ยน Firebase config ในรอบนี้

## รันในเครื่อง

ต้องมี Node.js แล้วใช้คำสั่ง:

```bash
npm install
npm run dev
```

เปิด URL ที่ Vite แสดงใน Terminal

## Build

```bash
npm install
npm run build
```

ไฟล์พร้อมใช้งานจะอยู่ในโฟลเดอร์ `dist`

## Deploy Vercel

1. อัปโฟลเดอร์นี้ขึ้น GitHub repository
2. เข้า Vercel แล้วเลือก Add New Project
3. Import repository
4. Vercel จะตรวจพบ Vite
5. Build command: `npm run build`
6. Output directory: `dist`
7. กด Deploy

ไฟล์ `vercel.json` ถูกเตรียมไว้แล้ว

## Deploy Netlify

วิธีแนะนำ:

1. อัปโฟลเดอร์นี้ขึ้น GitHub repository
2. เข้า Netlify > Add new site > Import an existing project
3. เลือก repository
4. Build command: `npm run build`
5. Publish directory: `dist`
6. กด Deploy

ไฟล์ `netlify.toml` ถูกเตรียมไว้แล้ว

## ไฟล์สำคัญ

- `src/App.jsx` — โค้ด BHS Scheduler Phase 6 ทั้งหมด
- `src/main.jsx` — จุดเริ่ม React
- `src/index.css` — Tailwind
- `package.json` — dependencies/scripts
- `tailwind.config.js` — Tailwind + animation plugin
- `vite.config.js` — Vite
- `vercel.json` — Vercel config
- `netlify.toml` — Netlify config

## หมายเหตุ

แพ็กเกจนี้ไม่ได้เปลี่ยน layout หรือ className ภายใน `App.jsx`; เป็นการห่อ Phase 6 เดิมให้รันเป็นเว็บ Vite สำหรับ Deploy เท่านั้น
