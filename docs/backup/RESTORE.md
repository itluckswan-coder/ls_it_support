# สำรองข้อมูลและนำกลับมาใช้ — แอปฝ่ายขาย ลักษวรรณ รีสอร์ต

## ไฟล์สำรองมีอะไรบ้าง
เปิดไฟล์ `luckswan-backup-<วันที่>.7z` ด้วย 7-Zip (Windows) หรือ Keka / The Unarchiver (Mac) แล้วใส่รหัส `BACKUP_PASSWORD`

| ไฟล์ | คืออะไร |
|---|---|
| `database.sql` | ข้อมูลทั้งหมดในระบบ: ใบเสนอราคา ใบฟังก์ชัน รอบพนักงาน ตารางการทำงาน พนักงาน ผู้ใช้ (รหัสผ่านแบบเข้ารหัส) ข้อมูลบริษัท รูปที่อัปโหลด |
| `ls_it_support.bundle` | โค้ดทั้งหมด พร้อมประวัติการแก้ไขทุกครั้ง |
| `source-code.zip` | โค้ดล่าสุด (ไม่มีประวัติ) — เปิดดูได้ง่าย |
| `BACKUP-INFO.txt` | วันเวลาที่สำรอง และเลข commit |

> ไฟล์นี้มีข้อมูลลูกค้า เก็บในที่ปลอดภัย และอย่าส่งรหัสผ่านไปพร้อมไฟล์

## สำรองข้อมูล (ทำเมื่อต้องการ)
1. GitHub → repo `ls_it_support` → **Actions** → **Backup data** → **Run workflow**
2. รอประมาณ 1–2 นาที แล้วเปิดรายการที่รันเสร็จ (เครื่องหมายถูกสีเขียว)
3. ที่หัวข้อ **Artifacts** กดดาวน์โหลด `luckswan-backup-<วันที่>` (ได้ไฟล์ .zip ที่มีไฟล์ .7z อยู่ข้างใน)
4. เก็บไฟล์ไว้ในที่ปลอดภัย เช่น Google Drive หรือ NAS — GitHub เก็บไว้ให้แค่ 90 วัน

ต้องตั้ง secret `BACKUP_PASSWORD` (อย่างน้อย 12 ตัวอักษร) ครั้งแรกครั้งเดียว: Settings → Secrets and variables → Actions → New repository secret

## นำกลับมาใช้ในที่ใหม่ (บัญชี GitHub / Cloudflare ใหม่)
ต้องมี Node.js 22 ขึ้นไป แตกไฟล์ .7z ไว้ในโฟลเดอร์หนึ่ง แล้วทำตามลำดับนี้

1. **โค้ด** — สร้าง repo ใหม่ (ว่าง) บน GitHub แล้ว
   ```
   git clone ls_it_support.bundle ls_it_support
   cd ls_it_support
   git remote set-url origin https://github.com/<บัญชีใหม่>/<repo ใหม่>.git
   git push -u origin main
   ```
2. **ฐานข้อมูล — ต้องทำก่อน deploy ครั้งแรก**
   ```
   cd worker
   npm install
   npx wrangler login
   npx wrangler d1 create luckswan-sales
   npx wrangler d1 execute luckswan-sales --remote --file=../../database.sql
   ```
   (แก้ path ของ `database.sql` ให้ตรงกับที่แตกไฟล์ไว้)
3. **Deploy** — ใน repo ใหม่ใส่ secrets `CLOUDFLARE_API_TOKEN` และ `CLOUDFLARE_ACCOUNT_ID` ของบัญชี Cloudflare ใหม่ (ดู `worker/DEPLOY.md`) แล้วไปที่ Actions → **Deploy to Cloudflare** → Run workflow
   ระบบจะใช้ฐานข้อมูลที่นำเข้าไว้ และเข้าสู่ระบบด้วยผู้ใช้/รหัสผ่านเดิมได้ทันที
4. ตั้ง `BACKUP_PASSWORD` ใน repo ใหม่ด้วย เพื่อสำรองข้อมูลต่อได้

> ถ้า deploy ไปก่อนแล้ว (ฐานข้อมูลว่างถูกสร้างไปแล้ว): ลบฐานข้อมูลนั้นด้วย `npx wrangler d1 delete luckswan-sales` แล้วทำข้อ 2 ใหม่ จากนั้นรัน deploy อีกครั้ง

## นำข้อมูลกลับเข้าระบบเดิม (กู้คืนทับ)
ทำเฉพาะเมื่อข้อมูลเสียหายจริง และสำรองข้อมูลปัจจุบันไว้ก่อนเสมอ — ขอให้ Claude ช่วยทีละขั้นได้
