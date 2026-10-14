# ติดตั้งฐานข้อมูลกลางบน Cloudflare

ระบบประกอบด้วย Cloudflare **Worker** (API + เว็บแอป) และฐานข้อมูล **D1** ทั้งหมดอยู่ที่อยู่เดียว เช่น `https://luckswan-sales.<ชื่อบัญชี>.workers.dev`

## วิธีที่ 1 — ให้ GitHub deploy ให้อัตโนมัติ (แนะนำ)
ทำครั้งเดียว หลังจากนั้นทุกครั้งที่ merge เข้า `main` ระบบจะอัปเดตเอง

1. สมัคร/เข้าสู่ระบบ https://dash.cloudflare.com แล้วเปิดเมนู **Workers & Pages** หนึ่งครั้ง (เพื่อให้ได้ชื่อ subdomain `*.workers.dev`)
2. คัดลอก **Account ID** — ตัวอักษร 0-9 a-f ยาว 32 ตัว อยู่ที่ Workers & Pages → แถบขวา “Account details” (ไม่ใช่ Zone ID, ชื่อบัญชี หรืออีเมล)
3. สร้าง API Token: My Profile → API Tokens → **Create Token** → เลือกแม่แบบ **Edit Cloudflare Workers** → กด *Add more* เพิ่มสิทธิ์ **Account · D1 · Edit** → Create แล้วคัดลอก token
4. ใน GitHub: repo `ls_it_support` → Settings → Secrets and variables → Actions → **New repository secret** สร้าง 2 ตัว
   - `CLOUDFLARE_API_TOKEN` = token จากข้อ 3
   - `CLOUDFLARE_ACCOUNT_ID` = Account ID จากข้อ 2
5. GitHub → แท็บ **Actions** → *Deploy to Cloudflare* → **Run workflow** (หรือ merge PR เข้า main)
6. เมื่อเสร็จ ดูที่อยู่เว็บในขั้น *Deploy worker and web app* (บรรทัดที่ลงท้าย `.workers.dev`)
7. เปิดที่อยู่นั้น → หน้า **ตั้งค่าครั้งแรก** → สร้างบัญชีผู้ดูแลระบบ → เพิ่มผู้ใช้งานจากหน้า “ผู้ใช้งาน”

## วิธีที่ 2 — deploy จากคอมพิวเตอร์
ต้องมี Node.js 22 ขึ้นไป
```bash
cd worker
npm install
npx wrangler login                     # เปิดเบราว์เซอร์ให้ล็อกอิน Cloudflare
npx wrangler d1 create luckswan-sales  # คัดลอก database_id ที่ได้ไปใส่ใน wrangler.toml
npx wrangler d1 migrations apply luckswan-sales --remote
npx wrangler deploy
```

## ย้ายข้อมูลเดิมขึ้นฐานข้อมูลกลาง
1. ในแอปเดิม (เครื่องที่มีข้อมูล) เมนูรูปโปรไฟล์ → **สำรองข้อมูล (ดาวน์โหลด)**
2. ในแอปใหม่ ล็อกอินเป็นผู้ดูแลระบบ → เมนูรูปโปรไฟล์ → **นำเข้าข้อมูลสำรองขึ้นคลาวด์** → เลือกไฟล์
   - ผู้ใช้จากไฟล์สำรองจะถูกสร้างพร้อมรหัสผ่านเริ่มต้น `luckswan` — ให้ทุกคนเปลี่ยนรหัสหลังเข้าใช้ครั้งแรก

## ระบบอนุมัติใบฟังก์ชัน
- ลำดับตำแหน่ง: พนักงานขาย / บัญชี < หัวหน้าแผนก < ผู้จัดการฝ่ายขาย < ผู้ดูแลระบบ
- ผู้จัดการฝ่ายขายและผู้ดูแลระบบ: แก้แล้วมีผลทันที
- ตำแหน่งอื่น: สร้าง/แก้/นำเข้ากำหนดการแล้วสถานะเป็น **รออนุมัติ** ส่งให้แผนกไม่ได้จนกว่าผู้มีตำแหน่งสูงกว่าจะอนุมัติ
- ผู้อนุมัติเห็นรายการในหน้า “หน้าหลัก” กด **ดูการเปลี่ยนแปลง** เทียบกับฉบับที่อนุมัติล่าสุด แล้ว **อนุมัติ** หรือ **ตีกลับ** (ต้องระบุเหตุผล)
- กฎเหล่านี้ตรวจที่เซิร์ฟเวอร์ ไม่สามารถข้ามได้จากหน้าเว็บ ทุกการอนุมัติ/ตีกลับถูกบันทึกในตาราง `audit`
- บัญชีที่มีสิทธิ์เฉพาะ “ดู” (เช่น บัญชี) แก้ใบฟังก์ชันไม่ได้

## ทดสอบบนเครื่อง
```bash
cd worker && npm install
npx wrangler d1 migrations apply luckswan-sales --local
npx wrangler dev --local --port 8787   # เปิด http://localhost:8787
```
