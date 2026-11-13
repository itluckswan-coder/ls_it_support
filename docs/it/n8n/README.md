# แจ้งเตือนใบฟังก์ชัน · ใบเสนอราคา เข้า LINE ผ่าน n8n

LUX ส่งเหตุการณ์ของใบฟังก์ชันและใบเสนอราคาไปที่ n8n แล้ว n8n จัดข้อความภาษาไทยส่งเข้ากลุ่ม LINE

```
LUX (Cloudflare) ──เหตุการณ์ + รหัสลับ──► n8n (luckswanswan.app.n8n.cloud) ──► LINE Messaging API ──► กลุ่ม LINE
```

## เหตุการณ์ที่ส่ง

| event | เมื่อไหร่ |
|---|---|
| `created` | สร้างใบฟังก์ชันใหม่ |
| `sent` | กด "ส่งให้ … แผนก" (ถ้า `beo.rev` > 0 คือส่งฉบับแก้ไข) |
| `pending` | มีการแก้ไขที่รออนุมัติ |
| `approved` / `rejected` | อนุมัติ / ตีกลับ (มีเหตุผลใน `reason`) |
| `ack` | แผนกกดรับทราบ (`acked` = แผนกที่เพิ่งรับทราบ, `waiting` = แผนกที่ยังไม่รับทราบ) |
| `quote_pending` | ใบเสนอราคา: พนักงานขายกด "ส่งขออนุมัติ" (`kind: 'quote'`) |
| `quote_approved` | ใบเสนอราคา: ผู้ดูแลระบบกด "อนุมัติใบเสนอราคา" |
| `quote_rejected` | ใบเสนอราคา: ผู้ดูแลระบบกด "ตีกลับ" (มีเหตุผลใน `reason`) |

ข้อมูลที่ส่ง: เลขใบ, Function No. (ใบฟังก์ชัน), ชื่อพนักงานขาย (ใบเสนอราคา), ฉบับแก้ไข, ชื่องาน, ชื่อลูกค้า (บริษัท), วันงาน, จำนวนคน, รายชื่อแผนก, ผู้ทำรายการ, ลิงก์เปิดใบ
**ไม่ส่ง** ราคา ยอดเงิน เบอร์โทร หรืออีเมลลูกค้า

ถ้ายังไม่ได้ตั้ง secret ทั้งสองตัว (ข้อ 3) LUX ทำงานตามปกติและไม่ส่งอะไร

ตอนนี้เปิดไว้เฉพาะขั้นอนุมัติ (รออนุมัติ · อนุมัติ · ตีกลับ) ทั้งใบฟังก์ชัน (`pending` · `approved` · `rejected`) และใบเสนอราคา (`quote_pending` · `quote_approved` · `quote_rejected`) — เปิดตัวอื่นเพิ่มได้ที่ `SEND` ในกล่อง จัดข้อความ

โค้ดของกล่อง จัดข้อความ อยู่ในไฟล์ `code-node.js` ด้วย (เหมือนในไฟล์ workflow) — เวลาอัปเดตโค้ด เปิดกล่องนั้น ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดจากไฟล์นี้ ไม่ต้องนำเข้า workflow ใหม่

## ตั้งค่า (ทำครั้งเดียว)

### 1. Credentials ใน n8n
n8n › **Overview › Credentials › Create credential › Header Auth** สร้าง 2 ตัว

| Credential name | Name | Value |
|---|---|---|
| `LINE Messaging API` | `Authorization` | `Bearer ` + Channel access token (จาก developers.line.biz) |
| `LUX secret` | `X-LUX-Secret` | รหัสลับที่ตั้งเอง (สุ่มยาว 32 ตัวขึ้นไป เช่น จากโปรแกรมสร้างรหัสผ่าน) |

ตั้งรหัสลับเป็นตัวอักษรภาษาอังกฤษกับตัวเลข ไม่มีช่องว่าง (ระบบ LUX ตัดช่องว่างหน้า/หลังออกให้ แต่ฝั่ง n8n ไม่ตัด)

ห้ามส่ง token หรือรหัสลับในแชท อีเมล หรือใส่ในไฟล์ใน repo นี้ (repo เป็นสาธารณะ)

### 2. นำเข้า workflow
1. n8n › **Create workflow** › เมนู **⋯ › Import from File** › เลือก `docs/it/n8n/lux-function-line.json`
2. เปิดกล่อง **รับเหตุการณ์จาก LUX** › Credential เลือก `LUX secret`
3. เปิดกล่อง **ส่งเข้ากลุ่ม LINE** › Credential เลือก `LINE Messaging API`
4. กล่อง **จัดข้อความ**: บรรทัดบนสุด `SEND` เปิด/ปิดแต่ละเหตุการณ์ (`true` / `false`) · `GROUP_ID` คือกลุ่ม LINE ที่จะส่งเข้า
5. **Save** แล้ว **Publish / Active**
6. เปิดกล่อง **รับเหตุการณ์จาก LUX** › แท็บ **Production URL** › คัดลอก (ควรเป็น `https://luckswanswan.app.n8n.cloud/webhook/lux-function`)

### 3. บอก LUX ว่าจะส่งไปที่ไหน (GitHub Secrets)
GitHub › repo `ls_it_support` › **Settings › Secrets and variables › Actions › New repository secret**
(ลิงก์ตรง: https://github.com/itluckswan-coder/ls_it_support/settings/secrets/actions/new)

| Name | Secret |
|---|---|
| `N8N_WEBHOOK_URL` | Production URL จากข้อ 2.6 |
| `N8N_SECRET` | รหัสลับเดียวกับ credential `LUX secret` |

จากนั้น Actions › **Deploy to Cloudflare** › **Run workflow** (ระบบจะนำค่าไปตั้งใน Cloudflare ให้ · ดูขั้น "Set the n8n notification secrets" ต้องขึ้น `n8n secrets set`)

### 4. ทดสอบ
สร้าง / ส่ง / รับทราบใบฟังก์ชันสักใบใน LUX › ข้อความควรเด้งในกลุ่ม LINE ภายในไม่กี่วินาที
**ทางลัด:** ผู้ดูแลระบบกดชื่อตัวเองมุมขวาบนของ LUX › **ทดสอบแจ้งเตือน LINE** ระบบจะส่งข้อความ 🔔 ทดสอบ และบอกผลจาก n8n เป็นภาษาไทย
- ✅ รับแล้ว (200) แต่กลุ่มไม่ได้ข้อความ → ดู n8n › workflow › **Executions** ว่ากล่องไหนสีแดง (กล่อง ส่งเข้ากลุ่ม LINE: 401 = token ผิด ต้องมี `Bearer ` นำหน้า · 400 = Group ID ผิดหรือบอทไม่อยู่ในกลุ่ม)
- ❌ 403 → รหัสลับไม่ตรง: Value ของ credential `LUX secret` ต้องตรงกับ `N8N_SECRET` และช่อง Name ต้องเป็น `X-LUX-Secret`
- ❌ 404 → workflow ยังไม่ Publish หรือ `N8N_WEBHOOK_URL` ไม่ตรงกับ Production URL

**สำคัญ:** ถ้ารหัสลับไม่ตรง (403) หรือลิงก์ผิด (404) n8n **ไม่บันทึกอะไรใน Executions เลย** หน้า Executions ที่ว่างเปล่าจึงมักแปลว่าเป็นสองกรณีนี้ ใช้ปุ่มทดสอบเพื่อดูว่าเป็นกรณีไหน

## เปลี่ยนแปลงภายหลัง
- **เลือกเหตุการณ์ที่จะแจ้ง:** แก้ `SEND` ในกล่อง จัดข้อความ
- **ย้ายไปกลุ่มอื่น:** แก้ `GROUP_ID` (เอา Group ID จาก workflow "เก็บ Group ID")
- **แยกกลุ่มตามแผนก:** ทำได้ในกล่องจัดข้อความ (ส่ง `to` ตาม `waiting`/`depts`) — ขอให้ผู้พัฒนาช่วยได้
- **ย้าย n8n ไปเครื่องอื่น (เช่น NAS):** นำเข้าไฟล์เดิม แล้วเปลี่ยน `N8N_WEBHOOK_URL` ใน GitHub Secrets และรัน Deploy ใหม่
- **ปิดการแจ้งเตือนทั้งหมด:** ปิด workflow ใน n8n หรือลบ `N8N_WEBHOOK_URL` ออกจาก GitHub Secrets แล้วรัน Deploy
