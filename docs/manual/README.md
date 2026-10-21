# คู่มือการใช้งานระบบงานขาย (ไทย · English · မြန်မာ)

ต้นฉบับของเว็บคู่มือ: คู่มือฉบับสมบูรณ์ + คู่มือย่อตามแผนก ทั้ง 3 ภาษา พร้อมรูปหน้าจอจริง

| ไฟล์ | ใช้ทำอะไร |
| --- | --- |
| `content.th.js` | เนื้อหาภาษาไทย (ฉบับอ้างอิงหลัก): หมวด (`chapters`), คู่มือย่อตามแผนก (`quick`), ประวัติการปรับปรุง (`changelog`) |
| `content.en.js`, `content.my.js` | คำแปลอังกฤษและพม่า — โครงสร้างและ `id` ต้องตรงกับฉบับไทยทุกข้อ |
| `index.html` | หน้าเว็บคู่มือ (เปิดไฟล์นี้ในเบราว์เซอร์ได้เลย) |
| `img/<ชื่อรูป>-<th|en|my>.webp` | รูปหน้าจอแต่ละภาษา |
| `build.py` | รวมเนื้อหาเป็นไฟล์เดียว `manual.html` สำหรับเผยแพร่ |
| `tools/shots.js`, `tools/webp.py` | ถ่ายรูปหน้าจอใหม่จากแอป (รันแอปที่ `sales-app` ด้วย `python3 -m http.server 8765` ก่อน) |

## เพิ่มหมวดหรือฟังก์ชันใหม่

1. เพิ่ม object ใหม่ใน `chapters` (หรือ `quick`) ของ `content.th.js` — กำหนด `id`, `group`, `icon`, `title`, `who`, `intro`, `steps`, `shots`, `tips`
2. เพิ่ม object ที่ `id` เดียวกันใน `content.en.js` และ `content.my.js` (ใช้คำตามเมนูในแอป จาก `sales-app/js/i18n-words.js`)
3. ถ่ายรูปหน้าจอใหม่: เพิ่มขั้นใน `tools/shots.js` แล้วรัน `node docs/manual/tools/shots.js /tmp/manual-raw` และ `python3 docs/manual/tools/webp.py /tmp/manual-raw`
4. เพิ่มบรรทัดใน `changelog` และเปลี่ยน `meta.version` / `meta.updated` ทั้ง 3 ภาษา
5. รัน `python3 docs/manual/build.py <โฟลเดอร์>` แล้วเผยแพร่ `manual.html` พร้อมโฟลเดอร์ `img/` ใหม่
