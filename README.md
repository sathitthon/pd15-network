# บ้านเน็ต PD15

หน้า 3D แสดงสถานะเครือข่ายของ 2BE HOME ประดิษฐ์มนูธรรม 15 (ภาพรวม ไม่มีข้อมูลรายห้องหรือข้อมูลภายใน)

- router (MikroTik script `dash-push`) ส่งยอดสรุปทุก 5 นาที → Google Apps Script (`apps-script-Code.gs`) → JSON
- `index.html` ดึง JSON ทุก 1 นาที แล้วแสดงในโมเดล 3D (Three.js)
