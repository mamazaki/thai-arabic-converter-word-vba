# Thai-Arabic Number Converter for Microsoft Word

ชุดคำสั่ง VBA (Macro) สำหรับการแปลงตัวเลขไทยเป็นอาราบิก และอาราบิกเป็นไทย ในโปรแกรม Microsoft Word 
เหมาะสำหรับงานเอกสารราชการที่ต้องการสลับรูปแบบตัวเลขตามระเบียบงานสารบรรณ

## วิธีติดตั้ง
1. เปิด Microsoft Word
2. กด `Alt + F11` เพื่อเปิดหน้าต่าง Microsoft Visual Basic for Applications
3. ไปที่เมนู `Insert` > `Module`
4. คัดลอก Code จากไฟล์ `ConvertNumber.bas` ไปวางใน Module ที่สร้างขึ้น
5. กด Save และปิดหน้าต่าง VBA

## วิธีใช้งาน
1. ไปที่แถบเมนู **View (มุมมอง)** > **Macros (แมโคร)**
2. เลือกชื่อแมโครที่ต้องการ:
   - `Arabic2thai`: แปลง 123 -> ๑๒๓
   - `Thai2Arabic`: แปลง ๑๒๓ -> 123
3. กด **Run (เรียกใช้)**

## ข้อควรระวัง
- คำสั่งนี้จะเปลี่ยนตัวเลข **ทั้งเอกสาร** (wdFindContinue)
- ไฟล์ Word ที่มี Macro ต้องบันทึกเป็นนามสกุล `.docm` เท่านั้น
