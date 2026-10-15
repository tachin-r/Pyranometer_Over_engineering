# Pyranometer Over-engineering ☀️⚡

Modular High-Precision Pyranometer KiCad Hardware Project.

---

## 📑 Project Structure & Interactive Online Viewers

คุณสามารถกดเปิดดูวงจรและ PCB ของแต่ละบอร์ดผ่าน **KiCanvas** ได้ทันทีโดยไม่ต้องติดตั้งโปรแกรม KiCad:

| บอร์ด / โมดูล | คำอธิบาย | ดู Schematic (วงจร) | ดู PCB (ลายวงจร) |
| :--- | :--- | :---: | :---: |
| **L1_Analog** | Analog Front-End & Sensor Interface | [🔍 เปิด Schematic](https://kicanvas.org/https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/L1_Analog/L1_Analog.kicad_sch) | [🔍 เปิด PCB](https://kicanvas.org/https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/L1_Analog/L1_Analog.kicad_pcb) |
| **L2_MCU** | Microcontroller Unit (MCU) Board | [🔍 เปิด Schematic](https://kicanvas.org/https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/L2_MCU/L2_MCU.kicad_sch) | [🔍 เปิด PCB](https://kicanvas.org/https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/L2_MCU/L2_MCU.kicad_pcb) |
| **L3_Switching** | Power Management & Switching Regulators | [🔍 เปิด Schematic](https://kicanvas.org/https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/L3_Switching/L3_Switching.kicad_sch) | [🔍 เปิด PCB](https://kicanvas.org/https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/L3_Switching/L3_Switching.kicad_pcb) |
| **LM2596 HW-411** | Reverse Engineered Buck Converter Module | [🔍 เปิด Schematic](https://kicanvas.org/https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/reverse_component/LM2596_HW411/LM2596_HW411.kicad_sch) | [🔍 เปิด PCB](https://kicanvas.org/https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/reverse_component/LM2596_HW411/LM2596_HW411.kicad_pcb) |

---

## 🛠️ รายละเอียดของแต่ละบอร์ด (Boards Overview)

### 1. `L1_Analog`
* บอร์ดขยายและปรับสภาพสัญญาณแอนะล็อกความแม่นยำสูง (Precision Analog Signal Conditioning) จากเซนเซอร์วัดค่าความเข้มแสงอาทิตย์ (Pyranometer Thermopile / Photodiode)

### 2. `L2_MCU`
* บอร์ดประมวลผลหลัก (Microcontroller Board) รับสัญญาณ ADC, ประมวลผลทางคณิตศาสตร์ และสื่อสารข้อมูล

### 3. `L3_Switching`
* บอร์ดจ่ายพลังงาน (Power Supply Unit) ภาคจ่ายไฟแบบ Switching แปลงแรงดันไฟเลี้ยงสำหรับทั้งระบบ พร้อมวงจรกรองสัญญาณรบกวน

### 4. `reverse_component/LM2596_HW411`
* วงจรและ PCB ที่ถอดแบบ (Reverse-engineered) ของโมดูล Buck Converter LM2596 HW-411 เพื่อใช้อ้างอิงและประกอบการทดสอบ

---

## 🚀 การเปิดใช้งานด้วย KiCad
โปรเจกต์นี้สร้างขึ้นด้วย **KiCad**:
* แต่ละโฟลเดอร์มีไฟล์ `.kicad_pro` แยกของตนเอง สามารถดับเบิลคลิกไฟล์โปรเจกต์ในแต่ละโฟลเดอร์เพื่อเปิดแก้ไขได้โดยตรง
