# Pyranometer Over-engineering ☀️⚡

Modular High-Precision Pyranometer KiCad Hardware Project.

---

## 📑 Project Structure & Interactive Online Viewers

Explore the schematics and PCB layouts interactively directly in your browser powered by **KiCanvas** (no installation required):

| Board / Module | Description | Interactive Project (Schematic + PCB) | Schematic Only | PCB Only |
| :--- | :--- | :---: | :---: | :---: |
| **L1_Analog** | Analog Front-End & Sensor Interface | [🚀 Open L1 Board](https://kicanvas.org/?github=https://github.com/tachin-r/Pyranometer_Over_engineering/tree/main/L1_Analog) | [🔍 Schematic](https://kicanvas.org/?github=https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/L1_Analog/L1_Analog.kicad_sch) | [🔍 PCB](https://kicanvas.org/?github=https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/L1_Analog/L1_Analog.kicad_pcb) |
| **L2_MCU** | Microcontroller Processing Board | [🚀 Open L2 Board](https://kicanvas.org/?github=https://github.com/tachin-r/Pyranometer_Over_engineering/tree/main/L2_MCU) | [🔍 Schematic](https://kicanvas.org/?github=https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/L2_MCU/L2_MCU.kicad_sch) | [🔍 PCB](https://kicanvas.org/?github=https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/L2_MCU/L2_MCU.kicad_pcb) |
| **L3_Switching** | Power Management & Switching Regulators | [🚀 Open L3 Board](https://kicanvas.org/?github=https://github.com/tachin-r/Pyranometer_Over_engineering/tree/main/L3_Switching) | [🔍 Schematic](https://kicanvas.org/?github=https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/L3_Switching/L3_Switching.kicad_sch) | [🔍 PCB](https://kicanvas.org/?github=https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/L3_Switching/L3_Switching.kicad_pcb) |
| **LM2596 HW-411** | Reverse-Engineered Buck Converter Module | [🚀 Open LM2596](https://kicanvas.org/?github=https://github.com/tachin-r/Pyranometer_Over_engineering/tree/main/reverse_component/LM2596_HW411) | [🔍 Schematic](https://kicanvas.org/?github=https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/reverse_component/LM2596_HW411/LM2596_HW411.kicad_sch) | [🔍 PCB](https://kicanvas.org/?github=https://github.com/tachin-r/Pyranometer_Over_engineering/blob/main/reverse_component/LM2596_HW411/LM2596_HW411.kicad_pcb) |

---

## 🛠️ Boards Overview

### 1. `L1_Analog` (Sensor Front-End)
* Precision Analog Signal Conditioning stage designed for low-noise measurement of solar irradiance sensor signals (Pyranometer Thermopile / Photodiode).

### 2. `L2_MCU` (Processing & Control)
* Primary microcontroller board responsible for high-resolution ADC sampling, calibration algorithms, digital filtering, and telemetry communication.

### 3. `L3_Switching` (Power Distribution)
* Power supply unit (PSU) providing efficient switching regulation, clean power rail generation, and noise filtering for sensitive analog circuitry.

### 4. `reverse_component/LM2596_HW411`
* Reverse-engineered KiCad schematic and PCB layout of the common LM2596 HW-411 DC-DC step-down converter module, used for reference, testing, and validation.

---

## 🚀 Working with KiCad

This project is organized as a **modular multi-board system**:
* Each sub-directory contains its own standalone KiCad project file (`.kicad_pro`).
* To view or edit a board, open the corresponding `.kicad_pro` file inside its respective directory.
