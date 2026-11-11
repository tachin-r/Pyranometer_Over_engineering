#include <SPI.h>
#include <WiFi.h>
#include <WebServer.h>
#include <math.h>

// ==============================================================================
// ⚙️ เลือกรุ่นฮาร์ดแวร์ที่กำลังทดสอบ (Hardware Target Selection)
// ==============================================================================
// เลือกว่ากำลังรันโค้ดบน "บอร์ดไข่ปลา (Perfboard Prototype)" หรือ "บอร์ดจริง (PCB L1_Analog)"
// เพียงแค่ Comment / Uncomment 1 บรรทัดด้านล่างนี้:
#define TARGET_PERFBOARD    // <-- กำลังทดสอบบนบอร์ดทดลองไข่ปลา (Perfboard)
// #define TARGET_PCB       // <-- ใช้งานบนแผ่นปริ้นท์จริง (Custom PCB L1_Analog)
// ==============================================================================

// ================= ตั้งค่า Wi-Fi =================
const char* ssid     = "Khaitapad69";     // <-- ใส่ชื่อ Wi-Fi ของคุณ (รองรับ 2.4GHz)
const char* password = "Isuskuy555";     // <-- ใส่รหัสผ่าน Wi-Fi ของคุณ

WebServer server(80);

// ================= ESP32-C6 Pinout Definitions =================
// กำหนดขา SPI และควบคุม ADS1220:
#define PIN_DRDY  5   // DRDY (Data Ready จาก ADS1220 ขา 14)
#define PIN_MISO  4   // MISO / DOUT (จาก ADS1220 ขา 15)
#define PIN_MOSI  3   // MOSI / DIN  (เข้า ADS1220 ขา 16)
#define PIN_SCLK  2   // SCLK        (เข้า ADS1220 ขา 1)
#define PIN_CS    1   // CS          (เข้า ADS1220 ขา 2)
#define PIN_LED   15  // LED แสดงสถานะบนบอร์ด ESP32-C6

// ================= ADS1220 Commands =================
#define CMD_RESET 0x06
#define CMD_START 0x08
#define CMD_RDATA 0x10
#define CMD_WREG  0x40

const float VREF = 4.5;       // แรงดันอ้างอิงจาก REF5045 (+4.5Vref ต่อเข้า REFP0 ขา 9)
const float R_TIA = 2200.0;   // ตัวต้านทาน Feedback R1 = 2.2k บน OPA388/OPA380
const uint32_t DRDY_TIMEOUT_MS = 150;

// ================= พารามิเตอร์ NTC Thermistor (TTC3A103F39HAEY) =================
// อ้างอิงตาม Datasheet Thinking Electronic TTC 03 Type:
// R25 = 10 kΩ (±1%) ที่ 25 °C, B25/85 = 3975 K (±1.5%)
// วงจรแบ่งแรงดัน: +4.5Vref -> R7 (10 kΩ) -> [โหนด NTC] -> RT1 (NTC 10 kΩ) -> GND
const float NTC_R25 = 10000.0;     // 10 kΩ
const float NTC_T25 = 298.15;      // 25 °C ในหน่วย Kelvin (25 + 273.15)
const float NTC_BETA = 3975.0;     // Beta B25/85 = 3975 K
const float R_PULLUP = 10000.0;    // R7 = 10 kΩ
const float V_SUPPLY_NTC = 4.5;    // ต่อกับรางไฟ +4.5Vref (Ratiometric กับ VREF)

// ================= ADS1220 MUX Mapping ตามบอร์ด =================
#if defined(TARGET_PERFBOARD)
  // -------------------------------------------------------------
  // [1] การแมปสำหรับบอร์ดทดลองไข่ปลา (Perfboard Prototype)
  // บน Perfboard: AIN0 = Vop, AIN1 = Vbias (สลับกับบน PCB)
  // -------------------------------------------------------------
  #define HW_TARGET_NAME "Perfboard Prototype"
  #define CH_VOP_IDX    0   // Vop อยู่ที่ AIN0
  #define CH_VBIAS_IDX  1   // Vbias อยู่ที่ AIN1
  #define CH_TEMP_IDX   2   // Temp (REF5045) อยู่ที่ AIN2
  #define CH_NTC_IDX    3   // NTC อยู่ที่ AIN3

  const char* CH_LABELS[4] = {
    "AIN0 (Vop - TIA Out)",
    "AIN1 (Vbias)",
    "AIN2 (Temp REF5045)",
    "AIN3 (NTC Board)"
  };

#else
  // -------------------------------------------------------------
  // [2] การแมปสำหรับแผ่นปริ้นท์จริง (Custom PCB L1_Analog)
  // บน PCB: AIN0 = Vbias, AIN1 = Vop (ตาม Schematic KiCad L1_Analog)
  // -------------------------------------------------------------
  #define HW_TARGET_NAME "Custom PCB (L1_Analog)"
  #define CH_VOP_IDX    1   // Vop อยู่ที่ AIN1
  #define CH_VBIAS_IDX  0   // Vbias อยู่ที่ AIN0
  #define CH_TEMP_IDX   2   // Temp (REF5045) อยู่ที่ AIN2
  #define CH_NTC_IDX    3   // NTC อยู่ที่ AIN3

  const char* CH_LABELS[4] = {
    "AIN0 (Vbias)",
    "AIN1 (Vop - TIA Out)",
    "AIN2 (Temp REF5045)",
    "AIN3 (NTC Board)"
  };
#endif

// ตาราง MUX Single-ended vs AVSS (GND), Gain = 1, PGA Bypass
// 0x81 (AIN0), 0x91 (AIN1), 0xA1 (AIN2), 0xB1 (AIN3)
const uint8_t MUX_CONFIGS[4] = {
  0b10000001, // CH0: AIN0
  0b10010001, // CH1: AIN1
  0b10100001, // CH2: AIN2
  0b10110001  // CH3: AIN3
};

SPISettings adsSpiSettings(1000000, MSBFIRST, SPI_MODE1);

// ตัวแปรเก็บค่าล่าสุดสำหรับส่งขึ้นหน้าเว็บและ Serial
int32_t g_raw[4] = {0};
float   g_volt[4] = {0.0};
int16_t g_rawAdsTemp = 0;
float   g_adsTempC = 0.0;
float   g_vDiodeActual = 0.0;
float   g_iPhoto_uA = 0.0;
float   g_refTempC = 0.0;
float   g_rNtc = 0.0;
float   g_ntcTempC = 0.0;
bool    g_allSuccess = true;

unsigned long lastReadTime = 0;
const unsigned long READ_INTERVAL = 1000; // อ่านค่าทุกๆ 1 วินาที

// ================= ฟังก์ชันคำนวณอุณหภูมิ NTC จากแรงดัน =================
float calculateNtcTemp(float vNtc) {
  if (vNtc <= 0.01 || vNtc >= (V_SUPPLY_NTC - 0.01)) {
    return -999.0; // ค่าผิดพลาด (เซนเซอร์หลุดหรือลัดวงจร)
  }
  float rNtc = R_PULLUP * (vNtc / (V_SUPPLY_NTC - vNtc));
  g_rNtc = rNtc;

  // สมการ B-parameter: 1/T = 1/T0 + (1/B) * ln(R/R0)
  float steinhart = log(rNtc / NTC_R25) / NTC_BETA;
  steinhart += 1.0 / NTC_T25;
  float tempK = 1.0 / steinhart;
  return tempK - 273.15;
}

// ================= ฟังก์ชันควบคุม ADS1220 =================
void sendCommand(uint8_t cmd) {
  SPI.beginTransaction(adsSpiSettings);
  digitalWrite(PIN_CS, LOW);
  delayMicroseconds(2);
  SPI.transfer(cmd);
  delayMicroseconds(2);
  digitalWrite(PIN_CS, HIGH);
  SPI.endTransaction();
}

void writeRegister(uint8_t reg, uint8_t value) {
  SPI.beginTransaction(adsSpiSettings);
  digitalWrite(PIN_CS, LOW);
  delayMicroseconds(2);
  SPI.transfer(CMD_WREG | (reg << 2));
  SPI.transfer(value);
  delayMicroseconds(2);
  digitalWrite(PIN_CS, HIGH);
  SPI.endTransaction();
}

void writeAllRegisters(uint8_t r0, uint8_t r1, uint8_t r2, uint8_t r3) {
  SPI.beginTransaction(adsSpiSettings);
  digitalWrite(PIN_CS, LOW);
  delayMicroseconds(2);
  SPI.transfer(CMD_WREG | 0x03); 
  SPI.transfer(r0);
  SPI.transfer(r1);
  SPI.transfer(r2);
  SPI.transfer(r3);
  delayMicroseconds(2);
  digitalWrite(PIN_CS, HIGH);
  SPI.endTransaction();
}

bool waitForDRDY(uint32_t timeoutMs) {
  uint32_t startMillis = millis();
  while (digitalRead(PIN_DRDY) == HIGH) {
    if (millis() - startMillis >= timeoutMs) {
      return false;
    }
    yield();
  }
  return true;
}

void recoverADS1220() {
  Serial.println("[WARNING] ADS1220 DRDY Timeout! Resetting ADC...");
  sendCommand(CMD_RESET);
  delay(10);
}

bool readADC24(int32_t &rawOut) {
  if (!waitForDRDY(DRDY_TIMEOUT_MS)) {
    recoverADS1220();
    rawOut = 0;
    return false;
  }

  SPI.beginTransaction(adsSpiSettings);
  digitalWrite(PIN_CS, LOW);
  delayMicroseconds(2);
  SPI.transfer(CMD_RDATA);
  uint8_t b1 = SPI.transfer(0x00);
  uint8_t b2 = SPI.transfer(0x00);
  uint8_t b3 = SPI.transfer(0x00);
  delayMicroseconds(2);
  digitalWrite(PIN_CS, HIGH);
  SPI.endTransaction();

  int32_t raw = ((int32_t)b1 << 16) | ((int32_t)b2 << 8) | b3;
  if (raw & 0x800000) {
    raw |= 0xFF000000;
  }
  rawOut = raw;
  return true;
}

bool readChannel(uint8_t chIndex, int32_t &rawOut, float &voltOut) {
  writeAllRegisters(MUX_CONFIGS[chIndex], 0x00, 0x40, 0x00);
  sendCommand(CMD_START);

  if (readADC24(rawOut)) {
    voltOut = ((float)rawOut / 8388607.0) * VREF;
    return true;
  } else {
    voltOut = 0.0;
    return false;
  }
}

bool readADS1220Temp(int16_t &rawOut, float &tempOut) {
  writeRegister(1, 0x02); // TS = 1 (เลือก Internal Temperature Sensor)
  sendCommand(CMD_START);

  if (!waitForDRDY(DRDY_TIMEOUT_MS)) {
    recoverADS1220();
    rawOut = 0;
    tempOut = 0.0;
    return false;
  }

  SPI.beginTransaction(adsSpiSettings);
  digitalWrite(PIN_CS, LOW);
  delayMicroseconds(2);
  SPI.transfer(CMD_RDATA);
  uint8_t b1 = SPI.transfer(0x00);
  uint8_t b2 = SPI.transfer(0x00);
  SPI.transfer(0x00);
  delayMicroseconds(2);
  digitalWrite(PIN_CS, HIGH);
  SPI.endTransaction();

  int16_t raw = ((int16_t)b1 << 8) | b2;
  raw = raw >> 2;

  rawOut = raw;
  tempOut = (float)raw * 0.03125;
  return true;
}

// ================= ฟังก์ชัน Web Server (HTML & JSON) =================
void handleRoot() {
  String html = R"rawliteral(
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>ESP32-C6 ADS1220 Precision Pyranometer Monitor</title>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #121212; color: #e0e0e0; margin: 0; padding: 20px; }
    h2 { color: #00e676; text-align: center; margin-bottom: 5px; }
    .hw-mode { text-align: center; color: #64b5f6; font-size: 0.95em; margin-bottom: 20px; }
    .container { max-width: 650px; margin: auto; background: #1e1e1e; padding: 20px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.5); }
    table { width: 100%; border-collapse: collapse; margin-top: 15px; }
    th, td { padding: 10px; text-align: left; border-bottom: 1px solid #333; }
    th { color: #81c784; }
    .highlight { background: #263238; padding: 15px; border-radius: 8px; margin-top: 15px; border-left: 4px solid #00e676; }
    .highlight p { margin: 8px 0; }
    .status-ok { color: #00e676; font-weight: bold; }
    .status-err { color: #ff5252; font-weight: bold; }
    .badge { background: #004d40; color: #80cbc4; padding: 2px 6px; border-radius: 4px; font-size: 0.85em; }
    .badge-hw { background: #1565c0; color: #e3f2fd; padding: 3px 8px; border-radius: 4px; font-size: 0.85em; }
  </style>
</head>
<body>
  <div class="container">
    <h2>ESP32-C6 & ADS1220 Monitor</h2>
    <div class="hw-mode">โหมดฮาร์ดแวร์: <span class="badge-hw">)rawliteral" + String(HW_TARGET_NAME) + R"rawliteral(</span></div>
    <p>สถานะ ADC: <span id="status" class="status-ok">กำลังโหลด...</span></p>
    <table>
      <thead>
        <tr><th>ช่องสัญญาณ ADC</th><th>ค่า Raw (24-bit)</th><th>แรงดันที่วัดได้ (V)</th></tr>
      </thead>
      <tbody id="ch-table">
        <tr><td colspan="3">กำลังรอข้อมูล...</td></tr>
      </tbody>
    </table>

    <div class="highlight">
      <p><strong>⚡ V_diode (Vop - Vbias):</strong> <span id="vdiode">0.000</span> mV</p>
      <p><strong>💡 Photocurrent (I_photo):</strong> <span id="iphoto">0.000</span> &mu;A</p>
      <hr style="border: 0; border-top: 1px solid #37474f; margin: 10px 0;">
      <p><strong>🌡️ อุณหภูมิบอร์ด NTC (TTC3A103):</strong> <span id="ntctemp" style="color:#ffca28; font-size:1.15em; font-weight:bold;">0.00</span> &deg;C <span class="badge" id="rntc">10.00 k&Omega;</span></p>
      <p><strong>🌡️ อุณหภูมิ REF5045:</strong> <span id="reftemp">0.00</span> &deg;C</p>
      <p><strong>🔥 อุณหภูมิภายในชิป ADS1220:</strong> <span id="adstemp">0.00</span> &deg;C</p>
    </div>
  </div>

  <script>
    function fetchData() {
      fetch('/data')
        .then(response => response.json())
        .then(data => {
          const statusEl = document.getElementById('status');
          if (data.success) {
            statusEl.innerText = 'ปกติ (OK)';
            statusEl.className = 'status-ok';
          } else {
            statusEl.innerText = 'Timeout (กำลังกู้คืน...)';
            statusEl.className = 'status-err';
          }

          let rows = '';
          for (let i = 0; i < 4; i++) {
            rows += `<tr>
              <td>${data.labels[i]}</td>
              <td>${data.raw[i]}</td>
              <td>${data.volt[i].toFixed(5)} V</td>
            </tr>`;
          }
          document.getElementById('ch-table').innerHTML = rows;
          document.getElementById('vdiode').innerText = (data.vDiode * 1000).toFixed(3);
          document.getElementById('iphoto').innerText = data.iPhoto.toFixed(3);
          document.getElementById('ntctemp').innerText = data.ntcTemp.toFixed(2);
          document.getElementById('rntc').innerText = (data.rNtc / 1000.0).toFixed(2) + ' k\u03A9';
          document.getElementById('reftemp').innerText = data.refTemp.toFixed(2);
          document.getElementById('adstemp').innerText = data.adsTemp.toFixed(2);
        })
        .catch(err => console.error(err));
    }
    setInterval(fetchData, 1000);
    fetchData();
  </script>
</body>
</html>
)rawliteral";
  server.send(200, "text/html", html);
}

void handleData() {
  String json = "{";
  json += "\"success\":" + String(g_allSuccess ? "true" : "false") + ",";
  json += "\"hw\":\"" + String(HW_TARGET_NAME) + "\",";
  json += "\"labels\":[\"" + String(CH_LABELS[0]) + "\",\"" + String(CH_LABELS[1]) + "\",\"" + String(CH_LABELS[2]) + "\",\"" + String(CH_LABELS[3]) + "\"],";
  json += "\"raw\":[" + String(g_raw[0]) + "," + String(g_raw[1]) + "," + String(g_raw[2]) + "," + String(g_raw[3]) + "],";
  json += "\"volt\":[" + String(g_volt[0], 5) + "," + String(g_volt[1], 5) + "," + String(g_volt[2], 5) + "," + String(g_volt[3], 5) + "],";
  json += "\"vDiode\":" + String(g_vDiodeActual, 6) + ",";
  json += "\"iPhoto\":" + String(g_iPhoto_uA, 3) + ",";
  json += "\"rNtc\":" + String(g_rNtc, 1) + ",";
  json += "\"ntcTemp\":" + String(g_ntcTempC, 2) + ",";
  json += "\"refTemp\":" + String(g_refTempC, 2) + ",";
  json += "\"adsTemp\":" + String(g_adsTempC, 2);
  json += "}";
  server.send(200, "application/json", json);
}

// ================= SETUP =================
void setup() {
  Serial.begin(115200);
  delay(500);

  pinMode(PIN_LED, OUTPUT);
  digitalWrite(PIN_LED, LOW);

  pinMode(PIN_CS, OUTPUT);
  pinMode(PIN_DRDY, INPUT_PULLUP); 
  digitalWrite(PIN_CS, HIGH);

  // เริ่มต้น SPI
  SPI.begin(PIN_SCLK, PIN_MISO, PIN_MOSI, PIN_CS);

  delay(100);
  sendCommand(CMD_RESET);
  delay(10);

  // เชื่อมต่อ Wi-Fi
  Serial.println();
  Serial.print("Connecting to Wi-Fi: ");
  Serial.println(ssid);
  WiFi.mode(WIFI_STA);
  WiFi.begin(ssid, password);

  uint8_t wifiTimeout = 0;
  while (WiFi.status() != WL_CONNECTED && wifiTimeout < 20) {
    delay(500);
    Serial.print(".");
    wifiTimeout++;
  }

  if (WiFi.status() == WL_CONNECTED) {
    Serial.println("\n[Wi-Fi Connected!]");
    Serial.print("-> เปิดเบราว์เซอร์แล้วเข้าไปที่ IP: http://");
    Serial.println(WiFi.localIP());
  } else {
    Serial.println("\n[Wi-Fi Failed] ทำงานต่อในโหมด Offline (Serial Monitor เท่านั้น)");
  }

  // ตั้งค่า Web Server
  server.on("/", handleRoot);
  server.on("/data", handleData);
  server.begin();

  Serial.println("=========================================================================================");
  Serial.print("ADS1220 Precision Pyranometer Monitor | Active Mode: ");
  Serial.println(HW_TARGET_NAME);
  Serial.println("=========================================================================================");
}

// ================= LOOP =================
void loop() {
  server.handleClient();

  if (millis() - lastReadTime >= READ_INTERVAL) {
    lastReadTime = millis();

    digitalWrite(PIN_LED, HIGH);
    g_allSuccess = true;

    // 1. กวาดอ่านทั้ง 4 ช่องสัญญาณ
    for (uint8_t i = 0; i < 4; i++) {
      if (!readChannel(i, g_raw[i], g_volt[i])) {
        g_allSuccess = false;
      }
      delay(2);
    }

    // 2. อ่านอุณหภูมิภายในตัวชิป ADS1220
    if (!readADS1220Temp(g_rawAdsTemp, g_adsTempC)) {
      g_allSuccess = false;
    }

    // 3. ประมวลผลสัญญาณจริงของไดโอด (Vop - Vbias เสมอ โดยดึงจาก Index ที่ตรงกับบอร์ดที่เลือก)
    // V_diode = Vop - Vbias
    g_vDiodeActual = g_volt[CH_VOP_IDX] - g_volt[CH_VBIAS_IDX];
    g_iPhoto_uA = (g_vDiodeActual / R_TIA) * 1e6;

    // 4. แปลงอุณหภูมิ REF5045 จากช่อง Temp
    g_refTempC = (g_volt[CH_TEMP_IDX] - 0.509) / 0.00264;

    // 5. แปลงอุณหภูมิ NTC Thermistor บนบอร์ด (TTC3A103F39HAEY จากช่อง NTC)
    g_ntcTempC = calculateNtcTemp(g_volt[CH_NTC_IDX]);

    // 6. แสดงผลทาง Serial Monitor
    Serial.println("-----------------------------------------------------------------------------------------");
    Serial.print("[Active Hardware]: ");
    Serial.println(HW_TARGET_NAME);
    if (!g_allSuccess) {
      Serial.println("[STATUS] Some channels timed out and were recovered.");
    }
    
    for (uint8_t i = 0; i < 4; i++) {
      Serial.print("[");
      Serial.print(CH_LABELS[i]);
      Serial.print("]\tRaw: ");
      Serial.print(g_raw[i]);
      Serial.print("\tV: ");
      Serial.print(g_volt[i], 5);
      Serial.println(" V");
    }

    Serial.print("[ADS1220 Temp]\tRaw: ");
    Serial.print(g_rawAdsTemp);
    Serial.print("\tT: ");
    Serial.print(g_adsTempC, 2);
    Serial.println(" °C");

    Serial.println("--- Computed Values ---");
    Serial.print("-> Net Diode Voltage (Vop - Vbias): ");
    Serial.print(g_vDiodeActual * 1000.0, 3);
    Serial.print(" mV | I_photo: ");
    Serial.print(g_iPhoto_uA, 3);
    Serial.println(" uA");

    Serial.print("-> NTC Board Temperature (CH");
    Serial.print(CH_NTC_IDX);
    Serial.print("): ");
    Serial.print(g_ntcTempC, 2);
    Serial.print(" °C (R_ntc = ");
    Serial.print(g_rNtc / 1000.0, 2);
    Serial.println(" kOhm)");

    Serial.print("-> REF5045 Temperature (CH");
    Serial.print(CH_TEMP_IDX);
    Serial.print("): ");
    Serial.print(g_refTempC, 2);
    Serial.println(" °C");

    if (WiFi.status() == WL_CONNECTED) {
      Serial.print("-> Web Dashboard IP: http://");
      Serial.println(WiFi.localIP());
    }

    digitalWrite(PIN_LED, LOW);
  }
}
