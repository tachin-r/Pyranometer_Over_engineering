import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches

def generate_system_block_diagram():
    fig, ax = plt.subplots(figsize=(14, 8.5), dpi=300)
    ax.set_xlim(0, 14)
    ax.set_ylim(0, 8.8)
    ax.axis('off')

    # Color Palette - Professional Engineering
    c_bg_head = '#F0F7FF'
    c_bdr_head = '#0055AA'
    
    c_bg_afe = '#F0FDF4'
    c_bdr_afe = '#15803D'
    
    c_bg_mcu = '#FAF5FF'
    c_bdr_mcu = '#7E22CE'
    
    c_bg_pwr = '#FFF7ED'
    c_bdr_pwr = '#C2410C'

    # -------------------------------------------------------------
    # 1. Outer Container Boxes (Dashed Grouping)
    # -------------------------------------------------------------
    # Group 1: Optical & Sensing Head
    rect_g1 = patches.FancyBboxPatch(
        (0.4, 2.5), 3.4, 5.8,
        boxstyle="round,pad=0.15",
        facecolor=c_bg_head, edgecolor=c_bdr_head,
        linestyle='--', linewidth=1.5, alpha=0.6
    )
    ax.add_patch(rect_g1)
    ax.text(0.55, 8.1, "1. Optical & Sensing Head", fontsize=11, fontweight='bold', color=c_bdr_head)

    # Group 2: Analog Front-End (ADS1220 + REF5045)
    rect_g2 = patches.FancyBboxPatch(
        (4.2, 2.5), 4.7, 5.8,
        boxstyle="round,pad=0.15",
        facecolor=c_bg_afe, edgecolor=c_bdr_afe,
        linestyle='--', linewidth=1.5, alpha=0.6
    )
    ax.add_patch(rect_g2)
    ax.text(4.35, 8.1, "2. High-Precision Analog Front-End (24-bit)", fontsize=11, fontweight='bold', color=c_bdr_afe)

    # Group 3: Processing & IoT Platform
    rect_g3 = patches.FancyBboxPatch(
        (9.3, 2.5), 4.3, 5.8,
        boxstyle="round,pad=0.15",
        facecolor=c_bg_mcu, edgecolor=c_bdr_mcu,
        linestyle='--', linewidth=1.5, alpha=0.6
    )
    ax.add_patch(rect_g3)
    ax.text(9.45, 8.1, "3. Processing & IoT Web Platform", fontsize=11, fontweight='bold', color=c_bdr_mcu)

    # Group 4: Power Supply Architecture (Bottom)
    rect_g4 = patches.FancyBboxPatch(
        (0.4, 0.4), 13.2, 1.8,
        boxstyle="round,pad=0.15",
        facecolor=c_bg_pwr, edgecolor=c_bdr_pwr,
        linestyle='--', linewidth=1.5, alpha=0.6
    )
    ax.add_patch(rect_g4)
    ax.text(0.55, 1.95, "4. Power Supply Architecture (5.2 V System Rail & Low-Noise Roadmap)", fontsize=11, fontweight='bold', color=c_bdr_pwr)

    # -------------------------------------------------------------
    # 2. Section 1 Blocks (Optical & Sensors)
    # -------------------------------------------------------------
    # Block 1A: Photodiode + TIA
    b1a = patches.FancyBboxPatch(
        (0.6, 6.2), 3.0, 1.6,
        boxstyle="round,pad=0.1",
        facecolor='#FFFFFF', edgecolor='#0284C7', linewidth=1.8
    )
    ax.add_patch(b1a)
    ax.text(2.1, 7.5, "BPW 34 Silicon Photodiode", fontsize=9.5, fontweight='bold', ha='center', color='#0369A1')
    ax.text(2.1, 7.15, "+ Synthetic Sapphire Window (Al₂O₃)", fontsize=8, ha='center', color='#334155')
    ax.text(2.1, 6.75, "OPA380 TIA Front-End (R_TIA = 2.2 kΩ)", fontsize=8.2, fontweight='bold', ha='center', color='#0F172A')
    ax.text(2.1, 6.4, "Output: V_diode (V_op) & V_bias", fontsize=8, ha='center', color='#0284C7')

    # Block 1B: NTC Surface Temp
    b1b = patches.FancyBboxPatch(
        (0.6, 4.5), 3.0, 1.4,
        boxstyle="round,pad=0.1",
        facecolor='#FFFFFF', edgecolor='#0284C7', linewidth=1.5
    )
    ax.add_patch(b1b)
    ax.text(2.1, 5.55, "NTC Thermistor (10 kΩ ±1%)", fontsize=9.5, fontweight='bold', ha='center', color='#0369A1')
    ax.text(2.1, 5.2, "Thinking TTC3A103F39HAEY", fontsize=8.2, ha='center', color='#0F172A')
    ax.text(2.1, 4.85, "B₂₅/₈₅ = 3950 K (Surface Temp)", fontsize=8, ha='center', color='#475569')
    ax.text(2.1, 4.55, "Precision Voltage Divider", fontsize=7.8, ha='center', color='#0284C7')

    # Block 1C: REF5045 Temp Pin 3
    b1c = patches.FancyBboxPatch(
        (0.6, 2.8), 3.0, 1.4,
        boxstyle="round,pad=0.1",
        facecolor='#FFFFFF', edgecolor='#D97706', linewidth=1.5
    )
    ax.add_patch(b1c)
    ax.text(2.1, 3.85, "REF5045 TEMP (Pin 3)", fontsize=9.5, fontweight='bold', ha='center', color='#B45309')
    ax.text(2.1, 3.5, "On-Chip Thermal Sensor", fontsize=8.2, ha='center', color='#0F172A')
    ax.text(2.1, 3.15, "V_TEMP = 509 mV + 2.64 mV/°C", fontsize=8, ha='center', color='#475569')
    ax.text(2.1, 2.85, "Monitors Reference Die Drift", fontsize=7.8, ha='center', color='#B45309')

    # -------------------------------------------------------------
    # 3. Section 2 Blocks (AFE: ADS1220 + REF5045)
    # -------------------------------------------------------------
    # Block 2A: ADS1220 ADC Main Container
    b2a = patches.FancyBboxPatch(
        (4.4, 4.3), 4.3, 3.5,
        boxstyle="round,pad=0.12",
        facecolor='#FFFFFF', edgecolor='#16A34A', linewidth=2.0
    )
    ax.add_patch(b2a)
    ax.text(6.55, 7.5, "TI ADS1220 (24-bit ΔΣ ADC)", fontsize=10.5, fontweight='bold', ha='center', color='#15803D')
    
    # Internal details inside ADS1220
    box_mux = patches.FancyBboxPatch(
        (4.55, 5.0), 4.0, 2.2,
        boxstyle="round,pad=0.08",
        facecolor='#F0FDF4', edgecolor='#86EFAC', linewidth=1.0
    )
    ax.add_patch(box_mux)
    ax.text(6.55, 6.95, "Single-Ended Input MUX vs AVSS (GND):", fontsize=8.5, fontweight='bold', ha='center', color='#166534')
    ax.text(4.7, 6.6, "• AIN0 (0x81): REF5045 Pin 3 TEMP", fontsize=8, color='#0F172A')
    ax.text(4.7, 6.25, "• AIN1 (0x91): NTC Surface Temp Divider", fontsize=8, color='#0F172A')
    ax.text(4.7, 5.9, "• AIN2 (0xA1): V_diode (Photocurrent Signal)", fontsize=8, color='#0F172A')
    ax.text(4.7, 5.55, "• AIN3 (0xB1): V_bias (Baseline / Virtual GND)", fontsize=8, color='#0F172A')
    ax.text(4.7, 5.2, "• Internal Temp (TS = 1, 0.03125 °C/LSB)", fontsize=8, color='#047857')

    ax.text(6.55, 4.65, "V_REF = 4.500 V (REFP0/REFN0) | Gain = 1 (PGA Bypassed)", fontsize=7.8, ha='center', fontweight='bold', color='#15803D')
    ax.text(6.55, 4.4, "Data Rate: 20 SPS (Normal Mode, Sinc³ 50/60Hz Filter)", fontsize=7.5, ha='center', color='#475569')

    # Block 2B: REF5045AID Voltage Reference
    b2b = patches.FancyBboxPatch(
        (4.4, 2.7), 4.3, 1.3,
        boxstyle="round,pad=0.1",
        facecolor='#FFFFFF', edgecolor='#D97706', linewidth=1.8
    )
    ax.add_patch(b2b)
    ax.text(6.55, 3.65, "TI REF5045AID Precision Reference", fontsize=9.5, fontweight='bold', ha='center', color='#B45309')
    ax.text(6.55, 3.3, "Output: 4.500 V ±0.05% | Low Drift: 3 ppm/°C", fontsize=8.2, ha='center', color='#0F172A')
    ax.text(6.55, 2.95, "Load Capacitor: 10 µF X7R Ceramic (Low-ESR)", fontsize=8, ha='center', color='#475569')

    # -------------------------------------------------------------
    # 4. Section 3 Blocks (MCU & IoT Platform)
    # -------------------------------------------------------------
    # Block 3A: ESP32-C6 Super Mini
    b3a = patches.FancyBboxPatch(
        (9.5, 4.7), 3.9, 3.1,
        boxstyle="round,pad=0.12",
        facecolor='#FFFFFF', edgecolor='#9333EA', linewidth=2.0
    )
    ax.add_patch(b3a)
    ax.text(11.45, 7.5, "ESP32-C6 Super Mini (RISC-V)", fontsize=10.5, fontweight='bold', ha='center', color='#7E22CE')
    
    ax.text(9.7, 7.05, "Hardware SPI Master (1 MHz, Mode 1):", fontsize=8.2, fontweight='bold', color='#6B21A8')
    ax.text(9.7, 6.75, " SCLK: GPIO 6 (Green)  | MISO: GPIO 2 (Blue)", fontsize=7.8, color='#0F172A')
    ax.text(9.7, 6.5, " MOSI: GPIO 4 (Gray)   | CS:   GPIO 14 (Cyan)", fontsize=7.8, color='#0F172A')
    ax.text(9.7, 6.2, "Handshake: DRDY on GPIO 0 (Watchdog 150ms)", fontsize=7.8, fontweight='bold', color='#DC2626')
    ax.text(9.7, 5.95, "Status Indicator: Onboard LED (GPIO 15)", fontsize=7.8, color='#475569')

    box_math = patches.FancyBboxPatch(
        (9.65, 4.85), 3.6, 0.95,
        boxstyle="round,pad=0.06",
        facecolor='#FAF5FF', edgecolor='#D8B4FE', linewidth=1.0
    )
    ax.add_patch(box_math)
    ax.text(11.45, 5.55, "Firmware Mathematical Engine:", fontsize=8, fontweight='bold', ha='center', color='#7E22CE')
    ax.text(11.45, 5.25, "V_diode_actual = V_AIN2 - V_AIN3 (Pseudo-Diff)", fontsize=7.5, ha='center', color='#0F172A')
    ax.text(11.45, 4.95, "I_photo = V_diode_actual / 2,200 Ω", fontsize=7.5, fontweight='bold', ha='center', color='#9333EA')

    # Block 3B: IoT Web Platform & Serial
    b3b = patches.FancyBboxPatch(
        (9.5, 2.7), 3.9, 1.7,
        boxstyle="round,pad=0.1",
        facecolor='#FFFFFF', edgecolor='#9333EA', linewidth=1.5
    )
    ax.add_patch(b3b)
    ax.text(11.45, 4.1, "IoT Connectivity & User Interface", fontsize=9.5, fontweight='bold', ha='center', color='#7E22CE')
    ax.text(9.7, 3.7, "• Wi-Fi 2.4 GHz Embedded Web Server (Port 80)", fontsize=8, color='#0F172A')
    ax.text(9.7, 3.4, "• Real-time HTML5 Dashboard (Live AJAX 1 Hz)", fontsize=8, color='#0F172A')
    ax.text(9.7, 3.1, "• RESTful JSON API (/data) for Cloud Logging", fontsize=8, color='#0F172A')
    ax.text(9.7, 2.8, "• High-Speed USB Serial Output (115,200 bps)", fontsize=8, color='#475569')

    # -------------------------------------------------------------
    # 5. Section 4 Blocks (Power Supply - 5.2 V Switching)
    # -------------------------------------------------------------
    # Block 4A: 5.2V Switching Power Supply
    b4a = patches.FancyBboxPatch(
        (0.6, 0.6), 5.8, 1.35,
        boxstyle="round,pad=0.1",
        facecolor='#FFFFFF', edgecolor='#EA580C', linewidth=2.0
    )
    ax.add_patch(b4a)
    ax.text(3.5, 1.65, "Regulated DC Switching Power Supply (5.2 V)", fontsize=9.5, fontweight='bold', ha='center', color='#C2410C')
    ax.text(3.5, 1.35, "Dedicated 5.2 V Rail (Bypasses USB Dropout -> V_IN > 4.70 V)", fontsize=8.2, ha='center', color='#0F172A')
    ax.text(3.5, 1.05, "Headroom: 5.2 V - 4.5 V = 700 mV (Well above V_DO ≈ 200 mV)", fontsize=8, fontweight='bold', ha='center', color='#15803D')
    ax.text(3.5, 0.75, "Powers REF5045 V_IN & ESP32-C6 System Bus Stable on Breadboard", fontsize=7.8, ha='center', color='#475569')

    # Block 4B: Planned LDO Post-Regulator
    b4b = patches.FancyBboxPatch(
        (7.4, 0.6), 6.0, 1.35,
        boxstyle="round,pad=0.1",
        facecolor='#FFFFFF', edgecolor='#059669', linewidth=1.8, linestyle='--'
    )
    ax.add_patch(b4b)
    ax.text(10.4, 1.65, "Planned Capstone 2: Ultra-Low-Noise LDO Post-Regulator", fontsize=9.5, fontweight='bold', ha='center', color='#047857')
    ax.text(10.4, 1.35, "Architecture: 5.2V Switching -> Low-Noise LDO (TPS7A / LP5907)", fontsize=8.2, ha='center', color='#0F172A')
    ax.text(10.4, 1.05, "Filters High-Frequency Switching Ripple (< 4 µV_RMS Noise Floor)", fontsize=8, ha='center', color='#0F172A')
    ax.text(10.4, 0.75, "Targeting Metrology Grade Compliance with ISO 9060 Class A", fontsize=7.8, fontweight='bold', ha='center', color='#059669')

    # -------------------------------------------------------------
    # 6. Connecting Arrows & Routing (Clean & No Overlaps)
    # -------------------------------------------------------------
    # 6.1 Signals from Section 1 to ADS1220
    # AIN2 & AIN3 from Photodiode/TIA
    ax.annotate(
        "", xy=(4.4, 6.7), xytext=(3.6, 6.7),
        arrowprops=dict(arrowstyle="->", color="#0284C7", lw=1.8)
    )
    ax.text(4.0, 6.85, "AIN2 (V_diode)\nAIN3 (V_bias)", fontsize=7.5, ha='center', va='bottom', color='#0369A1',
            bbox=dict(boxstyle='round,pad=0.18', facecolor='#F0F9FF', edgecolor='#BAE6FD', alpha=0.95))

    # AIN1 from NTC
    ax.annotate(
        "", xy=(4.4, 5.8), xytext=(3.6, 5.2),
        arrowprops=dict(arrowstyle="->", color="#0284C7", lw=1.8)
    )
    ax.text(3.9, 5.4, "AIN1 (NTC)", fontsize=7.5, ha='center', va='center', color='#0369A1',
            bbox=dict(boxstyle='round,pad=0.18', facecolor='#F0F9FF', edgecolor='#BAE6FD', alpha=0.95))

    # AIN0 from REF5045 Temp Pin 3
    ax.annotate(
        "", xy=(4.4, 5.1), xytext=(3.6, 3.5),
        arrowprops=dict(arrowstyle="->", color="#D97706", lw=1.8)
    )
    ax.text(3.85, 4.3, "AIN0 (REF Temp)", fontsize=7.5, ha='center', va='center', color='#B45309',
            bbox=dict(boxstyle='round,pad=0.18', facecolor='#FFFBEB', edgecolor='#FDE68A', alpha=0.95))

    # 6.2 Reference Voltage from REF5045 to ADS1220 (Vertical Arrow)
    ax.annotate(
        "", xy=(6.55, 4.3), xytext=(6.55, 4.0),
        arrowprops=dict(arrowstyle="->", color="#B45309", lw=2.0)
    )
    ax.text(6.55, 4.15, "V_REF = 4.500 V (REFP0 / REFN0)", fontsize=7.5, fontweight='bold', ha='center', va='center', color='#B45309',
            bbox=dict(boxstyle='round,pad=0.18', facecolor='#FFFBEB', edgecolor='#FDE68A', alpha=0.95))

    # 6.3 Bus Connection from ADS1220 to ESP32-C6
    ax.annotate(
        "", xy=(9.5, 6.3), xytext=(8.7, 6.3),
        arrowprops=dict(arrowstyle="<->", color="#7E22CE", lw=2.0)
    )
    ax.text(9.1, 6.55, "SPI (1 MHz, Mode 1)\n+ DRDY Handshake", fontsize=7.5, fontweight='bold', ha='center', va='bottom', color='#7E22CE',
            bbox=dict(boxstyle='round,pad=0.2', facecolor='#FAF5FF', edgecolor='#E9D5FF', alpha=0.95))

    # 6.4 Internal Data Stream from ESP32-C6 to IoT Block
    ax.annotate(
        "", xy=(11.45, 4.4), xytext=(11.45, 4.7),
        arrowprops=dict(arrowstyle="->", color="#7E22CE", lw=1.8)
    )
    ax.text(11.45, 4.55, "Live Data Stream (1 Hz)", fontsize=7.2, ha='center', va='center', color='#6B21A8',
            bbox=dict(boxstyle='round,pad=0.15', facecolor='#FFFFFF', edgecolor='#D8B4FE', alpha=0.95))

    # 6.5 Power Distribution Arrows (From 5.2V Switching)
    # Up to REF5045 V_IN
    ax.annotate(
        "", xy=(5.2, 2.7), xytext=(3.5, 1.95),
        arrowprops=dict(arrowstyle="->", color="#EA580C", lw=2.0, linestyle='-')
    )
    ax.text(4.25, 2.3, "5.2 V V_IN Rail\n(Clean Switching)", fontsize=7.5, fontweight='bold', ha='center', va='center', color='#C2410C',
            bbox=dict(boxstyle='round,pad=0.18', facecolor='#FFF7ED', edgecolor='#FED7AA', alpha=0.95))

    # Across to Planned LDO Roadmap (Dashed Arrow)
    ax.annotate(
        "", xy=(7.4, 1.27), xytext=(6.4, 1.27),
        arrowprops=dict(arrowstyle="->", color="#059669", lw=1.8, linestyle='--')
    )
    ax.text(6.9, 1.45, "Upgrade Path", fontsize=7.5, fontweight='bold', ha='center', va='bottom', color='#059669',
            bbox=dict(boxstyle='round,pad=0.15', facecolor='#F0FDF4', edgecolor='#A7F3D0', alpha=0.9))

    # Save to file
    out_path = os.path.join('Report_doc', 'image', 'system_block_diagram.png')
    plt.tight_layout()
    plt.savefig(out_path, dpi=300, bbox_inches='tight')
    plt.close()
    print(f"[OK] High-quality system block diagram saved to: {out_path}")

if __name__ == '__main__':
    generate_system_block_diagram()

