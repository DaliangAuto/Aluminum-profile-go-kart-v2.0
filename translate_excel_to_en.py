# -*- coding: utf-8 -*-
"""Translate Chinese Excel BOM files to English versions."""
import openpyxl
import os

BASE = r"e:\showcase\手搓智驾1_底盘\Aluminum-profile-go-kart-v2.0"
os.chdir(BASE)

TRANS = {
    "铝型材小车物料清单V2.0": "Aluminum Profile Go-Kart BOM V2.0",
    "智能驾驶物料清单": "Autopilot Parts List",
    "日期": "Date",
    "作者": "Author",
    "大亮Auto": "Daliang Auto",
    "序号": "Item No.",
    "物料号": "Part No.",
    "类别": "Category",
    "名称": "Description",
    "型号": "Model",
    "规格": "Specification",
    "数量": "Qty",
    "单位": "Unit",
    "备注": "Remarks",
    "图片": "Picture",
    "购买链接": "Purchase Link",
    "说明：": "Notes:",
    "更新历史：": "Revision History:",
    "根": "pcs",
    "套": "set",
    "个": "pcs",
    "只": "pcs",
    "块": "pcs",
    "台": "unit",
    "铝型材框架及配件": "Aluminum profile frame and accessories",
    "主框架纵梁": "Main frame longitudinal beam",
    "主框架横梁": "Main frame cross beam",
    "主框架立柱": "Main frame upright",
    "欧标3030铝型材，厚度2.2mm": "EU 3030 aluminum profile, thickness 2.2mm",
    "长度900mm": "Length 900mm",
    "长度130mm": "Length 130mm",
    "长度80mm": "Length 80mm",
    "长度400mm": "Length 400mm",
    "长度480mm": "Length 480mm",
    "欧标3060铝型材，厚度2.2mm": "EU 3060 aluminum profile, thickness 2.2mm",
    "长度290mm": "Length 290mm",
    "电池底托": "Battery base tray",
    "长度190mm": "Length 190mm",
    "座椅纵梁": "Seat longitudinal beam",
    "长度300mm": "Length 300mm",
    "座椅横梁": "Seat cross beam",
    "座椅立柱": "Seat upright",
    "方向盘立柱": "Steering column",
    "方向盘横撑": "Steering cross brace",
    "脚踏横杆": "Footrest cross bar",
    "脚踏纵梁": "Footrest longitudinal beam",
    "T型螺母+内六角螺栓": "T-nut + hex socket bolt",
    "适配欧标3030型材": "For EU 3030 profile",
    "角码": "Corner bracket",
    "L形连接件": "L-shaped connector",
    "T形连接件": "T-shaped connector",
    "铝合金L形角码": "Aluminum L-shaped corner bracket",
    "活动铰链": "Hinge",
    "悬架，前轮，转向总成": "Suspension, front wheel, steering assembly",
    "悬架，前轮": "Suspension, front wheel",
    "转向总成": "Steering assembly",
    "刹车踏板及刹车盒": "Brake pedal and brake box",
    "刹车卡钳": "Brake caliper",
    "成品": "Finished product",
    "轮毂电机": "Hub motor",
    "方向盘折叠气弹簧": "Steering fold gas spring",
    "坐垫或者座椅": "Seat cushion or seat",
    "自行车后坐垫": "Bicycle pillion seat",
    "控制板+脚踏板": "Control board + foot pedal",
    "锂电池或者铅酸电池": "Li-ion or lead-acid battery",
    "电池": "Battery",
    "【6.5寸8寸10寸平衡车电机24v36v通用，正常可用，成色】": "【6.5/8/10 in. balance-scooter motor 24V/36V, usable】",
    "36V48Ⅴ卡丁车扭扭车主板控制板 平衡车改卡丁车双驱控制器": "36V/48V go-kart twist main board, balance-to-kart dual-drive controller",
    "厚度不同，价格不同，太薄可能强度不够，我用的是2.2厚度的.2.0强度应该也是够的，会便宜些，自己选择。": "Thickness affects price; too thin may lack strength. 2.2mm works; 2.0mm may suffice and is cheaper. Choose as needed.",
    "如果需要折叠方向盘，一端需要攻丝, 攻M6*20的": "If folding steering is needed, one end requires tapping for M6*20",
    "建议配平垫，否则容易松": "Suggest flat washer to prevent loosening",
    "10寸，轴径16mm": "10 inch, shaft diameter 16mm",
    "可以买二手，直流无刷轮毂电机，就是电动平衡车上的那种，电压一般是36V": "Used units OK; BLDC hub motor (e.g. from electric balance scooter), typically 36V",
    "缸径18mm，20公斤压力": "Bore 18mm, 20kg force",
    "中心距280mm，两端T型接头": "Center distance 280mm, T-joint both ends",
    "可直接买座椅或者自行车后坐垫，座椅安装孔横向距离16cm": "Can use seat or bicycle pillion; seat mounting holes spacing 16cm",
    "如果不会写程序，建议买成品板，买二手卡丁车控制板，带2个脚踏板，几十块钱很便宜，注意看电压要匹配，注意电源线端子是否是XT60，公母头要匹配，一个是公头，一个是母头。\n\n如果会写程序，可以用单片机+直流无刷电机控制器实现": "If no programming: buy finished control board (used go-kart board with 2 pedals, cheap). Match voltage; check XT60 connectors (male/female pair). If programming: MCU + BLDC controller.",
    "36V 15Ah，低速模式可以玩1-2个小时，前提是电池别虚标，千万别贪便宜": "36V 15Ah, 1–2h runtime in slow mode if battery is not overrated; avoid cheap low-quality units.",
    "宽度不超110mm，高度不超110": "Width ≤110mm, height ≤110mm",
    "电池电压需要与电机的电压匹配，如果轮毂电机36v，电池也要是36v，注意一起配充电器，如果买锂电池必须配相同电压锂电池充电器，最好同一家买充电器。\n如果买的成品控制板，注意电池接头的端子要是XT60": "Battery voltage must match motor (e.g. 36V hub = 36V battery). Include matching charger; Li-ion needs same-voltage Li-ion charger, preferably from same seller. If using finished control board, ensure battery connector is XT60.",
    "增加轮毂电机链接": "Added hub motor link",
    "可选，也可采用电路板自带的电机电刹，但效果不如机械刹车效果好，自行选择": "Optional; electric motor brake on the control board is possible but mechanical braking is usually more effective. Choose as needed.",
    "安装孔距需要与羊角匹配，孔距51mm": "Mount hole spacing must match knuckle; 51mm spacing",
    "左右各一只，同上配套，可选": "One per side, matching set as above, optional",
    "控制板": "Control board",
    "算力模块": "Compute module",
    "无刷电机驱动板": "BLDC motor driver board",
    "转向伺服电机": "Steering servo motor",
    "减速机": "Reducer / gearbox",
    "刹车电缸": "Brake electric cylinder",
    "摄像头": "Camera",
    "24v电源": "24V power supply",
    "19v电源": "19V power supply",
    "pwm,有霍尔驱动板": "PWM, with Hall sensor driver",
    "时代超群60法兰24v驱控一体": "Shidai Chaoqun 60 flange 24V integrated drive",
    "200w, 3000转，CAN": "200W, 3000rpm, CAN",
    "法兰不可超60，可选择57或者或者42直流伺服电机，轴径不超14mm才能匹配转向机，注意：更换伺服电机，控制板固件需根据通信方式或协议修改代码": "Flange ≤60; may use 57 or 42 BLDC servo. Shaft ≤14mm to match steering. Changing servo requires firmware/code update for protocol.",
    "行星减速机": "Planetary reducer",
    "速比至少1:10": "Ratio ≥1:10",
    "法兰不超60": "Flange ≤60",
    "直流电缸": "DC electric cylinder",
    "24v，行程20mm，速度60mm/s，推力200N": "24V, stroke 20mm, speed 60mm/s, force 200N",
    "拉动刹车线，配合外采可兼容手自动的线刹盒": "Pulls brake cable; works with external manual/auto compatible brake box",
    "IMX219": "IMX219",
    "转向伺服电机和刹车电缸供电": "Powers steering servo and brake cylinder",
    "算力模块供电": "Powers compute module",
    "非外购，由Daliang开发": "In-house, developed by Daliang",
    "【淘宝】": "[Taobao] ",
    "【闲鱼】": "[Xianyu] ",
    "退货运费险": "Return shipping insurance",
    "7天无理由退货": "7-day no-questions return",
    "点击链接直接打开 或者 淘宝搜索直接打开": "Click to open or search on Taobao",
    "点击链接直接打开": "Click to open",
    "快来捡漏": "Deal alert",
    "欧标铝型材2020铝合金型材4040流水线框架3030鱼缸架60工作台橱柜": "EU aluminum profile 2020/3030/4040",
    "气弹簧气撑杆床用支撑杆橱柜广告牌上翻门厢货车用气压液压杆": "Gas spring, strut, cabinet/bed support",
}

DISCLAIMER_MECH = """Please read the following notes in full before downloading, using or sharing this material:

1. Overall dimensions are compact for portability (e.g. VW ID.3 trunk). Adjust for your vehicle and use case.
2. No materials sold, no fees. Per platform rules, no purchase links provided. Search and buy from public channels based on images and specs.
3. For assembly issues, use the community group. Q&A and livestreams when available. Content is experience sharing only, not technical guarantee, safety commitment or commercial service.
4. Data or drawings may contain errors; cutting, modification or supplement may be required.
5. Daliang Auto publishes this under CC BY-NC 4.0. Personal learning, modification and non-commercial sharing allowed with attribution. Commercial use prohibited.
6. For learning and hobby only. Any machining, modification, installation, testing or use (including structure, mechanical/electrical safety, driving) is at user's risk. Daliang Auto assumes no liability. Comply with local laws.
7. Downloading, using, modifying or sharing constitutes acceptance of these terms."""


def translate_cell(val):
    if val is None or not isinstance(val, str):
        return val
    s = str(val).strip()
    if not s:
        return val
    if s in TRANS:
        return TRANS[s]
    if "请务必在下载、使用或传播本资料前，完整阅读以下说明" in s and "大亮Auto" in s:
        return DISCLAIMER_MECH
    result = s
    for zh, en in sorted(TRANS.items(), key=lambda x: -len(x[0])):
        if zh in result and zh not in ("大亮Auto", "Daliang"):
            result = result.replace(zh, en)
    return result


def process_excel(src, dst):
    wb = openpyxl.load_workbook(src, data_only=False)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    cell.value = translate_cell(cell.value)
    wb.save(dst)
    print(f"Saved: {dst}")


def main():
    pairs = [
        ("Mechanical system part list_CN.xlsx", "Mechanical system part list_EN.xlsx"),
        ("Electronics & Compute part list_CN.xlsx", "Electronics & Compute part list_EN.xlsx"),
    ]
    for src, dst in pairs:
        if os.path.exists(src):
            process_excel(src, dst)
        else:
            print(f"Skip (not found): {src}")


if __name__ == "__main__":
    main()
