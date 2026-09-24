import os
import json
import secrets
import smtplib
from email.message import EmailMessage
import streamlit as st
from streamlit_drawable_canvas import st_canvas
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from PIL import Image

# 網頁基本設定
st.set_page_config(page_title="內政部租賃合約線上簽署系統", layout="wide")
st.title(" 住宅租賃契約書 - 線上合約簽署系統")

# 17 種固定家具設備清單
FURNITURE_ITEMS = [
    ("bed_frame", "床架"), ("mattress", "床墊"), ("wardrobe", "衣櫃"), ("table", "桌子"), ("chair", "椅子"),
    ("sofa", "沙發"), ("tea_table", "茶几"), ("ac", "冷氣機"), ("fridge", "冰箱"), ("washer", "洗衣機"),
    ("tv", "電視機"), ("heater", "熱水器"), ("cooker", "電磁爐"), ("lighting", "燈具"), ("toilet", "馬桶"),
    ("sink", "洗手台"), ("shower", "蓮蓬頭")
]

# 存放每一份合約的資料檔、房東簽名、房東印鑑的資料夾(用短 ID 分開存,彼此不會互相覆蓋)。
DATA_DIR = "contract_data"
os.makedirs(DATA_DIR, exist_ok=True)


def data_path(cid):
    return os.path.join(DATA_DIR, f"{cid}.json")


def sign_path(cid):
    return os.path.join(DATA_DIR, f"{cid}_sign.png")


def seal_path(cid):
    return os.path.join(DATA_DIR, f"{cid}_seal.png")


# 【短網址修復】:原本把所有欄位壓縮成一大串 base64 直接塞進網址,欄位越填越多、網址越來越長,
# 傳到 Line 或某些瀏覽器就可能因為網址過長被截斷或回傳 414 錯誤。
# 改成:資料實際存在伺服器上的一個小檔案裡,網址只帶一組幾個字的短 ID 去對應那個檔案,
# 網址長度固定很短,不會再因為欄位變多而變長。
contract_id = None
decoded_data = {}
if "p" in st.query_params:
    raw_id = st.query_params["p"]
    if isinstance(raw_id, list):
        raw_id = raw_id[0]
    if raw_id:
        contract_id = raw_id.strip()
        if os.path.exists(data_path(contract_id)):
            try:
                with open(data_path(contract_id), "r", encoding="utf-8") as f:
                    decoded_data = json.load(f)
            except Exception:
                decoded_data = {}

is_tenant_view = contract_id is not None

if is_tenant_view and not decoded_data:
    st.error(" 找不到這個連結對應的合約資料,連結可能已失效或不正確,請跟房東確認連結是否傳錯。")
    st.stop()


def get_p(key, default=""):
    """高階字串還原工具:從已解析好的合約資料字典取值,沒有資料時回傳預設值。"""
    val = decoded_data.get(key)
    if val:
        return str(val).strip()
    return default


def send_contract_email(to_email, file_path, tenant_name):
    """簽署完成後,將產出的合約 Word 檔用 Email 自動寄送給房東填寫的信箱。

    需要在部署平台(Streamlit Community Cloud 的話是 App -> Settings -> Secrets)設定 SMTP 帳號密碼,格式如下:

        [smtp]
        host = "smtp.gmail.com"
        port = 465
        user = "your_gmail@gmail.com"
        password = "xxxxxxxxxxxxxxxx"   # Gmail 請用「應用程式密碼」,不是登入密碼
        sender = "your_gmail@gmail.com" # 可省略,預設會用 user

    尚未設定 Secrets 時不會讓整個流程失敗,只會提示合約已產生但信件未寄出。
    """
    if not to_email:
        return False
    if "smtp" not in st.secrets:
        st.warning(" 尚未設定寄信用的 SMTP 資訊(Secrets),合約已產生但 Email 未寄出,請至部署平台補上設定。")
        return False
    try:
        cfg = st.secrets["smtp"]
        msg = EmailMessage()
        msg["Subject"] = f"【自動通知】承租人「{tenant_name}」已完成合約線上簽署"
        msg["From"] = cfg.get("sender", cfg["user"])
        msg["To"] = to_email
        msg.set_content(f"您好,\n\n承租人「{tenant_name}」已完成線上簽署,合約檔案請見附件。\n\n(此為系統自動寄送之郵件,請勿直接回覆)")
        with open(file_path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
                filename=os.path.basename(file_path),
            )
        with smtplib.SMTP_SSL(cfg["host"], int(cfg.get("port", 465))) as smtp:
            smtp.login(cfg["user"], cfg["password"])
            smtp.send_message(msg)
        return True
    except Exception as e:
        st.warning(f" Email 寄送失敗(合約檔案仍可正常下載):{e}")
        return False


if is_tenant_view:
    st.info("【房客簽署頁面】以下資料由房東填寫並鎖定,僅供核對、不可修改。請確認無誤後,在最下方「承租人(房客)手寫簽名」欄位親筆簽名,再點擊底部按鈕完成簽署並下載合約。")
else:
    st.write("【房東專區】:填完合約細節並手寫簽名後,點擊底部即可生成『全資料同步連結』傳給房客,房客免註冊登入即可補簽。")

# ==================== 建立網頁左、中、右三欄版面 ====================
col1, col2, col3 = st.columns(3)

# 1. 基本與租期資料
with col1:
    st.header("1. 基本與租期資料")
    landlord_name = st.text_input("出租人姓名 *", value=get_p("l_name"), disabled=is_tenant_view)
    landlord_id = st.text_input("出租人身分證字號", value=get_p("l_id"), disabled=is_tenant_view)
    landlord_email = st.text_input("房東收件 Email(簽署完成後自動寄送合約)", value=get_p("l_email"), disabled=is_tenant_view)
    tenant_name = st.text_input("承租人姓名 *", value=get_p("t_name"), disabled=is_tenant_view)
    tenant_id = st.text_input("承租人身分證字號", value=get_p("t_id"), disabled=is_tenant_view)
    tenant_phone = st.text_input("承租人電話", value=get_p("t_phone"), disabled=is_tenant_view)
    tenant_address = st.text_input("承租人戶籍地址", value=get_p("t_addr"), disabled=is_tenant_view)
    address = st.text_input("租賃房屋地址 *", value=get_p("addr"), disabled=is_tenant_view)
    rent_amount = st.text_input("每月租金 (元)", value=get_p("rent"), disabled=is_tenant_view)
    deposit_amount = st.text_input("押金金額 (元)", value=get_p("dep"), disabled=is_tenant_view)

    st.subheader("租期時間")
    c_s1, c_s2, c_s3 = st.columns(3)
    s_year = c_s1.text_input("開始年(民國)", value=get_p("sy", "115"), disabled=is_tenant_view)
    s_month = c_s2.text_input("開始月", value=get_p("sm", "1"), disabled=is_tenant_view)
    s_day = c_s3.text_input("開始日", value=get_p("sd", "1"), disabled=is_tenant_view)

    c_e1, c_e2, c_e3 = st.columns(3)
    e_year = c_e1.text_input("結束年(民國)", value=get_p("ey", "116"), disabled=is_tenant_view)
    e_month = c_e2.text_input("結束月", value=get_p("em", "1"), disabled=is_tenant_view)
    e_day = c_e3.text_input("結束日", value=get_p("ed", "1"), disabled=is_tenant_view)

    car_idx = 1 if get_p("car") == "yes" else 0
    car_option = st.radio("汽車位需求", ["無汽車位", "有汽車位"], index=car_idx, disabled=is_tenant_view)
    moto_idx = 1 if get_p("moto") == "yes" else 0
    moto_option = st.radio("機車位需求", ["無機車位", "有機車位"], index=moto_idx, disabled=is_tenant_view)

# 2. 租賃期間費用約定
with col2:
    st.header("2. 租賃期間費用約定")
    mgmt_list = ["出租人負擔", "承租人負擔", "其他約定"]
    mgmt_idx = mgmt_list.index(get_p("m_p")) if get_p("m_p") in mgmt_list else 0
    mgmt_pay = st.radio("管理費負擔方", mgmt_list, index=mgmt_idx, key="mgmt", disabled=is_tenant_view)
    fee_mgmt_house = st.text_input("住宅管理費/月 (元)", value=get_p("m_h", "0"), disabled=is_tenant_view)
    fee_mgmt_car = st.text_input("車位管理費/月 (元)", value=get_p("m_c", "0"), disabled=is_tenant_view)
    txt_mgmt_other = st.text_input("管理費其他約定說明", value=get_p("m_o"), disabled=is_tenant_view)

    water_list = ["出租人負擔", "承租人負擔", "其他約定"]
    water_idx = water_list.index(get_p("w_p")) if get_p("w_p") in water_list else 1
    water_pay = st.radio("水費負擔方", water_list, index=water_idx, key="water", disabled=is_tenant_view)
    txt_water_other = st.text_input("水費其他約定說明", value=get_p("w_o"), disabled=is_tenant_view)

    elec_list = ["出租人負擔", "承租人負擔 (依當期平均電價)", "承租人負擔 (固定每度元)", "非度數計費其他約定"]
    elec_idx = elec_list.index(get_p("e_p")) if get_p("e_p") in elec_list else 1
    elec_pay = st.radio("電費計費方式", elec_list, index=elec_idx, key="elec", disabled=is_tenant_view)
    fee_elec_rate = st.text_input("固定每度電費 (元)", value=get_p("e_r", "0"), disabled=is_tenant_view)
    txt_elec_other = st.text_input("電費other約定說明", value=get_p("e_o"), disabled=is_tenant_view)

    gas_list = ["出租人負擔", "承租人負擔", "其他約定"]
    gas_idx = gas_list.index(get_p("g_p")) if get_p("g_p") in gas_list else 1
    gas_pay = st.radio("瓦斯費負擔方", gas_list, index=gas_idx, key="gas", disabled=is_tenant_view)
    txt_gas_other = st.text_input("瓦斯費其他約定說明", value=get_p("g_o"), disabled=is_tenant_view)

    net_list = ["出租人負擔", "承租人負擔", "其他約定"]
    net_idx = net_list.index(get_p("n_p")) if get_p("n_p") in net_list else 0
    net_pay = st.radio("網路費負擔方", net_list, index=net_idx, key="net", disabled=is_tenant_view)
    txt_net_other = st.text_input("網路費其他約定說明", value=get_p("n_o"), disabled=is_tenant_view)
    txt_other_fee = st.text_input("請輸入其他費用說明", value=get_p("oth_f"), disabled=is_tenant_view)

# 3. 設備清單、點收物品與手寫簽名
with col3:
    st.header("3. 附屬設備、點收與簽名")
    st.markdown("** 附屬設備清單**")
    fur_context = {}
    with st.expander("點擊展開常見家具清單"):
        for key, name in FURNITURE_ITEMS:
            c_f1, c_f2 = st.columns(2)
            default_chk = True if get_p(f"f_{key}") == "1" else False
            is_checked = c_f1.checkbox(name, value=default_chk, key=f"cb_{key}", disabled=is_tenant_view)
            num = c_f2.text_input("數量", value=get_p(f"fn_{key}", "1"), key=f"num_{key}", disabled=is_tenant_view)
            fur_context[key] = (is_checked, num)

    chk_other_f = st.checkbox("其他自訂設備", value=True if get_p("f_oth") == "1" else False, disabled=is_tenant_view)
    textarea_other = st.text_area("請輸入自訂家具備註", value=get_p("f_txt"), height=60, disabled=is_tenant_view)

    st.write("---")
    st.markdown("** 承租人點收物品**")
    c_k1, c_k2 = st.columns(2)
    chk_key_house = c_k1.checkbox("房屋鑰匙", value=True if get_p("k_h") == "1" else False, disabled=is_tenant_view)
    num_key_house = c_k2.text_input("房屋鑰匙數量", value=get_p("kn_h", "1"), disabled=is_tenant_view)
    chk_token = c_k1.checkbox("感應磁扣(卡)", value=True if get_p("k_t") == "1" else False, disabled=is_tenant_view)
    num_token = c_k2.text_input("感應磁扣數量", value=get_p("kn_t", "1"), disabled=is_tenant_view)
    chk_key_mail = c_k1.checkbox("信箱鑰匙", value=True if get_p("k_m") == "1" else False, disabled=is_tenant_view)
    num_key_mail = c_k2.text_input("信箱鑰匙數量", value=get_p("kn_m", "1"), disabled=is_tenant_view)
    chk_remote = c_k1.checkbox("車庫遙控器", value=True if get_p("k_r") == "1" else False, disabled=is_tenant_view)
    num_remote = c_k2.text_input("車庫遙控器數量", value=get_p("kn_r", "1"), disabled=is_tenant_view)

    st.write("---")
    st.markdown("** 出租人(房東)手寫簽名**")
    if is_tenant_view:
        # 房客檢視模式:房東簽名鎖定為唯讀,顯示房東生成連結時存下的簽名圖檔,不能再畫。
        canvas_l = None
        if os.path.exists(sign_path(contract_id)):
            st.image(sign_path(contract_id), width=280)
        else:
            st.caption("(尚未取得房東簽名圖檔)")
    else:
        canvas_l = st_canvas(fill_color="rgba(255,255,255,0)", stroke_width=3, stroke_color="#000000", background_color="#FFFFFF", height=100, width=280, drawing_mode="freedraw", key="canvas_l", return_image_data=True)

    # 【新增】出租人印鑑:緊接在房東簽名後面,讓房東可以掃描/拍照上傳印鑑圖片,一起附加到合約裡。
    st.markdown("** 出租人印鑑(可選,掃描或拍照上傳)**")
    if is_tenant_view:
        landlord_seal_file = None
        if os.path.exists(seal_path(contract_id)):
            st.image(seal_path(contract_id), width=150)
        else:
            st.caption("(房東未上傳印鑑圖檔)")
    else:
        landlord_seal_file = st.file_uploader("上傳印鑑圖片 (png/jpg)", type=["png", "jpg", "jpeg"], key="landlord_seal_upload")

    st.markdown("** 承租人(房客)手寫簽名**")
    canvas_t = st_canvas(fill_color="rgba(255,255,255,0)", stroke_width=3, stroke_color="#000000", background_color="#FFFFFF", height=100, width=280, drawing_mode="freedraw", key="canvas_t", return_image_data=True)
# ==================== 底部功能按鈕區 ====================
st.write("---")
b_col1, b_col2 = st.columns(2)

with b_col1:
    st.subheader("【房東步驟 1】:產生全資料同步網址")
    if is_tenant_view:
        st.caption("此區僅供房東使用,房客檢視模式下已隱藏,以避免資料被覆蓋。")
    else:
        if st.button(" 一鍵生成房客簽名連結", use_container_width=True):
            # 每次生成連結都給一組新的短 ID,資料、房東簽名、房東印鑑都分開存成這組 ID 專屬的檔案,
            # 不同筆合約之間不會互相覆蓋(舊版做法是全部共用同一個檔名,多筆合約同時進行時會互相蓋掉)。
            new_id = secrets.token_urlsafe(6)

            if canvas_l is not None and canvas_l.image_data is not None and canvas_l.image_data.any():
                Image.fromarray(canvas_l.image_data.astype('uint8'), 'RGBA').save(sign_path(new_id))

            if landlord_seal_file is not None:
                landlord_seal_file.seek(0)
                Image.open(landlord_seal_file).convert("RGBA").save(seal_path(new_id))

            # 大打包所有資料欄位
            payload = {
                "l_name": landlord_name, "l_id": landlord_id, "l_email": landlord_email, "t_name": tenant_name, "t_id": tenant_id, "t_phone": tenant_phone,
                "t_addr": tenant_address, "addr": address, "rent": rent_amount, "dep": deposit_amount,
                "sy": s_year, "sm": s_month, "sd": s_day, "ey": e_year, "em": e_month, "ed": e_day,
                "car": "yes" if car_option == "有汽車位" else "no", "moto": "yes" if moto_option == "有機車位" else "no",
                "m_p": mgmt_pay, "m_h": fee_mgmt_house, "m_c": fee_mgmt_car, "m_o": txt_mgmt_other,
                "w_p": water_pay, "w_o": txt_water_other, "e_p": elec_pay, "e_r": fee_elec_rate, "e_o": txt_elec_other,
                "g_p": gas_pay, "g_o": txt_gas_other, "n_p": net_pay, "n_o": txt_net_other, "oth_f": txt_other_fee,
                "k_h": "1" if chk_key_house else "0", "kn_h": num_key_house, "k_t": "1" if chk_token else "0", "kn_t": num_token,
                "k_m": "1" if chk_key_mail else "0", "kn_m": num_key_mail, "k_r": "1" if chk_remote else "0", "kn_r": num_remote,
                "f_oth": "1" if chk_other_f else "0", "f_txt": textarea_other
            }
            for k, (is_chk, num) in fur_context.items():
                payload[f"f_{k}"] = "1" if is_chk else "0"
                payload[f"fn_{k}"] = num

            with open(data_path(new_id), "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False)

            # 【短網址】:網址只帶這組短 ID,不再把整包資料塞進網址,徹底解決網址過長 / 414 的問題。
            share_url = f"https://gxbnexkrg8ixs4pe8s4ywh.streamlit.app/?p={new_id}"

            st.success(" 全資料同步網址生成成功!請點擊下方代碼框右上角的『Copy』一鍵複製傳給房客(傳 Line 100% 免登入、全欄位自動打勾帶入):")
            st.code(share_url, language="text")

with b_col2:
    if is_tenant_view:
        st.subheader("【房客步驟】:確認資料並完成簽署")
        btn_label = " 確認無誤,完成簽署並產生合約"
    else:
        st.subheader("【房客與房東步驟 2】:雙方簽完名後生成下載")
        btn_label = " 線上生成合約文件"
    if st.button(btn_label, use_container_width=True):
        if not landlord_name or not tenant_name or not address:
            st.error(" 錯誤:『出租人』、『承租人姓名』與『租賃房屋地址』為必填欄位!")
        else:
            with st.spinner("系統正在處理資料,請稍候..."):
                try:
                    context = {
                        "landlord_name": landlord_name, "landlord_id": landlord_id, "tenant_name": tenant_name, "tenant_id": tenant_id,
                        "tenant_phone": tenant_phone, "tenant_address": tenant_address, "address": address,
                        "rent_amount": rent_amount, "deposit_amount": deposit_amount,
                        "start_year": s_year, "start_month": s_month, "start_day": s_day,
                        "end_year": e_year, "end_month": e_month, "end_day": e_day,
                        "chk_car_yes": "■" if car_option == "有汽車位" else "□", "chk_car_no": "■" if car_option == "無汽車位" else "□",
                        "chk_moto_yes": "■" if moto_option == "有機車位" else "□", "chk_moto_no": "■" if moto_option == "無機車位" else "□",
                        "chk_mgmt_landlord": "■" if mgmt_pay == "出租人負擔" else "□", "chk_mgmt_tenant": "■" if mgmt_pay == "承租人負擔" else "□",
                        "chk_mgmt_other": "■" if mgmt_pay == "其他約定" else "□", "fee_mgmt_house": fee_mgmt_house, "fee_mgmt_car": fee_mgmt_car, "txt_mgmt_other": txt_mgmt_other,
                        "chk_water_landlord": "■" if water_pay == "出租人負擔" else "□", "chk_water_tenant": "■" if water_pay == "承租人負擔" else "□", "chk_water_other": "■" if water_pay == "其他約定" else "□", "txt_water_other": txt_water_other,
                        "chk_elec_landlord": "■" if elec_pay == "出租人負擔" else "□", "chk_elec_tenant_avg": "■" if elec_pay == "承租人負擔 (依當期平均電價)" else "□", "chk_elec_tenant_fixed": "■" if elec_pay == "承租人負擔 (固定每度元)" else "□",
                        "chk_elec_other": "■" if elec_pay == "非度數計費其他約定" else "□",
                        "fee_elec_rate": fee_elec_rate, "txt_elec_other": txt_elec_other,
                        "chk_gas_landlord": "■" if gas_pay == "出租人負擔" else "□", "chk_gas_tenant": "■" if gas_pay == "承租人負擔" else "□", "chk_gas_other": "■" if gas_pay == "其他約定" else "□", "txt_gas_other": txt_gas_other,
                        "chk_net_landlord": "■" if net_pay == "出租人負擔" else "□", "chk_net_tenant": "■" if net_pay == "承租人負擔" else "□", "chk_net_other": "■" if net_pay == "其他約定" else "□",
                        "txt_net_other": txt_net_other,
                        "txt_other_fee": txt_other_fee,
                        "chk_key_house": "■" if chk_key_house else "□", "num_key_house": num_key_house if chk_key_house else "",
                        "chk_token": "■" if chk_token else "□", "num_token": num_token if chk_token else "",
                        "chk_key_mail": "■" if chk_key_mail else "□", "num_key_mail": num_key_mail if chk_key_mail else "",
                        "chk_remote": "■" if chk_remote else "□", "num_remote": num_remote if chk_remote else ""
                    }
                    for k, (is_chk, num) in fur_context.items():
                        context[f"chk_{k}"] = "■" if is_chk else "□"
                        context[f"num_{k}"] = num if is_chk else ""
                    context["chk_other"] = "■" if chk_other_f else "□"
                    context["textarea_other"] = textarea_other if chk_other_f else ""

                    if not os.path.exists("template.docx"):
                        st.error("找不到 template.docx 檔案!")
                    else:
                        doc = DocxTemplate("template.docx")

                        # 雲端背景雙簽名智慧合體
                        # canvas_l 在房客唯讀模式下是 None(房東簽名畫布沒有渲染),先判斷是否存在再讀 image_data。
                        if canvas_l is not None and canvas_l.image_data is not None and canvas_l.image_data.any():
                            Image.fromarray(canvas_l.image_data.astype('uint8'), 'RGBA').save("wl.png")
                            context["landlord_sign"] = InlineImage(doc, "wl.png", width=Inches(1.2))
                        elif contract_id and os.path.exists(sign_path(contract_id)):
                            context["landlord_sign"] = InlineImage(doc, sign_path(contract_id), width=Inches(1.2))
                        else:
                            context["landlord_sign"] = ""

                        # 出租人印鑑:房東當場上傳的優先,否則用產生連結時存下的那份。
                        # 印鑑實體尺寸通常只有 1.5~2 公分左右,這裡設定寬度 0.7 吋(約 1.8 公分),
                        # 如果插進 Word 後還是覺得太大或太小,把下面兩個 SEAL_WIDTH_INCHES 改數字就好
                        # (例如想再小一點改成 0.5,想大一點改成 0.9)。
                        # 【範本提醒】要讓印鑑真的出現在 Word 檔裡,template.docx 需要在房東簽名欄位後方
                        # 加上 {{landlord_seal}} 這個合併欄位,以及在對應位置加上 {{landlord_id}} 顯示出租人身分證字號。
                        SEAL_WIDTH_INCHES = 0.7
                        if landlord_seal_file is not None:
                            landlord_seal_file.seek(0)
                            Image.open(landlord_seal_file).convert("RGBA").save("wseal.png")
                            context["landlord_seal"] = InlineImage(doc, "wseal.png", width=Inches(SEAL_WIDTH_INCHES))
                        elif contract_id and os.path.exists(seal_path(contract_id)):
                            context["landlord_seal"] = InlineImage(doc, seal_path(contract_id), width=Inches(SEAL_WIDTH_INCHES))
                        else:
                            context["landlord_seal"] = ""

                        if canvas_t.image_data is not None and canvas_t.image_data.any():
                            Image.fromarray(canvas_t.image_data.astype('uint8'), 'RGBA').save("wt.png")
                            context["tenant_sign"] = InlineImage(doc, "wt.png", width=Inches(1.2))
                        else:
                            context["tenant_sign"] = ""

                        time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                        out_word = f"住宅租賃契約書_{tenant_name}_{time_str}.docx"
                        doc.render(context)
                        doc.save(out_word)

                        if os.path.exists("wl.png"): os.remove("wl.png")
                        if os.path.exists("wt.png"): os.remove("wt.png")
                        if os.path.exists("wseal.png"): os.remove("wseal.png")

                        st.success(" 線上合約已成功產出!請點擊下載您的合約檔案:")
                        with open(out_word, "rb") as wf:
                            st.download_button(label=" 下載最終雙方簽署合約 Word 檔案 (.docx)", data=wf, file_name=f"住宅租賃契約書_{tenant_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

                        # 簽署完成後,自動把合約寄一份到房東填寫的 Email(尚未設定 SMTP Secrets 時只會提示,不影響下載)。
                        if landlord_email:
                            with st.spinner(f"正在將合約寄送至 {landlord_email} ..."):
                                if send_contract_email(landlord_email, out_word, tenant_name):
                                    st.success(f" 已自動將合約寄送至 {landlord_email}")
                except Exception as e:
                    st.error(f"生成失敗,原因:{str(e)}")
