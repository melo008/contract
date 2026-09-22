
import os
import urllib.parse
import streamlit as st
from streamlit_drawable_canvas import st_canvas
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches

from PIL import Image

# 網頁基本設定
st.set_page_config(page_title="內政部租賃合約線上簽署系統", layout="wide")
st.title(" 住宅租賃契約書 - 線上合約簽署系統")
st.write("【房東專區】：填完資料後可於底部產生『專屬簽名連結』傳給房客；【房客專區】：核對資料後，於右側手寫簽名即可。")

# 讀取網址上的參數（如果是房客點開連結，會自動帶入資料）
query_params = st.query_params

# 輔助函式：讀取網址參數
def get_param(key, default=""):
    return query_params.get(key, default)

# ==================== 建立網頁左、中、右三欄版面 ====================
col1, col2, col3 = st.columns(3)

# 1. 基本與租期資料
with col1:
    st.header("1. 基本與租期資料")
    landlord_name = st.text_input("出租人姓名 *", value=get_param("l_name"))
    tenant_name = st.text_input("承租人姓名 *", value=get_param("t_name"))
    tenant_id = st.text_input("承租人身分證字號", value=get_param("t_id"))
    tenant_phone = st.text_input("承租人電話", value=get_param("t_phone"))
    tenant_address = st.text_input("承租人戶籍地址", value=get_param("t_addr"))
    address = st.text_input("租賃房屋地址 *", value=get_param("addr"))
    rent_amount = st.text_input("每月租金 (元)", value=get_param("rent"))
    deposit_amount = st.text_input("押金金額 (元)", value=get_param("dep"))
    
    st.subheader("租期時間")
    c_s1, c_s2, c_s3 = st.columns(3)
    s_year = c_s1.text_input("開始年(民國)", value=get_param("sy", "115"))
    s_month = c_s2.text_input("開始月", value=get_param("sm", "1"))
    s_day = c_s3.text_input("開始日", value=get_param("sd", "1"))
    
    c_e1, c_e2, c_e3 = st.columns(3)
    e_year = c_e1.text_input("結束年(民國)", value=get_param("ey", "116"))
    e_month = c_e2.text_input("結束月", value=get_param("em", "1"))
    e_day = c_e3.text_input("結束日", value=get_param("ed", "1"))
    
    car_idx = 1 if get_param("car") == "yes" else 0
    car_option = st.radio("汽車位需求", ["無汽車位", "有汽車位"], index=car_idx)
    moto_idx = 1 if get_param("moto") == "yes" else 0
    moto_option = st.radio("機車位需求", ["無機車位", "有機車位"], index=moto_idx)

# 2. 租賃期間費用約定
with col2:
    st.header("2. 租賃期間費用約定")
    
    mgmt_list = ["出租人負擔", "承租人負擔", "其他約定"]
    mgmt_idx = mgmt_list.index(get_param("m_p")) if get_param("m_p") in mgmt_list else 0
    mgmt_pay = st.radio("管理費負擔方", mgmt_list, index=mgmt_idx, key="mgmt")
    fee_mgmt_house = st.text_input("住宅管理費/月 (元)", value=get_param("m_h", "0"))
    fee_mgmt_car = st.text_input("車位管理費/月 (元)", value=get_param("m_c", "0"))
    txt_mgmt_other = st.text_input("管理費其他約定說明", value=get_param("m_o"))
    
    water_list = ["出租人負擔", "承租人負擔", "其他約定"]
    water_idx = water_list.index(get_param("w_p")) if get_param("w_p") in water_list else 1
    water_pay = st.radio("水費負擔方", water_list, index=water_idx, key="water")
    txt_water_other = st.text_input("水費其他約定說明", value=get_param("w_o"))
    
    elec_list = ["出租人負擔", "承租人負擔 (依當期平均電價)", "承租人負擔 (固定每度元)", "非度數計費其他約定"]
    elec_idx = elec_list.index(get_param("e_p")) if get_param("e_p") in elec_list else 1
    elec_pay = st.radio("電費計費方式", elec_list, index=elec_idx, key="elec")
    fee_elec_rate = st.text_input("固定每度電費 (元)", value=get_param("e_r", "0"))
    txt_elec_other = st.text_input("電費其他約定說明", value=get_param("e_o"))
    
    gas_list = ["出租人負擔", "承租人負擔", "其他約定"]
    gas_idx = gas_list.index(get_param("g_p")) if get_param("g_p") in gas_list else 1
    gas_pay = st.radio("瓦斯費負擔方", gas_list, index=gas_idx, key="gas")
    txt_gas_other = st.text_input("瓦斯費其他約定說明", value=get_param("g_o"))
    
    net_list = ["出租人負擔", "承租人負擔", "其他約定"]
    net_idx = net_list.index(get_param("n_p")) if get_param("n_p") in net_list else 0
    net_pay = st.radio("網路費負擔方", net_list, index=net_idx, key="net")
    txt_net_other = st.text_input("網路費其他約定說明", value=get_param("n_o"))
    
    txt_other_fee = st.text_input("請輸入其他費用說明", value=get_param("oth_f"))

# 3. 設備清單、點收物品與手寫簽名板設定
with col3:
    st.header("3. 附屬設備、點收與簽名")
    
    furniture_items = [
        ("bed_frame", "床架"), ("mattress", "床墊"), ("wardrobe", "衣櫃"), ("table", "桌子"), ("chair", "椅子"),
        ("sofa", "沙發"), ("tea_table", "茶几"), ("ac", "冷氣機"), ("fridge", "冰箱"), ("washer", "洗衣機"),
        ("tv", "電視機"), ("heater", "熱水器"), ("cooker", "電磁爐"), ("lighting", "燈具"), ("toilet", "馬桶"),
        ("sink", "洗手台"), ("shower", "蓮蓬頭")
    ]
    
    st.markdown("** 附屬設備清單**")
    fur_context = {}
    with st.expander("點擊展開常見家具清單"):
        for key, name in furniture_items:
            c_f1, c_f2 = st.columns(2)
            default_chk = True if get_param(f"f_{key}") == "1" else False
            is_checked = c_f1.checkbox(name, value=default_chk, key=f"cb_{key}")
            num = c_f2.text_input("數量", value=get_param(f"fn_{key}", "1"), key=f"num_{key}")
            fur_context[key] = (is_checked, num)
            
    chk_other_f = st.checkbox("其他自訂設備", value=True if get_param("f_oth")=="1" else False)
    textarea_other = st.text_area("請輸入自訂家具備註", value=get_param("f_txt"), height=60)

    st.write("---")
    st.markdown("** 承租人點收物品**")
    c_k1, c_k2 = st.columns(2)
    chk_key_house = c_k1.checkbox("房屋鑰匙", value=True if get_param("k_h")=="1" else False)
    num_key_house = c_k2.text_input("房屋鑰匙數量", value=get_param("kn_h", "1"))
    
    c_k3, c_k4 = st.columns(2)
    chk_token = c_k3.checkbox("感應磁扣(卡)", value=True if get_param("k_t")=="1" else False)
    num_token = c_k4.text_input("感應磁扣數量", value=get_param("kn_t", "1"))
    
    c_k5, c_k6 = st.columns(2)
    chk_key_mail = c_k5.checkbox("信箱鑰匙", value=True if get_param("k_m")=="1" else False)
    num_key_mail = c_k6.text_input("信箱鑰匙數量", value=get_param("kn_m", "1"))
    
    c_k7, c_k8 = st.columns(2)
    chk_remote = c_k7.checkbox("車庫遙控器", value=True if get_param("k_r")=="1" else False)
    num_remote = c_k8.text_input("車庫遙控器數量", value=get_param("kn_r", "1"))

    st.write("---")
    st.markdown("** 出租人手寫簽名**")
    canvas_l = st_canvas(fill_color="rgba(255,255,255,0)", stroke_width=3, stroke_color="#000000", background_color="#FFFFFF", height=100, width=280, drawing_mode="freedraw", key="canvas_l", return_image_data=True)
    
    st.markdown("** 承租人手寫簽名**")
    canvas_t = st_canvas(fill_color="rgba(255,255,255,0)", stroke_width=3, stroke_color="#000000", background_color="#FFFFFF", height=100, width=280, drawing_mode="freedraw", key="canvas_t", return_image_data=True)

# ==================== 底部功能按鈕區 ====================
st.write("---")
b_col1, b_col2 = st.columns(2)

with b_col1:
    st.subheader("產生專屬簽名網址")
    if st.button(" 一鍵生成房客簽名連結", use_container_width=True):
        # 將畫面上填好的資料加密編碼進網址中
        params = {
            "l_name": landlord_name, "t_name": tenant_name, "t_id": tenant_id, "t_phone": tenant_phone,
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
            params[f"f_{k}"] = "1" if is_chk else "0"
            params[f"fn_{k}"] = num
            
        # 串接目前的網頁基礎網址
        encoded_params = urllib.parse.urlencode(params)
        share_url = f"https://streamlit.app?{encoded_params}" # 注意：上線後換成你的免費網址！
        st.info("請複製下方網址，用 Line 或簡訊傳給房客，房客打開就能直接簽名：")
        st.code(share_url, language="text")

with b_col2:
    st.subheader("雙方簽完名後生成下載")
    if st.button("線上生成合約文件", use_container_width=True):
        if not landlord_name or not tenant_name or not address:
            st.error(" 錯誤：『出租人』、『承租人姓名』與『租賃房屋地址』為必填欄位！")
        else:
            with st.spinner("系統正在處理資料，請稍候..."):
                try:
                    context = {
                        "landlord_name": landlord_name, "tenant_name": tenant_name, "tenant_id": tenant_id,
                        "tenant_phone": tenant_phone, "tenant_address": tenant_address, "address": address,
                        "rent_amount": rent_amount, "deposit_amount": deposit_amount,
                        "start_year": s_year, "start_month": s_month, "start_day": s_day,
                        "end_year": e_year, "end_month": e_month, "end_day": e_day,
                        "chk_car_yes": "■" if car_option == "有汽車位" else "□", "chk_car_no": "■" if car_option == "無汽車位" else "□",
                        "chk_moto_yes": "■" if moto_option == "有機車位" else "□", "chk_moto_no": "■" if moto_option == "無機車位" else "□",
                        "chk_mgmt_landlord": "■" if mgmt_pay == "出租人負擔" else "□", "chk_mgmt_tenant": "■" if mgmt_pay == "承租人負擔" else "□",
                        "chk_mgmt_other": "■" if mgmt_pay == "其他約定" else "□", "fee_mgmt_house": fee_mgmt_house, "fee_mgmt_car": fee_mgmt_car, "txt_mgmt_other": txt_mgmt_other,
                        "chk_water_landlord": "■" if water_pay == "出租人負擔" else "□", "chk_water_tenant": "■" if water_pay == "承租人負擔" else "□", "chk_water_other": "■" if water_pay == "其他約定" else "□", "txt_water_other": txt_water_other,
                        "chk_elec_landlord": "■" if elec_pay == "出租人負擔" else "□", "chk_elec_tenant_avg": "■" if elec_pay == "承租人負擔 (依當期平均電價)" else "□", "chk_elec_tenant_fixed": "■" if elec_pay == "承租人負擔 (固定每度元)" else "□", "chk_elec_other": "■" if elec_pay == "非以度數計費其他約定" else "□", "fee_elec_rate": fee_elec_rate, "txt_elec_other": txt_elec_other,
                        "chk_gas_landlord": "■" if gas_pay == "出租人負擔" else "□", "chk_gas_tenant": "■" if gas_pay == "承租人負擔" else "□", "chk_gas_other": "■" if gas_pay == "其他約定" else "□", "txt_gas_other": txt_gas_other,
                        "chk_net_landlord": "■" if net_pay == "出租人負擔" else "□", "chk_net_tenant": "■" if net_pay == "承租人負擔" else "□", "chk_net_other": "■" if net_pay == "其他約定" else "□", "txt_net_other": txt_net_other,
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
                        st.error("找不到 template.docx 檔案！")
                    else:
                        doc = DocxTemplate("template.docx")
                        if canvas_l.image_data is not None:
                            Image.fromarray(canvas_l.image_data.astype('uint8'), 'RGBA').save("wl.png")
                            context["landlord_sign"] = InlineImage(doc, "wl.png", width=Inches(1.2))
                        else: context["landlord_sign"] = ""
                        if canvas_t.image_data is not None:
                            Image.fromarray(canvas_t.image_data.astype('uint8'), 'RGBA').save("wt.png")
                            context["tenant_sign"] = InlineImage(doc, "wt.png", width=Inches(1.2))
                        else: context["tenant_sign"] = ""

                                           # 渲染並儲存 Word 檔
                    time_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                    out_word = f"線上合約_{tenant_name}_{time_str}.docx"
                    doc.render(context)
                    doc.save(out_word)
                    
                    # 清理簽名暫存圖
                    if os.path.exists("wl.png"): os.remove("wl.png")
                    if os.path.exists("wt.png"): os.remove("wt.png")

                    st.success("🎉 線上合約已成功產出！請點擊下方按鈕下載：")
                    
                    # 雲端版專用：提供完美包含手寫簽名與小數點的 Word 下載按鈕
                    with open(out_word, "rb") as word_file:
                        st.download_button(
                            label=" 下載最終合約 Word 檔案 (.docx)",
                            data=word_file,
                            file_name=f"住宅租賃契約書_{tenant_name}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"生成失敗，原因：{str(e)}")

