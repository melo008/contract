import os
import urllib.parse
import io
import base64
import streamlit as st
from streamlit_drawable_canvas import st_canvas
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from PIL import Image

# 網頁基本設定
st.set_page_config(page_title="內政部租賃合約線上簽署系統", layout="wide")
st.title("🏠 住宅租賃契約書 - 線上合約簽署系統")
st.write("【房東專區】：填完合約細節並手寫簽名後，點擊底部即可生成『極精簡短網址』傳給房客，房客免註冊登入即可補簽。")

def get_short_url(long_url):
    """呼叫網路通用的免費 TinyURL API"""
    try:
        import requests
        api_url = f"http://tinyurl.com{urllib.parse.quote(long_url)}"
        res = requests.get(api_url, timeout=5)
        if res.status_code == 200 and res.text:
            return res.text
    except:
        pass
    return long_url

# 【核心修正】：使用 Streamlit 官方新版規範字典解碼，確保房客點開時資料 100% 帶入
def get_param(key, default=""):
    try:
        if key in st.query_params:
            return st.query_params[key]
    except:
        pass
    return default

# ==================== 建立網頁左、中、右三欄版面 ====================
col1, col2, col3 = st.columns(3)

# 1. 基本與租期資料
with col1:
    st.header("1. 基本與租期資料")
    landlord_name = st.text_input("出租人姓名 *", value=get_param("l_name"))
    tenant_name = st.text_input("承租人姓名 *", value=get_param("t_name"))
    tenant_id = st.text_input("承租人身分證字號")
    tenant_phone = st.text_input("承租人電話")
    tenant_address = st.text_input("承租人戶籍地址")
    address = st.text_input("租賃房屋地址 *", value=get_param("addr"))
    rent_amount = st.text_input("每月租金 (元)", value=get_param("rent"))
    deposit_amount = st.text_input("押金金額 (元)")
    
    st.subheader("租期時間")
    c_s1, c_s2, c_s3 = st.columns(3)
    s_year = c_s1.text_input("開始年(民國)", value="115")
    s_month = c_s2.text_input("開始月", value="1")
    s_day = c_s3.text_input("開始日", value="1")
    
    c_e1, c_e2, c_e3 = st.columns(3)
    e_year = c_e1.text_input("結束年(民國)", value="116")
    e_month = c_e2.text_input("結束月", value="1")
    e_day = c_e3.text_input("結束日", value="1")
    
    car_option = st.radio("汽車位需求", ["無汽車位", "有汽車位"], index=0)
    moto_option = st.radio("機車位需求", ["無機車位", "有機車位"], index=0)

# 2. 租賃期間費用約定
with col2:
    st.header("2. 租賃期間費用約定")
    mgmt_pay = st.radio("管理費負擔方", ["出租人負擔", "承租人負擔", "其他約定"], index=0, key="mgmt")
    fee_mgmt_house = st.text_input("住宅管理費/月 (元)", value="0")
    fee_mgmt_car = st.text_input("車位管理費/月 (元)", value="0")
    txt_mgmt_other = st.text_input("管理費其他約定說明")
    
    water_pay = st.radio("水費負擔方", ["出租人負擔", "承租人負擔", "其他約定"], index=1, key="water")
    txt_water_other = st.text_input("水費其他約定說明")
    
    elec_pay = st.radio("電費計費方式", ["出租人負擔", "承租人負擔 (依當期平均電價)", "承租人負擔 (固定每度元)", "非度數計費其他約定"], index=1, key="elec")
    fee_elec_rate = st.text_input("固定每度電費 (元)", value="0")
    txt_elec_other = st.text_input("電費other約定說明")
    
    gas_pay = st.radio("瓦斯費負擔方", ["出租人負擔", "承租人負擔", "其他約定"], index=1, key="gas")
    txt_gas_other = st.text_input("瓦斯費其他約定說明")
    
    net_pay = st.radio("網路費負擔方", ["出租人負擔", "承租人負擔", "其他約定"], index=0, key="net")
    txt_net_other = st.text_input("網路費其他約定說明")
    txt_other_fee = st.text_input("請輸入其他費用說明")

# 3. 設備清單、點收物品與手寫簽名
with col3:
    st.header("3. 附屬設備、點收與簽名")
    furniture_items = [
        ("bed_frame", "床架"), ("mattress", "床墊"), ("wardrobe", "衣櫃"), ("table", "桌子"), ("chair", "椅子"),
        ("sofa", "沙發"), ("tea_table", "茶几"), ("ac", "冷氣機"), ("fridge", "冰箱"), ("washer", "洗衣機"),
        ("tv", "電視機"), ("heater", "熱水器"), ("cooker", "電磁爐"), ("lighting", "燈具"), ("toilet", "馬桶"),
        ("sink", "洗手台"), ("shower", "蓮蓬頭")
    ]
    
    st.markdown("**🏢 附屬設備清單**")
    fur_context = {}
    with st.expander("點擊展開常見家具清單"):
        for key, name in furniture_items:
            c_f1, c_f2 = st.columns(2)
            is_checked = c_f1.checkbox(name, value=False, key=f"cb_{key}")
            num = c_f2.text_input("數量", value="1", key=f"num_{key}")
            fur_context[key] = (is_checked, num)
            
    chk_other_f = st.checkbox("其他自訂設備")
    textarea_other = st.text_area("請輸入自訂家具備註", height=60)

    st.write("---")
    st.markdown("**🔑 承租人點收物品**")
    c_k1, c_k2 = st.columns(2)
    chk_key_house = c_k1.checkbox("房屋鑰匙")
    num_key_house = c_k2.text_input("房屋鑰匙數量", value="1")
    chk_token = c_k1.checkbox("感應磁扣(卡)")
    num_token = c_k2.text_input("感應磁扣數量", value="1")
    chk_key_mail = c_k1.checkbox("信箱鑰匙")
    num_key_mail = c_k2.text_input("信箱鑰匙數量", value="1")
    chk_remote = c_k1.checkbox("車庫遙控器")
    num_remote = c_k2.text_input("車庫遙控器數量", value="1")

    st.write("---")
    canvas_l = st_canvas(fill_color="rgba(255,255,255,0)", stroke_width=3, stroke_color="#000000", background_color="#FFFFFF", height=100, width=280, drawing_mode="freedraw", key="canvas_l", return_image_data=True)
    st.markdown("**✒️ 出租人(房東)手寫簽名**")
    
    canvas_t = st_canvas(fill_color="rgba(255,255,255,0)", stroke_width=3, stroke_color="#000000", background_color="#FFFFFF", height=100, width=280, drawing_mode="freedraw", key="canvas_t", return_image_data=True)
    st.markdown("**✒️ 承租人(房客)手寫簽名**")
# ==================== 底部功能按鈕區 ====================
st.write("---")
b_col1, b_col2 = st.columns(2)

with b_col1:
    st.subheader("【房東步驟 1】：產生專屬短網址")
    if st.button("🔗 一鍵生成房客簽名連結", use_container_width=True):
        if canvas_l.image_data is not None and canvas_l.image_data.any():
            Image.fromarray(canvas_l.image_data.astype('uint8'), 'RGBA').save("landlord_last_sign.png")
            
        params = {
            "l_name": landlord_name, 
            "t_name": tenant_name, 
            "addr": address, 
            "rent": rent_amount
        }
        encoded_params = urllib.parse.urlencode(params)
        raw_long_url = f"https://streamlit.app?{encoded_params}"
        
        with st.spinner("正在為您進行網址精簡縮短..."):
            short_url = get_short_url(raw_long_url)
            
        st.success("🎉 萬用短網址生成成功！傳給房客點開即可（完全免註冊、免登入）：")
        st.code(short_url, language="text")

with b_col2:
    st.subheader("【房客與房東步驟 2】：雙方簽完名後生成下載")
    if st.button("🚀 線上生成合約文件", use_container_width=True):
        if not landlord_name or not tenant_name or not address:
            st.error("❌ 錯誤：『出租人』、『承租人姓名』與『租賃房屋地址』為必填欄位！")
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
                        "chk_net_landlord": "■" if net_pay == "出租人負擔" else "□", "chk_net_tenant": "■" if net_pay == "承租人負擔" else "□", "chk_net_other": "■" if net_pay == "其他約定" else "□", "txt_net_other": net_pay if 'net_pay' in locals() else txt_net_other,
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
                        
                        if canvas_l.image_data is not None and canvas_l.image_data.any():
                            Image.fromarray(canvas_l.image_data.astype('uint8'), 'RGBA').save("wl.png")
                            context["landlord_sign"] = InlineImage(doc, "wl.png", width=Inches(1.2))
                        elif os.path.exists("landlord_last_sign.png"):
                            context["landlord_sign"] = InlineImage(doc, "landlord_last_sign.png", width=Inches(1.2))
                        else:
                            context["landlord_sign"] = ""
                            
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

                        st.success("🎉 線上合約已成功產出！請點擊下載您的合約檔案：")
                        with open(out_word, "rb") as wf:
                            st.download_button(label="📥 下載最終雙方簽署合約 Word 檔案 (.docx)", data=wf, file_name=f"住宅租賃契約書_{tenant_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
                except Exception as e:
                    st.error(f"生成失敗，原因：{str(e)}")
