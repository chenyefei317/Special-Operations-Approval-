import datetime
import io
import os
import zipfile
import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
import requests
import streamlit as st
from streamlit_drawable_canvas import st_canvas

# 尝试安全导入 python-docx，防止云端未安装时直接崩溃
try:
  from docx import Document
  from docx.oxml.ns import qn
  from docx.shared import Inches, Pt, RGBColor

  DOCX_AVAILABLE = True
except ImportError:
  DOCX_AVAILABLE = False

# ================= 1. 页面配置与初始化 =================
st.set_page_config(
    page_title="特种作业安全许可与在线审批系统",
    page_icon="🛡️",
    layout="centered",
    initial_sidebar_state="expanded",
)

hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    html, body, [class*="css"] {
        font-family: "华文宋体", SimSun, serif;
    }
    @media print {
        .stSidebar, button, input {
            display: none !important;
        }
    }
    </style>
    """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)


# ================= 2. 6666 访问密码拦截验证 =================
if "authenticated" not in st.session_state:
  st.session_state.authenticated = False

if not st.session_state.authenticated:
  col_l, col_t = st.columns([1, 6])
  with col_l:
    try:
      st.image("logo.png", width=110)
    except Exception:
      pass
  with col_t:
    st.markdown("## 特种作业安全许可与在线审批系统")

  st.markdown("---")
  st.info("🔒 本系统为内部合规平台，请输入 **4位访问密码** 进入系统。")

  with st.form("password_form"):
    pwd_input = st.text_input(
        "请输入 4 位访问密码：", type="password", max_chars=4
    )
    submit_pwd = st.form_submit_button("进入系统", use_container_width=True)

  if submit_pwd:
    try:
      correct_pwd = st.secrets.get("APP_PASSWORD", "6666")
    except Exception:
      correct_pwd = "6666"

    if pwd_input == correct_pwd:
      st.session_state.authenticated = True
      st.rerun()
    else:
      st.error("❌ 密码错误，请重新输入（默认密码为 6666）！")

  st.markdown("---")
  st.markdown(
      "<div style='text-align: center; color: gray; font-size: 14px;'>"
      "本系统为内部合规平台，严禁商业用途 | 开发者：陈野菲 wx：13761704998"
      "</div>",
      unsafe_allow_html=True,
  )
  st.stop()


# ================= 3. GB 30871-2022 现场安全评估标准库 =================
GB30871_GENERAL_ITEMS = [
    "【通用-安全交底】作业前已对参加作业的所有人员进行安全措施交底（告知作业内容、危险有害因素、防范措施及应急处置）。",
    "【通用-持证与防护】特种作业人员持有有效的国家特种作业操作证；全员已按规定佩戴符合标准的个体防护装备（安全帽、安全鞋、防护手套等）。",
    "【通用-工器具检查】现场作业使用的工器具、移动式电动工具及安全附属设施完好无损，经检查合格且符合安全技术要求。",
    "【通用-监护人履职】安全监护人已接受专项培训并佩戴明显标识，承诺作业期间全程在岗监护，不擅自离开现场或从事与监护无关的工作。",
]

GB30871_SPECIFIC_ITEMS = {
    "动火作业（含切割）": [
        "【动火-可燃物清理】动火点周围 15m 范围内易燃易爆物品已彻底清除；下方及周围地沟、电缆沟、井口已采取严密防火封堵及遮盖措施（GB 30871 第5章）。",
        "【动火-气体检测】动火作业前已完成现场可燃及有毒有害气体检测分析，检测时间距动火开始不超过 30min，且检测数据合格。",
        "【动火-气瓶间距】气焊、气割作业时，乙炔瓶直立放置并设防倾倒措施，装有回火防止器；氧气瓶与乙炔瓶间距不小于 5m，二者与动火点间距均不小于 10m。",
        "【动火-消防器材】动火现场已配备足量且完好有效的灭火器、灭火毯、接火盆或消防水带，火花飞溅部位已设置挡火屏。",
        "【动火-电焊机规范】电焊机外壳可靠接地，焊把线绝缘良好无裸露；接地线直接接在被焊构件上，未使用厂房金属结构、管道或设备作为回路。",
    ],
    "临时用电作业": [
        "【临电-一机一闸】临时用电配电箱完好，严格执行“一机一闸一保护”，已安装合格的漏电保护器（RCD），经测试动作灵敏可靠（GB 30871 第10章）。",
        "【临电-接地与防爆】所有临时用电设备及金属外壳均已实施可靠保护接地或接零；存在爆炸或火灾危险场所的电气元件符合相应防爆等级要求。",
        "【临电-线路敷设】采用耐压等级不低于 500V 的绝缘铜芯导线；架空敷设时最大弧垂距地面不低于 2.5m（跨越道路不低于 5m）；地面走线已加装防砸、防碾压护套管。",
        "【临电-警示与防护】配电箱、开关箱设有明显的安全警示标识，防雨、防尘、防潮措施到位；检修或非作业期间严格执行上锁挂牌（LOTO）。",
    ],
    "登高作业（高处作业）": [
        "【登高-防坠落保护】作业人员已正确佩戴符合 GB 6095 标准的全身式双大钩安全带，严格执行“高挂低用”，系挂点牢固可靠（GB 30871 第8章）。",
        "【登高-设施验收】作业使用的脚手架、移动式升降平台、人字梯或高空作业车已检查验收合格，支腿稳固、护栏齐全、防滑及制动有效。",
        "【登高-防高空落物】作业人员随身配备工具袋，工器具已采取防坠绳系挂；高处作业正下方已拉设安全警戒线与警示牌，严禁无关人员进入。",
    ],
    "受限空间作业": [
        "【受限-隔绝与置换】受限空间与其他系统已采取加装盲板或拆除一段管道的方式可靠隔绝；内部已完成清洗、吹扫、置换（GB 30871 第6章）。",
        "【受限-气体分析】作业前 30min 内已完成氧含量、可燃气体及有毒有害气体检测且合格，作业中至少每 2h 复测或连续监测。",
        "【受限-通风照明】已采取自然通风或强制机械通风措施（严禁充纯氧）；受限空间内使用安全电压不大于 36V（潮湿狭小场所不大于 12V）的防爆照明。",
    ],
    "吊装作业": [
        "【吊装-机具资质】起重机械、吊索具、卡环经检查完好且在检验有效期内，司索工、指挥人员及起重工均持有效特种作业证（GB 30871 第9章）。",
        "【吊装-警戒试吊】吊装区域已设置安全警戒区；正式起吊前已进行试吊，确认制动、吊点及平衡状态正常，严格遵守“十不吊”规定。",
    ],
    "盲板抽堵作业": [
        "【盲板-系统泄压】设备管道已完成卸压、排空与置换，已绘制盲板位置图并统一编号，盲板材质、厚度及垫片符合介质压力要求（GB 30871 第7章）。",
    ],
    "动土作业": [
        "【动土-地下管线】已核查地下电缆、管线分布并现场交底；挖掘深度超过 1.2m 且存在危险物料积聚时，应执行受限空间作业相关规定（GB 30871 第11章）。",
    ],
    "断路作业": [
        "【断路-交通警示】已制定交通组织方案，在断路路口和相关道路设置交通警示标志、路栏及道路作业警示灯（GB 30871 第12章）。",
    ],
}


# ================= 4. 百度网盘自动上传函数 =================
def refresh_baidu_access_token():
  try:
    client_id = st.secrets.get("BAIDU_CLIENT_ID", "")
    client_secret = st.secrets.get("BAIDU_CLIENT_SECRET", "")
    refresh_token = st.secrets.get("BAIDU_REFRESH_TOKEN", "")
    if not client_id or not client_secret or not refresh_token:
      return None
    token_url = "https://pan.baidu.com/oauth/2.0/token"
    params = {
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "client_id": client_id,
        "client_secret": client_secret,
    }
    response = requests.get(token_url, params=params)
    res_data = response.json()
    return res_data.get("access_token")
  except Exception:
    return None


def upload_to_baidu_netdisk_with_auto_refresh(file_bytes, remote_filename):
  try:
    access_token = st.secrets.get("BAIDU_ACCESS_TOKEN", "")
    if not access_token:
      return False, "未配置网盘凭证，文件已在本地生成并可通过网页下载。"
  except Exception:
    return (
        False,
        "本地未检测到 secrets 配置，已跳过网盘同步，网页下载功能正常使用。",
    )

  sub_folder = "特种作业审批档案"
  target_path = f"/apps/慧瑞EHS合规档案/{sub_folder}/{remote_filename}"

  def send_upload_request(token):
    upload_url = f"https://pan.baidu.com/rest/2.0/xpan/file?method=upload&access_token={token}&path={target_path}&uploadid=&file=1"
    files = {"file": (remote_filename, file_bytes)}
    return requests.post(upload_url, files=files).json()

  try:
    result = send_upload_request(access_token)
    if "errno" in result and result["errno"] in [110, 111]:
      new_token = refresh_baidu_access_token()
      if new_token:
        result = send_upload_request(new_token)
      else:
        return False, "Token 已过期且自动刷新失败。"

    if "errno" in result and result["errno"] == 0:
      return True, f"成功同步至网盘：/apps/慧瑞EHS合规档案/{sub_folder}/"
    else:
      return False, f"网盘上传失败: {result.get('error_msg', '未知错误')}"
  except Exception as e:
    return False, f"网盘上传异常: {str(e)}"


# ================= 5. 侧边栏：Logo与带有“EHS野哥”文字提示的微信二维码 =================
with st.sidebar:
  try:
    st.image("logo.png", width=160)
  except Exception:
    pass

  st.markdown("### 📱 微信扫码与分享")
  st.write(
      "手机微信扫码下方二维码，随时随地在线填报特种作业审批并生成许可证。"
  )

  app_url = st.text_input(
      "应用公网链接 (URL)", value="https://special-work-permit.streamlit.app/"
  )

  if app_url:
    qr = qrcode.QRCode(version=1, box_size=10, border=2)
    qr.add_data(app_url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert(
        "RGB"
    )

    draw = ImageDraw.Draw(qr_img)
    w, h = qr_img.size
    box_w, box_h = 120, 36
    box_x = (w - box_w) / 2
    box_y = (h - box_h) / 2
    draw.rectangle(
        [box_x, box_y, box_x + box_w, box_y + box_h],
        fill="white",
        outline="black",
        width=2,
    )
    try:
      font = ImageFont.truetype("simsun.ttc", 14)
    except Exception:
      font = ImageFont.load_default()
    draw.text((box_x + 18, box_y + 10), "EHS野哥", fill="black", font=font)

    img_buffer = io.BytesIO()
    qr_img.save(img_buffer, format="PNG")
    st.image(
        Image.open(img_buffer),
        caption="【手机扫码：EHS野哥特种作业审批】",
        width=160,
    )

  st.markdown("---")
  st.markdown("### 🔗 快捷操作")
  st.info(
      "💡 **复制链接提示**：请直接复制上方输入框内的网址，或点击浏览器地址栏复制发送至微信群。"
  )

# ================= 6. 主界面头部 =================
col_logo, col_title = st.columns([1, 6])
with col_logo:
  try:
    st.image("logo.png", width=110)
  except Exception:
    pass
with col_title:
  st.markdown("## 特种作业安全许可与在线审批系统")

st.markdown(
    "本系统严格依据国家标准 **GB 30871-2022《危险化学品企业特殊作业安全规范》** 建设。手写签字为选填项，生成 Word 文档后可直接下载打印进行线下纸质签字。"
)

# ================= 7. 模块一：作业性质、类型与基础信息 =================
st.write("---")
st.subheader("1. 作业性质、类型与基础信息")

work_nature = st.radio(
    "作业性质 (必填)：",
    options=["1）承包商", "2）正式工"],
    horizontal=True,
    help="请选择本次作业单位的用工性质",
)

work_type_options = list(GB30871_SPECIFIC_ITEMS.keys())

col_w1, col_w2 = st.columns(2)
with col_w1:
  selected_work_type = st.selectbox(
      "特种作业类型 (依据国家标准下拉选择)：",
      options=work_type_options,
      index=0,
      help="切换作业类型后，下方自动匹配对应的 GB 30871 安全评估要求",
  )
with col_w2:
  work_location = st.text_input(
      "具体作业地点 / 区域 (必填)：",
      placeholder="例如：一车间北侧维修区 / 配电房",
  )

col_u1, col_u2 = st.columns(2)
with col_u1:
  work_unit = st.text_input(
      "作业单位 / 部门名称 (必填)：",
      placeholder="例如：XX建设工程有限公司 / 设备工程部",
  )
with col_u2:
  worker_names = st.text_input(
      "作业人员姓名及特种作业证号 (必填)：",
      placeholder="例如：张三 (T350123...)、李四",
  )

work_content = st.text_input(
    "作业内容及主要工器具描述 (必填)：",
    placeholder="例如：管道支架角磨机切割与电焊焊接 / 临时接入 380V 检修电源箱",
)

# ================= 8. 模块二：签发日期与截止日期（以 72 小时为限） =================
st.write("---")
st.subheader("2. 作业许可有效期限（最长不超过 72 小时）")

now_dt = datetime.datetime.now().replace(second=0, microsecond=0)
default_end_dt = now_dt + datetime.timedelta(hours=24)

col_d1, col_t1, col_d2, col_t2 = st.columns(4)
with col_d1:
  issue_date = st.date_input("签发日期：", value=now_dt.date(), key="issue_d")
with col_t1:
  issue_time = st.time_input("签发时间：", value=now_dt.time(), key="issue_t")
with col_d2:
  expire_date = st.date_input(
      "截止日期：", value=default_end_dt.date(), key="expire_d"
  )
with col_t2:
  expire_time = st.time_input(
      "截止时间：", value=default_end_dt.time(), key="expire_t"
  )

issue_datetime = datetime.datetime.combine(issue_date, issue_time)
expire_datetime = datetime.datetime.combine(expire_date, expire_time)
duration_seconds = (expire_datetime - issue_datetime).total_seconds()
duration_hours = duration_seconds / 3600.0

is_time_valid = False
if duration_hours <= 0:
  st.error("❌ 时间设置错误：【截止日期/时间】必须晚于【签发日期/时间】！")
elif duration_hours > 72.0:
  st.error(
      f"❌ 违反合规限值：当前设置的作业有效时长为 **{duration_hours:.1f} 小时**，已超过最长 **72 小时上限**！请缩短截止时间。"
  )
else:
  is_time_valid = True
  st.success(
      f"✅ 许可有效期核验通过：自 **{issue_datetime.strftime('%Y-%m-%d %H:%M')}** 至 **{expire_datetime.strftime('%Y-%m-%d %H:%M')}**（共计 **{duration_hours:.1f} 小时**，符合 ≤72 小时合规限值）。"
  )

# ================= 9. 模块三：现场安全评估（GB 30871-2022 标准打钩） =================
st.write("---")
st.subheader(
    f"3. 现场安全评估核查表（GB 30871-2022 - {selected_work_type}）"
)

current_assessment_items = (
    GB30871_SPECIFIC_ITEMS.get(selected_work_type, []) + GB30871_GENERAL_ITEMS
)


def toggle_all_gb_checks():
  target_val = st.session_state.get("select_all_gb", False)
  for idx in range(len(current_assessment_items)):
    st.session_state[f"gb_chk_{selected_work_type}_{idx}"] = target_val


st.checkbox(
    "☑ 选择全部评估项（确认现场已100%落实以下全部安全措施）",
    key="select_all_gb",
    on_change=toggle_all_gb_checks,
)

st.info(
    f"请对照 **GB 30871-2022** 针对【{selected_work_type}】的法定要求，在作业现场逐项核实并打钩确认："
)

assessment_results = []
for i, item_text in enumerate(current_assessment_items):
  chk_key = f"gb_chk_{selected_work_type}_{i}"
  if chk_key not in st.session_state:
    st.session_state[chk_key] = False
  checked = st.checkbox(item_text, key=chk_key)
  assessment_results.append(checked)

# ================= 10. 模块四：三方责任人签字与手写日期（选填项） =================
st.write("---")
st.subheader("✍️ 4. 责任人信息与手写签名（选填：留空则打印后线下手动签字）")

col_p1, col_p2, col_p3 = st.columns(3)
with col_p1:
  supervisor_name = st.text_input(
      "现场作业负责人姓名：", placeholder="选填（如：张三）"
  )
with col_p2:
  guardian_name = st.text_input(
      "安全监护人姓名：", placeholder="选填（如：李四）"
  )
with col_p3:
  approver_name = st.text_input("审批人姓名：", placeholder="选填（如：王五）")

st.markdown(
    "💡 **提示**：若不在此处画板手写签名，生成的 Word 文件中将保留空白签名栏，方便打印后纸质手写签字。"
)

col_c1, col_c2, col_c3 = st.columns(3)
with col_c1:
  st.markdown("📝 **1. 现场作业负责人签字（可选）：**")
  canvas_supervisor = st_canvas(
      stroke_width=3,
      stroke_color="#000000",
      background_color="#F8F9FA",
      height=160,
      width=220,
      drawing_mode="freedraw",
      key="canvas_supervisor",
      return_image_data=True,
  )
with col_c2:
  st.markdown("🛡️ **2. 安全监护人签字（可选）：**")
  canvas_guardian = st_canvas(
      stroke_width=3,
      stroke_color="#000000",
      background_color="#F8F9FA",
      height=160,
      width=220,
      drawing_mode="freedraw",
      key="canvas_guardian",
      return_image_data=True,
  )
with col_c3:
  st.markdown("✅ **3. 审批人签字（可选）：**")
  canvas_approver = st_canvas(
      stroke_width=3,
      stroke_color="#000000",
      background_color="#F8F9FA",
      height=160,
      width=220,
      drawing_mode="freedraw",
      key="canvas_approver",
      return_image_data=True,
  )

current_date_str = datetime.date.today().strftime("%Y年%m月%d日")
st.markdown(
    f"**📅 手写签发日期栏（可选，当前系统日期为 {current_date_str}）：**"
)
canvas_date = st_canvas(
    stroke_width=3,
    stroke_color="#000000",
    background_color="#F8F9FA",
    height=130,
    width=340,
    drawing_mode="freedraw",
    key="canvas_date_permit",
    return_image_data=True,
)


def is_canvas_blank(canvas_obj):
  return canvas_obj.image_data is None or (
      canvas_obj.json_data is not None
      and len(canvas_obj.json_data.get("objects", [])) == 0
  )


# ================= 11. 生成正式 Word《特种作业安全许可证》函数 =================
def generate_work_permit_docx(
    permit_no,
    w_nature,
    w_type,
    w_loc,
    w_unit,
    w_workers,
    w_content,
    dt_issue,
    dt_expire,
    dur_hrs,
    eval_items,
    eval_checks,
    sup_name,
    grd_name,
    app_name,
    sup_img_io,
    grd_img_io,
    app_img_io,
    date_img_io,
):
  if not DOCX_AVAILABLE:
    return None

  doc = Document()

  p_title = doc.add_paragraph()
  r_title = p_title.add_run(f"【特种作业安全许可证 / 审批单 — {w_type}】")
  r_title.bold = True
  r_title.font.size = Pt(15)
  r_title.font.name = "华文宋体"
  r_title.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

  p_meta = doc.add_paragraph()
  r_meta = p_meta.add_run(
      f"许可证编号：{permit_no}        执行标准：GB 30871-2022\n"
      f"作业性质：{w_nature}        作业类型：{w_type}\n"
      f"作业地点：{w_loc}        作业单位：{w_unit}\n"
      f"作业人员及证号：{w_workers}\n"
      f"作业内容及工器具：{w_content}\n"
      f"许可签发时间：{dt_issue.strftime('%Y-%m-%d %H:%M')}\n"
      f"许可截止时间：{dt_expire.strftime('%Y-%m-%d %H:%M')} （有效时长：{dur_hrs:.1f} 小时，符合 ≤72 小时合规限值）\n"
      "-------------------------------------------------------------------------"
  )
  r_meta.font.size = Pt(10.5)
  r_meta.font.name = "华文宋体"
  r_meta.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

  p_eval_head = doc.add_paragraph()
  r_eh = p_eval_head.add_run("一、GB 30871-2022 现场安全评估核查确认表：")
  r_eh.bold = True
  r_eh.font.size = Pt(11)
  r_eh.font.name = "华文宋体"
  r_eh.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

  eval_lines = []
  for idx, text in enumerate(eval_items):
    status_str = "[√ 已核实符合]" if eval_checks[idx] else "[× 未勾选]"
    eval_lines.append(f"{idx + 1}. {text}  —— {status_str}")

  p_eval_body = doc.add_paragraph()
  r_eb = p_eval_body.add_run("\n".join(eval_lines))
  r_eb.font.size = Pt(9.5)
  r_eb.font.name = "华文宋体"
  r_eb.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

  p_sig_head = doc.add_paragraph()
  r_sh = p_sig_head.add_run(
      "-------------------------------------------------------------------------\n"
      "二、三方安全责任确认与签字（可在线手写或打印后线下手签）："
  )
  r_sh.bold = True
  r_sh.font.size = Pt(11)
  r_sh.font.name = "华文宋体"
  r_sh.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

  table = doc.add_table(rows=2, cols=2)
  table.autofit = False

  c00 = table.cell(0, 0).paragraphs[0]
  r00 = c00.add_run(
      f"1. 现场作业负责人：{sup_name if sup_name else '（签字：       ）'}\n"
  )
  r00.font.name = "华文宋体"
  r00.font.size = Pt(10)
  r00.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")
  if sup_img_io:
    c00.add_run().add_picture(sup_img_io, width=Inches(1.6))
    sup_img_io.seek(0)

  c01 = table.cell(0, 1).paragraphs[0]
  r01 = c01.add_run(
      f"2. 安全监护人：{grd_name if grd_name else '（签字：       ）'}\n"
  )
  r01.font.name = "华文宋体"
  r01.font.size = Pt(10)
  r01.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")
  if grd_img_io:
    c01.add_run().add_picture(grd_img_io, width=Inches(1.6))
    grd_img_io.seek(0)

  c10 = table.cell(1, 0).paragraphs[0]
  r10 = c10.add_run(
      f"3. 审批人：{app_name if app_name else '（签字：       ）'}\n"
  )
  r10.font.name = "华文宋体"
  r10.font.size = Pt(10)
  r10.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")
  if app_img_io:
    c10.add_run().add_picture(app_img_io, width=Inches(1.6))
    app_img_io.seek(0)

  c11 = table.cell(1, 1).paragraphs[0]
  r11 = c11.add_run("4. 签发日期：\n")
  r11.font.name = "华文宋体"
  r11.font.size = Pt(10)
  r11.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")
  if date_img_io:
    c11.add_run().add_picture(date_img_io, width=Inches(1.6))
    date_img_io.seek(0)

  buffer = io.BytesIO()
  doc.save(buffer)
  buffer.seek(0)
  return buffer


# ================= 12. 提交校验与生成合规档案 =================
st.write("---")
if st.button(
    "📁 确认安全措施已落实，一键生成并下载《特种作业安全许可证》",
    use_container_width=True,
    type="primary",
):
  if not DOCX_AVAILABLE:
    st.error(
        "❌ 服务器未检测到 `python-docx` 库。请在 GitHub 仓库的 `requirements.txt` 文件中确保已正确写入 `python-docx` 并在后台重启应用！"
    )
  elif (
      not work_location.strip()
      or not work_unit.strip()
      or not worker_names.strip()
      or not work_content.strip()
  ):
    st.error("❌ 拦截：请完整填写【作业地点】、【作业单位】、【作业人员】及【作业内容】！")
  elif not is_time_valid:
    st.error(
        "❌ 拦截：作业许可有效时间设置不合规（必须大于 0 且不超过 72 小时），请重新核定签发与截止时间！"
    )
  elif not all(assessment_results):
    st.warning(
        "⚠️ 拦截：按照 GB 30871-2022 强制性规范要求，现场所有安全评估项必须全部核实打钩后方可签发作业许可！"
    )
  else:
    permit_number = f"PTW-{datetime.datetime.now().strftime('%Y%m%d-%H%M%S')}"
    st.success(
        f"✅ 特种作业安全许可证审批表生成成功！许可证编号：**{permit_number}**"
    )

    def canvas_to_bytes_io_optional(canvas_obj):
      if is_canvas_blank(canvas_obj):
        return None
      img = Image.fromarray(canvas_obj.image_data.astype("uint8"), "RGBA")
      b_io = io.BytesIO()
      img.save(b_io, format="PNG")
      b_io.seek(0)
      return b_io

    sup_io = canvas_to_bytes_io_optional(canvas_supervisor)
    grd_io = canvas_to_bytes_io_optional(canvas_guardian)
    app_io = canvas_to_bytes_io_optional(canvas_approver)
    dt_io = canvas_to_bytes_io_optional(canvas_date)

    permit_docx_buffer = generate_work_permit_docx(
        permit_number,
        work_nature,
        selected_work_type,
        work_location.strip(),
        work_unit.strip(),
        worker_names.strip(),
        work_content.strip(),
        issue_datetime,
        expire_datetime,
        duration_hours,
        current_assessment_items,
        assessment_results,
        supervisor_name.strip(),
        guardian_name.strip(),
        approver_name.strip(),
        sup_io,
        grd_io,
        app_io,
        dt_io,
    )

    if permit_docx_buffer:
      docx_filename = f"特种作业审批单_{selected_work_type}_{work_location.strip()}_{permit_number}.docx"

      zip_buffer = io.BytesIO()
      with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        zip_file.writestr(docx_filename, permit_docx_buffer.getvalue())
        if sup_io:
          zip_file.writestr(
              f"作业负责人签字_{supervisor_name.strip()}.png",
              sup_io.getvalue(),
          )
        if grd_io:
          zip_file.writestr(
              f"监护人签字_{guardian_name.strip()}.png", grd_io.getvalue()
          )
        if app_io:
          zip_file.writestr(
              f"审批人签字_{approver_name.strip()}.png", app_io.getvalue()
          )
        if dt_io:
          zip_file.writestr("手写签发日期.png", dt_io.getvalue())

        upload_to_baidu_netdisk_with_auto_refresh(
            permit_docx_buffer.getvalue(), docx_filename
        )

      zip_buffer.seek(0)

      st.markdown("---")
      st.success(
          "🎉 特种作业审批档案已生成完毕！推荐直接下载 Word 后进行打印与线下签字："
      )

      col_dl1, col_dl2 = st.columns(2)
      with col_dl1:
        st.download_button(
            label="📄 下载《特种作业安全许可证》(.docx) [最适合直接打印]",
            data=permit_docx_buffer.getvalue(),
            file_name=docx_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
      with col_dl2:
        st.download_button(
            label="📥 一键打包下载完整审批档案 (.ZIP)",
            data=zip_buffer,
            file_name=f"特种作业审批档案_{permit_number}.zip",
            mime="application/zip",
            use_container_width=True,
        )

      st.snow()

# ================= 13. 底部版权与开发者声明 =================
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray; font-size: 14px;'>"
    "本系统为内部合规平台，严禁商业用途 | 开发者：陈野菲 wx：13761704998"
    "</div>",
    unsafe_allow_html=True,
)
