import datetime
import io
import os
import zipfile
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
import pandas as pd
import qrcode
from PIL import Image
import requests
import streamlit as st
from streamlit_drawable_canvas import st_canvas

# ================= 1. 页面配置与初始化 =================
st.set_page_config(
    page_title="特种作业安全许可与在线审批系统",
    page_icon="🛡️",
    layout="centered",
    initial_sidebar_state="expanded",
)

# 注入华文宋体全局样式与隐藏默认元素
hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    html, body, [class*="css"] {
        font-family: "华文宋体", SimSun, serif;
    }
    </style>
    """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ================= 2. GB 30871-2022 现场安全评估标准库 =================
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
        "【临电-线路敷设】采用耐压等级不低于 500V 的绝缘铜芯导线；架空敷设时最大弧垂距地面不低于 2.5m（跨越道路不低于 5m）；地面走线已加装防砸、防碾压护套管，未在树木或脚手架上直接架设。",
        "【临电-警示与防护】配电箱、开关箱设有明显的安全警示标识，防雨、防尘、防潮措施到位；检修或非作业期间严格执行上锁挂牌（LOTO）。",
    ],
    "登高作业（高处作业）": [
        "【登高-防坠落保护】作业人员已正确佩戴符合 GB 6095 标准的全身式双大钩安全带，严格执行“高挂低用”，系挂点牢固可靠，必要时设置水平生命线或防坠器（GB 30871 第8章）。",
        "【登高-登高设施验收】作业使用的脚手架、移动式升降平台、人字梯或高空作业车已检查验收合格，支腿稳固、护栏齐全、防滑及制动有效。",
        "【登高-防高空落物】作业人员随身配备工具袋，工器具已采取防坠绳系挂；高处作业正下方已拉设安全警戒线与警示牌，严禁无关人员进入或违规立体交叉作业。",
        "【登高-身体与气象】已确认作业人员无高血压、心脏病、恐高症等职业禁忌证；现场无五级以上大风、暴雨、雷电、浓雾等恶劣天气。",
    ],
    "受限空间作业": [
        "【受限-隔绝与置换】受限空间与其他系统已采取加装盲板或拆除一段管道的方式可靠隔绝；内部已完成清洗、吹扫、置换（GB 30871 第6章）。",
        "【受限-气体分析】作业前 30min 内已完成氧含量（19.5%~21%，富氧不大于23.5%）、可燃气体及有毒有害气体检测且合格，作业中至少每 2h 复测或连续监测。",
        "【受限-通风与照明】已采取自然通风或强制机械通风措施（严禁充纯氧）；受限空间内使用安全电压不大于 36V（潮湿狭小场所不大于 12V）的防爆照明。",
        "【受限-应急救援】出入口保持畅通，现场已备妥正压式空气呼吸器、全身式安全带、救生索及三脚架等应急救援装备。",
    ],
    "吊装作业": [
        "【吊装-机具与资质】起重机械、吊索具、卡环经检查完好且在检验有效期内，司索工、指挥人员及起重工均持有效特种作业证（GB 30871 第9章）。",
        "【吊装-警戒与试吊】吊装区域已设置安全警戒区，非作业人员严禁入内；正式起吊前已进行试吊，确认制动、吊点及平衡状态正常，严格遵守“十不吊”规定。",
    ],
    "盲板抽堵作业": [
        "【盲板-泄压与位置图】设备管道已完成卸压、排空与置换，已绘制盲板位置图并统一编号，盲板材质、厚度及垫片符合介质压力要求（GB 30871 第7章）。",
    ],
    "动土作业": [
        "【动土-地下管线核查】已核查地下电缆、给排水、消防及燃气管道分布并现场交底确认；挖掘深度超过 1.2m 时已按规定放坡或加装护壁支撑（GB 30871 第11章）。",
    ],
    "断路作业": [
        "【断路-警示与应急通道】断路作业点周围已设置交通警示牌、围栏及夜间警示红灯，已制定应急消防车辆绕行方案并通知相关部门（GB 30871 第12章）。",
    ],
}


# ================= 3. 百度网盘自动上传函数（带 OAuth2 自动刷新） =================
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
  access_token = st.secrets.get("BAIDU_ACCESS_TOKEN", "")
  if not access_token:
    return (
        False,
        "未配置网盘凭证，文件已在本地生成并可通过网页下载/ZIP打包保存。",
    )

  sub_folder = "特种作业审批档案"
  target_path = f"/apps/慧瑞EHS合规档案/{sub_folder}/{remote_filename}"

  def send_upload_request(token):
    upload_url = f"https://pan.baidu.com/rest/2.0/xpan/file?method=upload&access_token={token}&path={target_path}&uploadid=&file=1"
    files = {"file": (remote_filename, file_bytes)}
    return requests.post(upload_url, files=files).json()

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


# ================= 4. 侧边栏：Logo与微信扫码分享 =================
with st.sidebar:
  try:
    st.image("logo.png", width=160)
  except Exception:
    st.image(
        "https://upload.wikimedia.org/wikipedia/commons/thumb/c/c5/Ikea_logo.svg/800px-Ikea_logo.svg.png",
        width=160,
    )

  st.markdown("### 📱 微信扫码与分享")
  st.write("现场作业人员与审批人可直接使用手机微信扫码，在线填报并完成手写签字。")

  app_url = st.text_input(
      "应用公网链接 (URL)", value="https://special-work-permit.streamlit.app/"
  )

  if app_url:
    qr = qrcode.make(app_url)
    img_buffer = io.BytesIO()
    qr.save(img_buffer, format="PNG")
    st.image(
        Image.open(img_buffer), caption="微信扫码快速进入特种作业审批", width=160
    )
    st.info(
        "💡 **合规提示**：特种作业许可证有效期限最长不得超过 **8 小时**，超时须重新办理审批。"
    )

# ================= 5. 主界面头部 =================
col_logo, col_title = st.columns([1, 6])
with col_logo:
  try:
    st.image("logo.png", width=110)
  except Exception:
    st.image(
        "https://upload.wikimedia.org/wikipedia/commons/thumb/c/c5/Ikea_logo.svg/800px-Ikea_logo.svg.png",
        width=110,
    )
with col_title:
  st.markdown("## 特种作业安全许可与在线审批系统")

st.markdown(
    "本系统严格依据国家标准 **GB 30871-2022《危险化学品企业特殊作业安全规范》** 建设。请现场作业负责人、监护人及审批人逐项核实现场安全条件，核定作业有效时间（**以 8 小时为限**），并在底部完成三方亲笔手写签字。"
)

# ================= 6. 模块一：作业类型与基础信息 =================
st.write("---")
st.subheader("1. 作业类型与基础信息")

work_type_options = list(GB30871_SPECIFIC_ITEMS.keys())

col_w1, col_w2 = st.columns(2)
with col_w1:
  selected_work_type = st.selectbox(
      "特种作业类型 (依据国家标准下拉选择)：",
      options=work_type_options,
      index=0,
      help="切换作业类型后，下方将自动匹配对应的 GB 30871 现场安全评估要求",
  )
with col_w2:
  work_location = st.text_input(
      "具体作业地点 / 区域 (必填)：",
      placeholder="例如：一车间北侧维修区 / 配电房",
  )

col_u1, col_u2 = st.columns(2)
with col_u1:
  work_unit = st.text_input(
      "作业单位 / 部门 (必填)：",
      placeholder="例如：设备工程部 / 外部委外施工队",
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

# ================= 7. 模块二：签发日期与截止日期（8小时为限） =================
st.write("---")
st.subheader("2. 作业许可有效期限（严格以 8 小时为限）")

now_dt = datetime.datetime.now().replace(second=0, microsecond=0)
default_end_dt = now_dt + datetime.timedelta(hours=8)

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
elif duration_hours > 8.0:
  st.error(
      f"❌ 违反合规限值：当前设置的作业有效时长为 **{duration_hours:.1f} 小时**，已超过国家及公司规定的 **8 小时上限**！请缩短截止时间。"
  )
else:
  is_time_valid = True
  st.success(
      f"✅ 许可有效期核验通过：自 **{issue_datetime.strftime('%Y-%m-%d %H:%M')}** 至 **{expire_datetime.strftime('%Y-%m-%d %H:%M')}**（共计 **{duration_hours:.1f} 小时**，符合 ≤8 小时合规限值）。"
  )

# ================= 8. 模块三：现场安全评估（GB 30871-2022 标准打钩） =================
st.write("---")
st.subheader(f"3. 现场安全评估核查表（GB 30871-2022 - {selected_work_type}）")

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

# ================= 9. 模块四：三方责任人手动签字与手写日期 =================
st.write("---")
st.subheader("✍️ 4. 审批人、现场作业负责人、监护人 手动签字栏")

col_p1, col_p2, col_p3 = st.columns(3)
with col_p1:
  supervisor_name = st.text_input("现场作业负责人姓名 (必填)：")
with col_p2:
  guardian_name = st.text_input("安全监护人姓名 (必填)：")
with col_p3:
  approver_name = st.text_input("审批人姓名 (必填)：")

st.markdown("**请在下方对应画板内完成三方亲笔手写签字：**")

col_c1, col_c2, col_c3 = st.columns(3)
with col_c1:
  st.markdown("📝 **1. 现场作业负责人签字：**")
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
  st.markdown("🛡️ **2. 安全监护人签字：**")
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
  st.markdown("✅ **3. 审批人签字：**")
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
    f"**📅 手写签发日期确认栏（当前系统日期为 {current_date_str}，请在下方手写日期）：**"
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


# 判断画布是否为空的辅助函数
def is_canvas_blank(canvas_obj):
  return canvas_obj.image_data is None or (
      canvas_obj.json_data is not None
      and len(canvas_obj.json_data.get("objects", [])) == 0
  )


# ================= 10. 生成正式 Word《特种作业安全许可证》函数 =================
def generate_work_permit_docx(
    permit_no,
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
  doc = Document()

  # 标题
  p_title = doc.add_paragraph()
  r_title = p_title.add_run(f"【特种作业安全许可证 / 审批单 — {w_type}】")
  r_title.bold = True
  r_title.font.size = Pt(15)
  r_title.font.name = "华文宋体"
  r_title.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

  # 基本信息与有效期段落
  p_meta = doc.add_paragraph()
  r_meta = p_meta.add_run(
      f"许可证编号：{permit_no}        执行标准：GB 30871-2022\n"
      f"作业类型：{w_type}        作业地点：{w_loc}\n"
      f"作业单位：{w_unit}        作业人员及证号：{w_workers}\n"
      f"作业内容及工器具：{w_content}\n"
      f"许可签发时间：{dt_issue.strftime('%Y-%m-%d %H:%M')}\n"
      f"许可截止时间：{dt_expire.strftime('%Y-%m-%d %H:%M')} （有效时长：{dur_hrs:.1f} 小时，符合 ≤8 小时合规限值）\n"
      "-------------------------------------------------------------------------"
  )
  r_meta.font.size = Pt(10.5)
  r_meta.font.name = "华文宋体"
  r_meta.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

  # GB 30871 评估明细
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

  # 三方手写签名表格
  p_sig_head = doc.add_paragraph()
  r_sh = p_sig_head.add_run(
      "-------------------------------------------------------------------------\n"
      "二、三方安全责任确认与亲笔手写签字："
  )
  r_sh.bold = True
  r_sh.font.size = Pt(11)
  r_sh.font.name = "华文宋体"
  r_sh.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

  table = doc.add_table(rows=2, cols=2)
  table.autofit = False

  # 单元格 1：现场作业负责人
  c00 = table.cell(0, 0).paragraphs[0]
  r00 = c00.add_run(f"1. 现场作业负责人（{sup_name}）手写签字：\n")
  r00.font.name = "华文宋体"
  r00.font.size = Pt(10)
  r00.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")
  c00.add_run().add_picture(sup_img_io, width=Inches(1.6))
  sup_img_io.seek(0)

  # 单元格 2：安全监护人
  c01 = table.cell(0, 1).paragraphs[0]
  r01 = c01.add_run(f"2. 安全监护人（{grd_name}）手写签字：\n")
  r01.font.name = "华文宋体"
  r01.font.size = Pt(10)
  r01.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")
  c01.add_run().add_picture(grd_img_io, width=Inches(1.6))
  grd_img_io.seek(0)

  # 单元格 3：审批人
  c10 = table.cell(1, 0).paragraphs[0]
  r10 = c10.add_run(f"3. 审批人（{app_name}）手写签字：\n")
  r10.font.name = "华文宋体"
  r10.font.size = Pt(10)
  r10.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")
  c10.add_run().add_picture(app_img_io, width=Inches(1.6))
  app_img_io.seek(0)

  # 单元格 4：手写签发日期
  c11 = table.cell(1, 1).paragraphs[0]
  r11 = c11.add_run("4. 手写签发日期：\n")
  r11.font.name = "华文宋体"
  r11.font.size = Pt(10)
  r11.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")
  c11.add_run().add_picture(date_img_io, width=Inches(1.6))
  date_img_io.seek(0)

  buffer = io.BytesIO()
  doc.save(buffer)
  buffer.seek(0)
  return buffer


# ================= 11. 提交校验与生成合规档案 =================
st.write("---")
if st.button(
    "📁 确认安全措施已落实，一键签发《特种作业安全许可证》",
    use_container_width=True,
    type="primary",
):
  if (
      not work_location.strip()
      or not work_unit.strip()
      or not worker_names.strip()
      or not work_content.strip()
  ):
    st.error("❌ 拦截：请完整填写【作业地点】、【作业单位】、【作业人员】及【作业内容】！")
  elif not is_time_valid:
    st.error(
        "❌ 拦截：作业许可有效时间设置不合规（必须大于 0 且不超过 8 小时），请重新核定签发与截止时间！"
    )
  elif not all(assessment_results):
    st.warning(
        "⚠️ 拦截：按照 GB 30871-2022 强制性规范要求，现场所有安全评估项必须全部核实打钩后方可签发作业许可！"
    )
  elif (
      not supervisor_name.strip()
      or not guardian_name.strip()
      or not approver_name.strip()
  ):
    st.error("❌ 拦截：请完整填写【现场作业负责人】、【安全监护人】与【审批人】姓名！")
  elif is_canvas_blank(canvas_supervisor):
    st.warning("⚠️ 拦截：请【现场作业负责人】在对应画板内完成手写签字！")
  elif is_canvas_blank(canvas_guardian):
    st.warning("⚠️ 拦截：请【安全监护人】在对应画板内完成手写签字！")
  elif is_canvas_blank(canvas_approver):
    st.warning("⚠️ 拦截：请【审批人】在对应画板内完成手写签字！")
  elif is_canvas_blank(canvas_date):
    st.warning("⚠️ 拦截：请在【手写签发日期确认栏】内完成手写日期！")
  else:
    permit_number = f"PTW-{datetime.datetime.now().strftime('%Y%m%d-%H%M%S')}"
    st.success(
        f"✅ 特种作业安全许可证签发成功！许可证编号：**{permit_number}**"
    )

    # 转换四个手写画布为图片流
    def canvas_to_bytes_io(canvas_obj):
      img = Image.fromarray(canvas_obj.image_data.astype("uint8"), "RGBA")
      b_io = io.BytesIO()
      img.save(b_io, format="PNG")
      b_io.seek(0)
      return b_io

    sup_io = canvas_to_bytes_io(canvas_supervisor)
    grd_io = canvas_to_bytes_io(canvas_guardian)
    app_io = canvas_to_bytes_io(canvas_approver)
    dt_io = canvas_to_bytes_io(canvas_date)

    # 生成 Word 审批单
    permit_docx_buffer = generate_work_permit_docx(
        permit_number,
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

    docx_filename = f"特种作业审批单_{selected_work_type}_{work_location.strip()}_{permit_number}.docx"

    # 打包 ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
      zip_file.writestr(docx_filename, permit_docx_buffer.getvalue())
      zip_file.writestr(
          f"作业负责人签字_{supervisor_name.strip()}.png", sup_io.getvalue()
      )
      zip_file.writestr(
          f"监护人签字_{guardian_name.strip()}.png", grd_io.getvalue()
      )
      zip_file.writestr(
          f"审批人签字_{approver_name.strip()}.png", app_io.getvalue()
      )
      zip_file.writestr("手写签发日期.png", dt_io.getvalue())

      # 自动同步至百度网盘
      upload_to_baidu_netdisk_with_auto_refresh(
          permit_docx_buffer.getvalue(), docx_filename
      )

    zip_buffer.seek(0)

    st.markdown("---")
    st.success("🎉 特种作业审批档案已生成完毕，点击下方按钮即可下载归档：")

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
      st.download_button(
          label="📄 下载《特种作业安全许可证》(.docx)",
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

# ================= 12. 底部版权与开发者声明 =================
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray; font-size: 14px;'>"
    "本系统为内部合规平台，严禁商业用途 | 开发者：陈野菲"
    "</div>",
    unsafe_allow_html=True,
)