from flask import Flask, request, jsonify, render_template_string, session, Response
import requests
import os
import secrets
from datetime import datetime, timedelta, date
from ics import Calendar, Event
import re # 引入正則表達式模組

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

# --- 前端 HTML 和 CSS 完全保持不變 ---
html = '''
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SCU 課表查詢系統</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-color: #0d1117; --surface-color: #161b22; --primary-accent-color: #c9a45d; --secondary-accent-color: #58a6ff;
            --text-color: #c9d1d9; --text-muted-color: #8b949e; --border-color: #30363d; --hover-glow: 0 0 20px rgba(201, 164, 93, 0.4);
            --error-bg: rgba(248, 81, 73, 0.1); --error-text: #f85149; --success-bg: rgba(63, 185, 80, 0.1); --success-text: #3fb950;
        }
        body {
            font-family: 'Noto Sans TC', Arial, sans-serif; background-color: var(--bg-color); color: var(--text-color);
            margin: 0; padding: 2rem; display: flex; flex-direction: column; align-items: center; min-height: 100vh; box-sizing: border-box;
        }
        .content-wrapper {
            width: 100%; max-width: 1200px; padding: 2.5rem; background-color: var(--surface-color); border: 1px solid var(--border-color);
            border-radius: 12px; box-shadow: 0 8px 24px rgba(0,0,0,0.4); transition: all 0.5s ease-out;
        }
        h2, h3 { color: var(--secondary-accent-color); text-align: center; margin: 0 0 2rem 0; font-weight: 500; }
        .form-group { margin-bottom: 1.5rem; }
        label { display: block; margin-bottom: 0.5rem; font-weight: 400; color: var(--text-muted-color); }
        input[type="text"], input[type="password"] {
            width: 100%; padding: 12px 15px; background-color: var(--bg-color); border: 1px solid var(--border-color);
            border-radius: 6px; box-sizing: border-box; color: var(--text-color); font-size: 16px; transition: border-color 0.3s ease, box-shadow 0.3s ease;
        }
        input:focus { outline: none; border-color: var(--primary-accent-color); box-shadow: 0 0 8px rgba(201, 164, 93, 0.3); }
        button, .styled-button {
            background-color: var(--primary-accent-color); color: var(--bg-color); padding: 12px 25px; border: none; border-radius: 6px; cursor: pointer;
            font-size: 16px; font-weight: 700; width: 100%; transition: all 0.3s ease; text-decoration: none; display: inline-block; box-sizing: border-box; text-align: center;
        }
        button:hover, .styled-button:hover { background-color: #e6bf7a; transform: translateY(-2px); box-shadow: var(--hover-glow); }
        button:disabled, .styled-button.disabled { background-color: #8b949e; cursor: not-allowed; transform: none; box-shadow: none; }
        .message { margin-top: 1.5rem; padding: 12px; border-radius: 6px; text-align: center; font-weight: 500; }
        .success { background-color: var(--success-bg); color: var(--success-text); }
        .error { background-color: var(--error-bg); color: var(--error-text); }
        .loading { color: var(--text-muted-color); }

        #courseContent { display: none; opacity: 0; transform: translateY(20px); transition: opacity 0.8s ease-out, transform 0.8s ease-out; }
        #courseContent.visible { display: block; opacity: 1; transform: translateY(0); }
        #mobileControls { display: none; flex-direction: column; gap: 0.5rem; margin-bottom: 1.5rem; }
        #mobileControls .day-selector { display: flex; justify-content: space-between; align-items: center; }
        #mobileControls button { padding: 8px 12px; font-size: 14px; width: auto; flex-grow: 1; }
        #mobileControls #prevDay, #mobileControls #nextDay { flex-grow: 0; width: 50px; }
        #currentDayDisplay { color: var(--primary-accent-color); font-size: 1.2em; font-weight: 700; text-align: center; flex-grow: 2; }

        /* --- CSS Grid 表格 --- */
        #courseData { margin-top: 1rem; }
        .course-grid {
            display: grid; grid-template-columns: 129px repeat(7, 1fr); gap: 4px;
            min-width: 900px; box-sizing: border-box; font-size: 0.85em;
        }
        .grid-cell {
            padding: 8px; border-radius: 6px; background-color: var(--surface-color);
            display: flex; align-items: center; justify-content: center;
            text-align: center; min-height: 60px; transition: all 0.3s ease-in-out;
            overflow: hidden; text-overflow: ellipsis;
        }
        .grid-header, .grid-time-header {
            color: var(--secondary-accent-color); font-weight: 500; background-color: var(--bg-color);
            position: sticky; top: 0; z-index: 2;
        }
        .grid-time-header { left: 0; z-index: 3; }
        .grid-slot-time { color: var(--text-muted-color); font-weight: 500; background-color: var(--bg-color); position: sticky; left: 0; z-index: 1; }
        .grid-course a { color: var(--text-color); text-decoration: none; display: flex; align-items: center; justify-content: center; width: 100%; height: 100%; }
        .grid-course.has-course { cursor: pointer; }
        .grid-course.has-course:hover { transform: scale(1.05); background-color: var(--primary-accent-color); box-shadow: var(--hover-glow); z-index: 10; }
        .grid-course.has-course:hover a { color: var(--bg-color); font-weight: 500; }
        .grid-course.empty { background-color: transparent; border: 1px dashed var(--border-color); }
        #courseTableTitle { grid-column: 1 / -1; text-align: center; font-size: 1.2em; letter-spacing: 1px; color: var(--text-color); padding-bottom: 1rem; background-color: var(--surface-color); border-radius: 6px; }

        /* --- 可折疊空行的樣式 --- */
        .row-expandable { cursor: pointer; position: relative; }
        .row-expandable::after { content: '+'; position: absolute; right: 10px; top: 50%; transform: translateY(-50%); font-size: 1.5em; color: var(--text-muted-color); transition: transform 0.3s ease; }
        .row-expandable.expanded::after { content: '−'; }
        .is-collapsed-merged { grid-column: 1 / -1 !important; height: 25px !important; min-height: 25px !important; border: 1px solid var(--border-color) !important; }
        .cell-hidden-by-collapse { display: none !important; }
        .grid-slot-time .content-collapsed { display: none; }
        .grid-slot-time .content-normal { display: block; }
        .grid-slot-time.is-collapsed-merged .content-normal { display: none; }
        .grid-slot-time.is-collapsed-merged .content-collapsed { display: inline; }

        /* --- 手機版響應式 --- */
        @media (max-width: 768px) {
            body { padding: 0; }
            .content-wrapper { max-width: 100%; border-radius: 0; padding: 1.5rem 1rem; border: none; box-shadow: none; }
            #mobileControls { display: flex; }
            .course-grid { min-width: unset; }
            .day-hidden { display: none !important; }
            .course-grid.mobile-full-view { grid-template-columns: 35px repeat(7, 1fr); gap: 2px; font-size: 0.6em; }
            .mobile-full-view .grid-cell { padding: 3px 2px; min-height: unset; overflow: visible; text-overflow: clip; word-break: break-all; line-height: 1.2; }
            .mobile-full-view .grid-course.has-course:hover { transform: none; }
            .row-expandable::after { right: 5px; }
        }

        /* --- 導出按鈕樣式 --- */
        .export-buttons { display: none; flex-direction: column; gap: 1rem; margin-top: 2rem; padding-top: 2rem; border-top: 1px solid var(--border-color); }
        @media (min-width: 768px) { .export-buttons { flex-direction: row; } .export-buttons > * { flex: 1; } }

        /* --- 打印/導出專用樣式 --- */
        .is-printing { background-color: #ffffff !important; min-width: 1200px !important; }
        .is-printing .grid-cell { background-color: #ffffff !important; color: #000000 !important; border: 1px solid #ddd !important; }
        .is-printing .grid-slot-time, .is-printing .grid-header, .is-printing .grid-time-header { background-color: #f2f2f2 !important; color: #000000 !important; }
        .is-printing .grid-course a { color: #000000 !important; }
        .is-printing .row-expandable::after { display: none !important; }
        .is-printing .grid-header, .is-printing .grid-time-header, .is-printing .grid-slot-time { position: static !important; }
    </style>
</head>
<body>
    <div class="content-wrapper">
        <h2 id="mainTitle">登入東吳大學系統</h2>
        <form id="loginForm">
            <div class="form-group"> <label for="userid">學號:</label> <input type="text" id="userid" name="userid" required autocomplete="username"> </div>
            <div class="form-group"> <label for="password">密碼:</label> <input type="password" id="password" name="password" required autocomplete="current-password"> </div>
            <button type="submit" id="loginBtn">登入並查詢課表</button>
        </form>
        <div id="message"></div>
        <div id="courseContent">
            <div id="mobileControls">
                <div class="day-selector"> <button id="prevDay"><</button> <span id="currentDayDisplay"></span> <button id="nextDay">></button> </div>
                <button id="toggleViewBtn">顯示整週</button>
            </div>
            <div id="courseData"></div>
            <div class="export-buttons" id="exportContainer">
                <button id="exportPngBtn">導出為 PNG</button>
                <button id="exportPdfBtn">導出為 PDF</button>
                <a href="/api/export/ics" class="styled-button" id="exportIcsBtn">添加到日曆 (.ics)</a>
            </div>
        </div>
    </div>

    <script>
        // JavaScript 邏輯完全保持不變
        let currentDayIndex = 0;
        let currentViewMode = 'today';
        const dayNames = ['週一', '週二', '週三', '週四', '週五', '週六', '週日'];
        function setupMobileView(){const t=window.innerWidth<=768,e=document.getElementById("mobileControls");if(!document.querySelector(".course-grid"))return;e.style.display=t?"flex":"none",t?("full-mobile"!==currentViewMode&&(currentDayIndex=(new Date).getDay(),currentDayIndex=0===currentDayIndex?6:currentDayIndex-1,currentViewMode="today")):currentViewMode="full-desktop",updateTableView()}function updateTableView(){const t=window.innerWidth<=768,e=document.querySelector(".course-grid");if(!e)return;const o=e.querySelectorAll(".grid-slot-time"),n=e.querySelectorAll("[data-day-index]");e.classList.remove("mobile-full-view"),e.style.gridTemplateColumns="",o.forEach(t=>{t.classList.remove("is-collapsed-merged","row-expandable","expanded")}),e.querySelectorAll(".grid-course").forEach(t=>{t.classList.remove("cell-hidden-by-collapse")}),n.forEach(t=>t.classList.remove("day-hidden")),t&&"today"===currentViewMode?(document.getElementById("toggleViewBtn").textContent="顯示整週",e.style.gridTemplateColumns="35px 1fr",n.forEach(t=>{t.classList.toggle("day-hidden",t.dataset.dayIndex!=currentDayIndex)})):t&&"full-mobile"===currentViewMode?(document.getElementById("toggleViewBtn").textContent="僅顯示今日",e.classList.add("mobile-full-view")):document.getElementById("toggleViewBtn").textContent="僅顯示今日",o.forEach(o=>{const n=o.dataset.slotIndex;let d=!1;t&&"today"===currentViewMode?e.querySelector(`.grid-course[data-slot-index="${n}"][data-day-index="${currentDayIndex}"]`)?.dataset.isEmpty==="true"&&(d=!0):"true"===o.dataset.isWeekEmpty&&(d=!0),d&&(o.classList.add("row-expandable","is-collapsed-merged"),e.querySelectorAll(`.grid-course[data-slot-index="${n}"]`).forEach(t=>{t.classList.add("cell-hidden-by-collapse")}))});const d=document.getElementById("currentDayDisplay");d.textContent=dayNames[currentDayIndex],document.getElementById("prevDay").disabled=0===currentDayIndex&&"today"===currentViewMode,document.getElementById("nextDay").disabled=6===currentDayIndex&&"today"===currentViewMode}async function performExport(t){const e=document.querySelector(".course-grid"),o=`export${t}Btn`,n=document.getElementById(o),d=n.textContent;n.textContent="生成中...",n.disabled=!0;const s=currentViewMode;e.classList.add("is-printing"),e.classList.remove("mobile-full-view"),e.style.gridTemplateColumns="",e.querySelectorAll(".is-collapsed-merged").forEach(t=>{t.classList.remove("is-collapsed-merged","expanded");const o=t.dataset.slotIndex;e.querySelectorAll(`.grid-course[data-slot-index="${o}"]`).forEach(t=>t.classList.remove("cell-hidden-by-collapse"))}),e.querySelectorAll(".day-hidden").forEach(t=>t.classList.remove("day-hidden"));try{await new Promise(t=>setTimeout(t,100));const l=await html2canvas(e,{scale:2,useCORS:!0,backgroundColor:"#ffffff"});if("Png"===t){const i=document.createElement("a");i.download="course_schedule.png",i.href=l.toDataURL("image/png"),i.click()}else if("Pdf"===t){const{jsPDF:a}=window.jspdf,c=l.toDataURL("image/png"),r=new a({orientation:l.width>l.height?"landscape":"portrait",unit:"px",format:[l.width,l.height]});r.addImage(c,"PNG",0,0,l.width,l.height),r.save("course_schedule.pdf")}}catch(u){console.error("Export failed:",u),alert("導出失敗，請查看控制台日誌。")}finally{e.classList.remove("is-printing"),updateTableView(),n.textContent=d,n.disabled=!1}}document.addEventListener("DOMContentLoaded",()=>{document.getElementById("courseData").addEventListener("click",t=>{const e=t.target.closest(".row-expandable");if(!e)return;const o=e.dataset.slotIndex,n=document.getElementById("courseData").querySelectorAll(`.grid-course[data-slot-index="${o}"]`);e.classList.toggle("expanded"),e.classList.toggle("is-collapsed-merged"),n.forEach(t=>t.classList.toggle("cell-hidden-by-collapse"))}),document.getElementById("exportPngBtn").addEventListener("click",()=>performExport("Png")),document.getElementById("exportPdfBtn").addEventListener("click",()=>performExport("Pdf")),document.getElementById("prevDay").addEventListener("click",()=>{0<currentDayIndex&&(currentDayIndex--,updateTableView())}),document.getElementById("nextDay").addEventListener("click",()=>{6>currentDayIndex&&(currentDayIndex++,updateTableView())}),document.getElementById("toggleViewBtn").addEventListener("click",()=>{window.innerWidth<=768&&(currentViewMode="today"===currentViewMode?"full-mobile":"today",updateTableView())}),window.addEventListener("resize",setupMobileView)}),document.getElementById("loginForm").addEventListener("submit",async function(t){t.preventDefault();const e=document.getElementById("userid").value,o=document.getElementById("password").value,n=document.getElementById("loginBtn"),d=document.getElementById("message");n.disabled=!0,n.textContent="登入中...",d.innerHTML='<div class="loading">正在登入...</div>',document.getElementById("courseContent").classList.remove("visible");try{const s=await fetch("/api/login",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({userid:e,password:o})}),l=await s.json();"success"===l.status?(d.innerHTML='<div class="success">登入成功！正在獲取課表...</div>',await getCourseTable(l.data)):(d.innerHTML=`<div class="error">登入失敗: ${l.message}</div>`,n.disabled=!1,n.textContent="登入並查詢課表")}catch(i){d.innerHTML=`<div class="error">發生錯誤: ${i.message}</div>`,n.disabled=!1,n.textContent="登入並查詢課表"}});async function getCourseTable(t){const e=document.getElementById("message");try{const o=await fetch("/api/course",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(t)}),n=await o.json();if("success"===n.status){document.getElementById("loginForm").style.display="none",document.getElementById("mainTitle").textContent="您的課表",e.innerHTML='<div class="success">課表獲取成功！</div>';const d=document.getElementById("courseContent");document.getElementById("courseData").innerHTML=n.content,d.classList.add("visible"),document.getElementById("exportContainer").style.display="flex",setupMobileView()}else e.innerHTML=`<div class="error">課表獲取失敗: ${n.message}</div>`}catch(s){e.innerHTML=`<div class="error">課表獲取錯誤: ${s.message}</div>`}finally{const l=document.getElementById("loginBtn");l.disabled=!1,l.textContent="重新查詢"}}
    </script>
</body>
</html>
'''

# 原始系統的基礎 URL
BASE_URL = "https://psv.scu.edu.tw/portal"

@app.route('/')
def index():
    return render_template_string(html)

@app.route('/api/login', methods=['POST'])
def api_login():
    # 登入邏輯保持不變
    data = request.get_json()
    userid, password = data.get('userid'), data.get('password')
    url = f"{BASE_URL}/jsonApi.php"
    payload = { "libName": "Login", "userid": userid, "password": password }
    headers = { "Content-Type": "application/x-www-form-urlencoded", "User-Agent": "Mozilla/5.0", "Referer": BASE_URL }
    try:
        response = requests.post(url, data=payload, headers=headers)
        response.raise_for_status()
        response_data = response.json()
        if response_data.get('status') == 'success':
            user_info = response_data.get('message', {})
            return jsonify({ "status": "success", "message": "登入成功", "data": { "sessionID": user_info.get('sessionID'), "userId": user_info.get('userId'), "sessionCode": user_info.get('sessionCode'), "name": user_info.get('name'), "unit": user_info.get('unit'), }})
        else: return jsonify({"status": "error", "message": response_data.get('message', '登入失敗')})
    except requests.RequestException as e: return jsonify({"status": "error", "message": f"連線錯誤: {str(e)}"}), 500
    except Exception as e: return jsonify({"status": "error", "message": f"處理登入回傳失敗: {str(e)}"}), 500

def cours_table_td_data(s):
    if s is None: return ''
    s = s.replace(' ', '').replace('��', '').replace('<br/>', '<br>').replace('<br><br>', '<br>')
    return s.strip()

def process_course_data(raw_data):
    num_slots = len(raw_data)
    temp_grid = [[{'course_id': '', 'course_text': '', 'span': 0, 'raw_text': ''} for _ in range(7)] for _ in range(num_slots)]
    is_row_week_empty = [True] * num_slots
    for slot_idx, slot_data in enumerate(raw_data):
        for day_idx in range(7):
            api_day_key, api_courid_key = f'day{day_idx + 1}', f'day{day_idx + 1}Courid'
            raw_text = slot_data.get(api_day_key, '')
            course_text = cours_table_td_data(raw_text)
            course_id = slot_data.get(api_courid_key, '')
            if course_id and course_text.strip(): is_row_week_empty[slot_idx] = False
            temp_grid[slot_idx][day_idx] = {'course_id': course_id, 'course_text': course_text, 'span': 1, 'raw_text': raw_text}
    for day_idx in range(7):
        for slot_idx in range(num_slots - 2, -1, -1):
            current_cell_data, next_cell_data = temp_grid[slot_idx][day_idx], temp_grid[slot_idx + 1][day_idx]
            if (current_cell_data['course_id'] and current_cell_data['course_id'] == next_cell_data['course_id'] and
                current_cell_data['raw_text'] == next_cell_data['raw_text'] and next_cell_data['span'] > 0):
                current_cell_data['span'] += next_cell_data['span']
                next_cell_data['span'] = 0
                if is_row_week_empty[slot_idx]: is_row_week_empty[slot_idx] = False
    return temp_grid, is_row_week_empty

@app.route('/api/course', methods=['POST'])
def api_course():
    data = request.get_json()
    session_id, user_id, session_code, user_name, user_unit = data.get('sessionID'), data.get('userId'), data.get('sessionCode'), data.get('name'), data.get('unit')
    if not all([session_id, user_id, session_code, user_name, user_unit]):
        return jsonify({"status": "error", "message": "缺少必要的登入資訊來獲取課表"}), 400
    course_api_url = f"{BASE_URL}/jsonApi.php"
    course_payload = { "libName": "CourseTable", "api_loginstr": session_id, "api_loginID": user_id, "api_encodeID": session_code, "api_stuname": user_name, "api_clsname": user_unit }
    headers = { "Content-Type": "application/x-www-form-urlencoded", "User-Agent": "Mozilla/5.0", "Referer": BASE_URL }
    try:
        response = requests.post(course_api_url, data=course_payload, headers=headers)
        response.raise_for_status()
        course_data = response.json()
        if course_data.get('status') == 'success':
            message_data = course_data.get('message', {})
            sub_result = message_data.get('SubRESULT', [])
            time_info = message_data.get('time', '').strip().replace(' ', '')
            year, semester = int(time_info[0:3]), int(time_info[6])
            session['course_data'] = { 'sub_result': sub_result, 'year': year, 'semester': semester }
            temp_grid, is_row_week_empty = process_course_data(sub_result)
            course_table_title = f"{year} 學年度 第 {semester} 學期"
            course_time = ['08:10 <br> 09:00', '09:10 <br> 10:00', '10:10 <br> 11:00', '11:10 <br> 12:00', '12:10 <br> 13:00', '13:10 <br> 14:00', '14:10 <br> 15:00', '15:10 <br> 16:00', '16:10 <br> 17:00', '17:10 <br> 18:20', '18:25 <br> 19:15', '19:20 <br> 20:10', '20:20 <br> 21:10', '21:15 <br> 22:05']
            week_days = ['週一 <br> Mon', '週二 <br> TUE', '週三 <br> WED', '週四 <br> THU', '週五 <br> FRI', '週六 <br> SAT', '週日 <br> SUN']
            num_slots = len(sub_result)
            grid_html = '<div class="course-grid">'
            grid_html += f'<div id="courseTableTitle">{course_table_title}</div>'
            grid_html += '<div class="grid-cell grid-time-header"></div>'
            for i, day_name in enumerate(week_days):
                grid_html += f'<div class="grid-cell grid-header" data-day-index="{i}">{day_name}</div>'
            for slot_idx in range(num_slots):
                is_week_empty_attr = 'true' if is_row_week_empty[slot_idx] else 'false'
                slot_label = sub_result[slot_idx].get("slot", "")
                time_period_text = course_time[slot_idx] if 0 <= slot_idx < len(course_time) else ""
                content_normal = f'<span class="content-normal">{slot_label}<br>{time_period_text}</span>'
                content_collapsed = f'<span class="content-collapsed">{slot_label} {time_period_text.replace("<br>", "-")}</span>'
                grid_html += f'<div class="grid-cell grid-slot-time" data-slot-index="{slot_idx}" data-is-week-empty="{is_week_empty_attr}">{content_normal}{content_collapsed}</div>'
                for day_idx in range(7):
                    cell_to_render = temp_grid[slot_idx][day_idx]
                    if cell_to_render['span'] > 0:
                        course_text, course_id = cell_to_render['course_text'], cell_to_render['course_id']
                        rowspan_attr = f'style="grid-row-end: span {cell_to_render["span"]};"' if cell_to_render['span'] > 1 else ''
                        is_empty_attr = 'true' if not (course_id and course_text.strip()) else 'false'
                        if is_empty_attr == 'false':
                            link_url = f"https://mobile.sys.scu.edu.tw/performance/performance/{year}/{semester}/{course_id}"
                            grid_html += f'<div class="grid-cell grid-course has-course" data-is-empty="false" data-slot-index="{slot_idx}" data-day-index="{day_idx}" {rowspan_attr}><a href="{link_url}" target="_blank">{course_text}</a></div>'
                        else:
                            grid_html += f'<div class="grid-cell grid-course empty" data-is-empty="true" data-slot-index="{slot_idx}" data-day-index="{day_idx}" {rowspan_attr}> </div>'
            grid_html += '</div>'
            return jsonify({"status": "success", "content": grid_html})
        else:
            return jsonify({"status": "error", "message": course_data.get('message', '獲取課表失敗')})
    except Exception as e:
        return jsonify({"status": "error", "message": f"處理課表數據失敗: {str(e)}"}), 500

# --- 新增的 ICS 導出路由 ---

# **新增**：智能解析課程字串的輔助函數
def parse_course_details(text):
    details = {
        'name': text,
        'location': None,
        'teacher': None,
        'week_type': None # 'odd', 'even', or None
    }
    
    # 1. 移除 HTML 標籤並清理
    clean_text = re.sub(r'<br\s*/?>', ' ', text).strip()
    
    # 2. 提取教師 (通常是結尾的2-4個中文字)
    teacher_match = re.search(r'([\u4e00-\u9fa5]{2,4})$', clean_text)
    if teacher_match:
        details['teacher'] = teacher_match.group(1).strip()
        clean_text = clean_text[:teacher_match.start()].strip()
        
    # 3. 提取單雙週
    if '雙' in clean_text:
        details['week_type'] = 'even'
        clean_text = clean_text.replace('雙', '').strip()
    elif '單' in clean_text:
        details['week_type'] = 'odd'
        clean_text = clean_text.replace('單', '').strip()

    # 4. 提取地點 (常見模式：D0309, U19, 0141)
    location_match = re.search(r'([A-Z]?\d{3,4})$', clean_text)
    if location_match:
        details['location'] = location_match.group(1).strip()
        clean_text = clean_text[:location_match.start()].strip()

    # 5. 剩餘部分作為課程名稱
    details['name'] = clean_text.strip() or "未命名課程"

    return details


@app.route('/api/export/ics')
def export_ics():
    if 'course_data' not in session:
        return "錯誤：課表資訊不存在。請先查詢課表。", 400

    course_data = session['course_data']
    sub_result, year, semester = course_data['sub_result'], course_data['year'], course_data['semester']
    
    temp_grid, _ = process_course_data(sub_result)
    
    gregorian_year = year + 1911
    start_date = date(gregorian_year, 2, 12) if semester == 2 else date(gregorian_year, 9, 10)
    start_weekday = start_date.weekday()
    end_date = start_date + timedelta(weeks=18)

    course_time_mapping = {
        '1': ("08:10", "09:00"), '2': ("09:10", "10:00"), '3': ("10:10", "11:00"), '4': ("11:10", "12:00"), 'E': ("12:10", "13:00"), 
        '5': ("13:10", "14:00"), '6': ("14:10", "15:00"), '7': ("15:10", "16:00"), '8': ("16:10", "17:00"), '9': ("17:10", "18:20"), 
        'A': ("18:25", "19:15"), 'B': ("19:20", "20:10"), 'C': ("20:20", "21:10"), 'D': ("21:15", "22:05")
    }

    cal = Calendar()
    for slot_idx in range(len(temp_grid)):
        for day_idx in range(7):
            cell = temp_grid[slot_idx][day_idx]
            if cell['span'] > 0 and cell['course_id']:
                slot_label = sub_result[slot_idx].get("slot")
                end_slot_idx = slot_idx + cell['span'] - 1
                end_slot_label = sub_result[end_slot_idx].get("slot")

                if slot_label not in course_time_mapping or end_slot_label not in course_time_mapping:
                    continue

                start_hm = course_time_mapping[slot_label][0].split(':')
                end_hm = course_time_mapping[end_slot_label][1].split(':')

                # **使用新的解析函數**
                details = parse_course_details(cell['course_text'])

                # 計算第一次上課的日期
                days_to_add = day_idx - start_weekday
                first_class_date = start_date + timedelta(days=days_to_add)

                # 處理單雙週的開始日期
                # isocalendar()[1] 返回年份中的第幾週
                start_week_num = first_class_date.isocalendar()[1]
                if details['week_type'] == 'even' and start_week_num % 2 != 0:
                    first_class_date += timedelta(weeks=1) # 如果是雙週課但開始於單數週，則推遲一週
                elif details['week_type'] == 'odd' and start_week_num % 2 == 0:
                    first_class_date += timedelta(weeks=1) # 如果是單週課但開始於雙數週，則推遲一週
                
                # 如果調整後超出學期，則跳過
                if first_class_date > end_date:
                    continue

                begin_time = first_class_date.strftime(f'%Y-%m-%d {start_hm[0]}:{start_hm[1]}:00')
                end_time = first_class_date.strftime(f'%Y-%m-%d {end_hm[0]}:{end_hm[1]}:00')
                
                description_parts = []
                if details['teacher']: description_parts.append(f"教師: {details['teacher']}")
                if details['week_type']: description_parts.append(f"週別: {'雙週' if details['week_type'] == 'even' else '單週'}")

                e = Event()
                e.name = details['name']
                e.begin = begin_time
                e.end = end_time
                e.location = details['location'] or ''
                e.description = "\n".join(description_parts)
                
                # 設置重複規則
                rrule_parts = ['FREQ=WEEKLY']
                if details['week_type']:
                    rrule_parts.append('INTERVAL=2') # 單雙週都是每兩週重複一次
                rrule_parts.append(f'UNTIL={end_date.strftime("%Y%m%d")}T235959Z')
                e.rrule = ";".join(rrule_parts)
                
                cal.events.add(e)

    return Response(str(cal), mimetype="text/calendar", headers={"Content-disposition": "attachment; filename=course_schedule.ics"})


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host='0.0.0.0', port=port)