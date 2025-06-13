from flask import Flask, request, jsonify, render_template_string, session, Response
import requests
import os
import secrets
from datetime import datetime, timedelta, date
from ics import Calendar, Event
import re
import pytz

app = Flask(__name__)
# 優先從環境變數讀取 SECRET_KEY，如果沒有就隨機生成一個 (方便本地測試)
app.secret_key = os.environ.get('SECRET_KEY', secrets.token_hex(16))
# --- 前端 HTML/CSS/JS 保持不變 ---
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
        body { font-family: 'Noto Sans TC', Arial, sans-serif; background-color: var(--bg-color); color: var(--text-color); margin: 0; padding: 2rem; display: flex; flex-direction: column; align-items: center; min-height: 100vh; box-sizing: border-box; }
        .content-wrapper { width: 100%; max-width: 1200px; padding: 2.5rem; background-color: var(--surface-color); border: 1px solid var(--border-color); border-radius: 12px; box-shadow: 0 8px 24px rgba(0,0,0,0.4); transition: all 0.5s ease-out; }
        h2, h3 { color: var(--secondary-accent-color); text-align: center; margin: 0 0 2rem 0; font-weight: 500; }
        .form-group { margin-bottom: 1.5rem; }
        label { display: block; margin-bottom: 0.5rem; font-weight: 400; color: var(--text-muted-color); }
        input[type="text"], input[type="password"] { width: 100%; padding: 12px 15px; background-color: var(--bg-color); border: 1px solid var(--border-color); border-radius: 6px; box-sizing: border-box; color: var(--text-color); font-size: 16px; transition: border-color 0.3s ease, box-shadow 0.3s ease; }
        input:focus { outline: none; border-color: var(--primary-accent-color); box-shadow: 0 0 8px rgba(201, 164, 93, 0.3); }
        button, .styled-button { background-color: var(--primary-accent-color); color: var(--bg-color); padding: 12px 25px; border: none; border-radius: 6px; cursor: pointer; font-size: 16px; font-weight: 700; width: 100%; transition: all 0.3s ease; text-decoration: none; display: inline-block; box-sizing: border-box; text-align: center; }
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
        #courseData { margin-top: 1rem; }
        .course-grid { display: grid; grid-template-columns: 129px repeat(7, 1fr); gap: 4px; min-width: 900px; box-sizing: border-box; font-size: 0.85em; }
        .grid-cell { padding: 8px; border-radius: 6px; background-color: var(--surface-color); display: flex; align-items: center; justify-content: center; text-align: center; min-height: 60px; transition: all 0.3s ease-in-out; overflow: hidden; text-overflow: ellipsis; }
        .grid-header, .grid-time-header { color: var(--secondary-accent-color); font-weight: 500; background-color: var(--bg-color); position: sticky; top: 0; z-index: 2; }
        .grid-time-header { left: 0; z-index: 3; }
        .grid-slot-time { color: var(--text-muted-color); font-weight: 500; background-color: var(--bg-color); position: sticky; left: 0; z-index: 1; }
        .grid-course a { color: var(--text-color); text-decoration: none; display: flex; align-items: center; justify-content: center; width: 100%; height: 100%; }
        .grid-course.has-course { cursor: pointer; }
        .grid-course.has-course:hover { transform: scale(1.05); background-color: var(--primary-accent-color); box-shadow: var(--hover-glow); z-index: 10; }
        .grid-course.has-course:hover a { color: var(--bg-color); font-weight: 500; }
        .grid-course.empty { background-color: transparent; border: 1px dashed var(--border-color); }
        #courseTableTitle { grid-column: 1 / -1; text-align: center; font-size: 1.2em; letter-spacing: 1px; color: var(--text-color); padding-bottom: 1rem; background-color: var(--surface-color); border-radius: 6px; }
        .row-expandable { cursor: pointer; position: relative; }
        .row-expandable::after { content: '+'; position: absolute; right: 10px; top: 50%; transform: translateY(-50%); font-size: 1.5em; color: var(--text-muted-color); transition: transform 0.3s ease; }
        .row-expandable.expanded::after { content: '−'; }
        .is-collapsed-merged { grid-column: 1 / -1 !important; height: 25px !important; min-height: 25px !important; border: 1px solid var(--border-color) !important; }
        .cell-hidden-by-collapse { display: none !important; }
        .grid-slot-time .content-collapsed { display: none; }
        .grid-slot-time .content-normal { display: block; }
        .grid-slot-time.is-collapsed-merged .content-normal { display: none; }
        .grid-slot-time.is-collapsed-merged .content-collapsed { display: inline; }
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
        .export-buttons { display: none; flex-direction: column; gap: 1rem; margin-top: 2rem; padding-top: 2rem; border-top: 1px solid var(--border-color); }
        @media (min-width: 768px) { .export-buttons { flex-direction: row; } .export-buttons > * { flex: 1; } }
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
                <button class="styled-button" id="exportIcsBtn">添加到日曆 (.ics)</button>
            </div>
        </div>
    </div>
<!-- 您只需要替換掉您 HTML 檔案中 <script> 到 </script> 的部分 -->
<script>
    let currentDayIndex = 0;
    let currentViewMode = "today";
    const dayNames = ["週一", "週二", "週三", "週四", "週五", "週六", "週日"];

    function setupMobileView() {
        const isMobile = window.innerWidth <= 768;
        const mobileControls = document.getElementById("mobileControls");
        if (!document.querySelector(".course-grid")) return;

        mobileControls.style.display = isMobile ? "flex" : "none";

        if (isMobile) {
            if (currentViewMode !== "full-mobile") {
                let today = new Date().getDay();
                currentDayIndex = (today === 0) ? 6 : today - 1;
                currentViewMode = "today";
            }
        } else {
            currentViewMode = "full-desktop";
        }
        updateTableView();
    }

    function updateTableView() {
        const isMobile = window.innerWidth <= 768;
        const grid = document.querySelector(".course-grid");
        if (!grid) return;

        const timeSlots = grid.querySelectorAll(".grid-slot-time");
        const allCells = grid.querySelectorAll("[data-day-index]");

        // Reset all view-specific classes and styles
        grid.classList.remove("mobile-full-view");
        grid.style.gridTemplateColumns = "";
        timeSlots.forEach(slot => {
            slot.classList.remove("is-collapsed-merged", "row-expandable", "expanded");
        });
        grid.querySelectorAll(".grid-course").forEach(cell => {
            cell.classList.remove("cell-hidden-by-collapse");
        });
        allCells.forEach(cell => cell.classList.remove("day-hidden"));

        // Apply view-specific logic
        if (isMobile && currentViewMode === "today") {
            document.getElementById("toggleViewBtn").textContent = "顯示整週";
            grid.style.gridTemplateColumns = "35px 1fr";
            allCells.forEach(cell => {
                cell.classList.toggle("day-hidden", cell.dataset.dayIndex != currentDayIndex);
            });
        } else if (isMobile && currentViewMode === "full-mobile") {
            document.getElementById("toggleViewBtn").textContent = "僅顯示今日";
            grid.classList.add("mobile-full-view");
        } else { // Desktop view
             document.getElementById("toggleViewBtn").textContent = "僅顯示今日";
        }
        
        // Collapse empty rows
        timeSlots.forEach(slot => {
            const slotIndex = slot.dataset.slotIndex;
            let shouldCollapse = false;
            if (isMobile && currentViewMode === "today") {
                // In mobile today view, collapse if THIS day is empty
                const dayCell = grid.querySelector(`.grid-course[data-slot-index="${slotIndex}"][data-day-index="${currentDayIndex}"]`);
                if (dayCell?.dataset.isEmpty === 'true') {
                    shouldCollapse = true;
                }
            } else {
                // In desktop or mobile full week view, collapse if the WHOLE week is empty
                if (slot.dataset.isWeekEmpty === 'true') {
                    shouldCollapse = true;
                }
            }

            if (shouldCollapse) {
                slot.classList.add("row-expandable", "is-collapsed-merged");
                grid.querySelectorAll(`.grid-course[data-slot-index="${slotIndex}"]`).forEach(cell => {
                    cell.classList.add("cell-hidden-by-collapse");
                });
            }
        });

        const currentDayDisplay = document.getElementById("currentDayDisplay");
        currentDayDisplay.textContent = dayNames[currentDayIndex];
        document.getElementById("prevDay").disabled = (currentDayIndex === 0 && currentViewMode === "today");
        document.getElementById("nextDay").disabled = (currentDayIndex === 6 && currentViewMode === "today");
    }

    async function performExport(type) {
        const grid = document.querySelector('.course-grid');
        const btnId = `export${type}Btn`;
        const button = document.getElementById(btnId);
        const originalText = button.textContent;
        button.textContent = "生成中...";
        button.disabled = true;

        // Prepare grid for printing
        grid.classList.add('is-printing');
        grid.classList.remove('mobile-full-view');
        grid.style.gridTemplateColumns = '';
        grid.querySelectorAll('.is-collapsed-merged').forEach(slot => {
            slot.classList.remove('is-collapsed-merged', 'expanded');
            const slotIndex = slot.dataset.slotIndex;
            grid.querySelectorAll(`.grid-course[data-slot-index="${slotIndex}"]`).forEach(cell => {
                cell.classList.remove('cell-hidden-by-collapse');
            });
        });
        grid.querySelectorAll('.day-hidden').forEach(cell => cell.classList.remove('day-hidden'));

        try {
            await new Promise(resolve => setTimeout(resolve, 100)); // Allow DOM to update
            const canvas = await html2canvas(grid, {
                scale: 2,
                useCORS: true,
                backgroundColor: '#ffffff'
            });

            if (type === 'Png') {
                const link = document.createElement('a');
                link.download = 'course_schedule.png';
                link.href = canvas.toDataURL('image/png');
                link.click();
            } else if (type === 'Pdf') {
                const { jsPDF } = window.jspdf;
                const imgData = canvas.toDataURL('image/png');
                const pdf = new jsPDF({
                    orientation: canvas.width > canvas.height ? 'landscape' : 'portrait',
                    unit: 'px',
                    format: [canvas.width, canvas.height]
                });
                pdf.addImage(imgData, 'PNG', 0, 0, canvas.width, canvas.height);
                pdf.save('course_schedule.pdf');
            }
        } catch (error) {
            console.error('Export failed:', error);
            alert('導出失敗，請查看控制台日誌。');
        } finally {
            // Revert grid to original state
            grid.classList.remove('is-printing');
            updateTableView(); // Re-apply current view settings
            button.textContent = originalText;
            button.disabled = false;
        }
    }

    document.addEventListener("DOMContentLoaded", () => {
        document.getElementById("courseData").addEventListener("click", event => {
            const expandableRow = event.target.closest(".row-expandable");
            if (!expandableRow) return;

            const slotIndex = expandableRow.dataset.slotIndex;
            const cellsToToggle = document.getElementById("courseData").querySelectorAll(`.grid-course[data-slot-index="${slotIndex}"]`);
            
            expandableRow.classList.toggle("expanded");
            expandableRow.classList.toggle("is-collapsed-merged");
            cellsToToggle.forEach(cell => cell.classList.toggle("cell-hidden-by-collapse"));
        });

        document.getElementById("exportPngBtn").addEventListener("click", () => performExport("Png"));
        document.getElementById("exportPdfBtn").addEventListener("click", () => performExport("Pdf"));
        
        // --- ▼▼▼ 這就是修改後的核心部分 ▼▼▼ ---
        document.getElementById("exportIcsBtn").addEventListener("click", function() {
            // 使用 window.location.href 來觸發下載
            // 這會讓瀏覽器在當前的 context 中發起請求，從而帶上 session cookie
            window.location.href = '/api/export/ics';
        });
        // --- ▲▲▲ 修改結束 ▲▲▲ ---

        document.getElementById("prevDay").addEventListener("click", () => {
            if (currentDayIndex > 0) {
                currentDayIndex--;
                updateTableView();
            }
        });
        document.getElementById("nextDay").addEventListener("click", () => {
            if (currentDayIndex < 6) {
                currentDayIndex++;
                updateTableView();
            }
        });
        document.getElementById("toggleViewBtn").addEventListener("click", () => {
            if (window.innerWidth <= 768) {
                currentViewMode = (currentViewMode === "today") ? "full-mobile" : "today";
                updateTableView();
            }
        });
        window.addEventListener("resize", setupMobileView);
    });

    document.getElementById("loginForm").addEventListener("submit", async function(event) {
        event.preventDefault();
        const userid = document.getElementById("userid").value;
        const password = document.getElementById("password").value;
        const loginBtn = document.getElementById("loginBtn");
        const messageDiv = document.getElementById("message");

        loginBtn.disabled = true;
        loginBtn.textContent = "登入中...";
        messageDiv.innerHTML = '<div class="loading">正在登入...</div>';
        document.getElementById("courseContent").classList.remove("visible");

        try {
            const response = await fetch("/api/login", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ userid: userid, password: password }),
            });
            const result = await response.json();

            if (result.status === "success") {
                messageDiv.innerHTML = '<div class="success">登入成功！正在獲取課表...</div>';
                await getCourseTable(result.data);
            } else {
                messageDiv.innerHTML = `<div class="error">登入失敗: ${result.message}</div>`;
                loginBtn.disabled = false;
                loginBtn.textContent = "登入並查詢課表";
            }
        } catch (error) {
            messageDiv.innerHTML = `<div class="error">發生錯誤: ${error.message}</div>`;
            loginBtn.disabled = false;
            loginBtn.textContent = "登入並查詢課表";
        }
    });

    async function getCourseTable(loginData) {
        const messageDiv = document.getElementById("message");
        try {
            const response = await fetch("/api/course", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(loginData),
            });
            const result = await response.json();
            if (result.status === "success") {
                document.getElementById("loginForm").style.display = "none";
                document.getElementById("mainTitle").textContent = "您的課表";
                messageDiv.innerHTML = '<div class="success">課表獲取成功！</div>';
                const courseContent = document.getElementById("courseContent");

                document.getElementById("courseData").innerHTML = result.content;
                courseContent.classList.add("visible");
                document.getElementById("exportContainer").style.display = "flex";
                setupMobileView();
            } else {
                messageDiv.innerHTML = `<div class="error">課表獲取失敗: ${result.message}</div>`;
            }
        } catch (error) {
            messageDiv.innerHTML = `<div class="error">課表獲取錯誤: ${error.message}</div>`;
        } finally {
            const loginBtn = document.getElementById("loginBtn");
            loginBtn.disabled = false;
            loginBtn.textContent = "重新查詢";
        }
    }
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
    s = s.replace(' ', ' ').replace('��', '').replace('<br/>', '<br>').replace('<br><br>', '<br>')
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
            session['course_data'] = { 'sub_result': sub_result }
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
                content_collapsed = f'<span class="content-collapsed">{slot_label} {time_period_text.replace("<br>", " - ")}</span>'
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

def normalize_slot(slot_str):
    if not slot_str: return ""
    full_width = "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ"
    half_width = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    translate_table = str.maketrans(full_width, half_width)
    return slot_str.upper().translate(translate_table)

# --- **修改後的 ICS 導出路由** ---
@app.route('/api/export/ics')
def export_ics():
    if 'course_data' not in session:
        return "錯誤：課表資訊不存在。請先查詢課表。", 400

    course_data = session['course_data']
    sub_result = course_data['sub_result']
    temp_grid, _ = process_course_data(sub_result)
    
    tz = pytz.timezone('Asia/Taipei')
    today = datetime.now(tz).date()
    
    # 判斷學期結束日期
    current_month = today.month
    if current_month >= 9 or current_month <= 1:  # 上學期 (9月-1月)
        semester_end = date(today.year + (1 if current_month >= 9 else 0), 1, 31)
    else:  # 下學期 (2月-7月)
        semester_end = date(today.year, 7, 31)
    
    # 找到本週一作為起始點
    start_of_this_week = today - timedelta(days=today.weekday())

    course_time_mapping = {
        '1': ("08:10", "09:00"), '2': ("09:10", "10:00"), '3': ("10:10", "11:00"), '4': ("11:10", "12:00"),
        'E': ("12:10", "13:00"), '5': ("13:10", "14:00"), '6': ("14:10", "15:00"), '7': ("15:10", "16:00"),
        '8': ("16:10", "17:00"), '9': ("17:10", "18:20"), 'A': ("18:25", "19:15"), 'B': ("19:20", "20:10"),
        'C': ("20:20", "21:10"), 'D': ("21:15", "22:05")
    }
    
    # 手動構建 ICS 內容以確保符合 RFC 5545 規範
    def fold_line(line):
        """按照 RFC 5545 規範進行 75 字元換行"""
        if len(line) <= 75:
            return line
        
        result = []
        while len(line) > 75:
            result.append(line[:75])
            line = ' ' + line[75:]  # 續行需要空格開頭
        if line:
            result.append(line)
        return '\r\n'.join(result)
    
    ics_lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//SCU Course Schedule//EN"
    ]
    
    # 計算需要創建事件的週數範圍
    current_week_start = start_of_this_week
    week_count = 0
    
    while current_week_start <= semester_end and week_count < 25:  # 最多25週防止無限循環
        for slot_idx, row in enumerate(temp_grid):
            for day_idx, cell in enumerate(row):
                if cell['span'] > 0 and cell['course_id']:
                    slot_label = normalize_slot(sub_result[slot_idx].get("slot"))
                    if not slot_label or slot_label not in course_time_mapping: 
                        continue

                    start_hm = course_time_mapping[slot_label][0].split(':')
                    
                    end_slot_idx = slot_idx + cell['span'] - 1
                    end_slot_label = normalize_slot(sub_result[end_slot_idx].get("slot"))
                    if not end_slot_label or end_slot_label not in course_time_mapping: 
                        continue
                    end_hm = course_time_mapping[end_slot_label][1].split(':')

                    course_date = current_week_start + timedelta(days=day_idx)
                    
                    # 檢查課程日期是否超過學期結束
                    if course_date > semester_end:
                        continue
                    
                    # 處理單雙週課程
                    raw_text = cell['raw_text']
                    week_number = course_date.isocalendar()[1]
                    
                    # 跳過不符合單雙週條件的課程
                    if '單' in raw_text and week_number % 2 == 0:
                        continue
                    elif '雙' in raw_text and week_number % 2 != 0:
                        continue
                    
                    begin_dt = tz.localize(datetime(
                        course_date.year, course_date.month, course_date.day, 
                        int(start_hm[0]), int(start_hm[1])
                    ))
                    end_dt = tz.localize(datetime(
                        course_date.year, course_date.month, course_date.day, 
                        int(end_hm[0]), int(end_hm[1])
                    ))
                    
                    # 清理課程名稱，移除 HTML 標籤和實體字符
                    summary_text = cell['course_text'].replace('<br>', ' ').replace('<br/>', ' ').strip()
                    summary_text = summary_text.replace('&nbsp;', '').replace('&nbsp', '')
                    summary_text = re.sub(r'\s+', ' ', summary_text).strip()

                    if not summary_text: 
                        continue
                    
                    # 生成唯一 UID (包含日期和課程資訊)
                    import uuid
                    uid = f"{course_date.strftime('%Y%m%d')}-{slot_idx}-{day_idx}-{uuid.uuid4()}"
                    
                    # 格式化時間為 UTC
                    dtstart = begin_dt.astimezone(pytz.UTC).strftime('%Y%m%dT%H%M%SZ')
                    dtend = end_dt.astimezone(pytz.UTC).strftime('%Y%m%dT%H%M%SZ')
                    
                    # 添加事件
                    ics_lines.extend([
                        "BEGIN:VEVENT",
                        fold_line(f"DTSTART:{dtstart}"),
                        fold_line(f"DTEND:{dtend}"),
                        fold_line(f"SUMMARY:{summary_text}"),
                        fold_line(f"UID:{uid}"),
                        "END:VEVENT"
                    ])
        
        # 移動到下一週
        current_week_start += timedelta(weeks=1)
        week_count += 1
    
    ics_lines.append("END:VCALENDAR")
    ics_content = '\r\n'.join(ics_lines)
    return Response(
        ics_content,
        mimetype="text/calendar",
        headers={"Content-disposition": "attachment; filename=course_schedule.ics"}
    )


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host='0.0.0.0', port=port)