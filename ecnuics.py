import re
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
import os
import pytz
from icalendar import Calendar, Event
import qrcode
import socket
import threading
import http.server
import socketserver

def read_html_file(file_path):
    """读取HTML文件"""
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()

def parse_course_table(html_content):
    """解析HTML得到课表数据"""
    soup = BeautifulSoup(html_content, 'html.parser')
    
    course_table = soup.find('table', id='manualArrangeCourseTable')
    if not course_table:
        raise Exception("未找到课程表格")
    
    courses = []
    cells = course_table.find_all('td', style=lambda s: s and ('background-color: rgb(148, 174, 243)' in s or 'backGround-Color:rgb(148, 174, 243)' in s))
    
    for cell in cells:
        try:
            course_text = cell.get_text().strip()
            if not course_text:
                continue
            
            if cell.has_attr('title'):
                course_text = cell.get('title')
            
            pattern = r'(.*?)\s+(.*?)(?:\((.*?)\))?(?:;)?\s*\((.*?)\)'
            match = re.search(pattern, course_text)
            if not match:
                print(f"无法解析课程信息: {course_text}")
                continue
            
            teacher = match.group(1).strip() if match.group(1) else ""
            course_name = match.group(2).strip() if match.group(2) else ""
            course_id = match.group(3) if match.group(3) else ""
            location_info = match.group(4).strip() if match.group(4) else ""
            
            week_pattern = r'(\d+)-(\d+)([单双])?'
            week_match = re.search(week_pattern, location_info)
            if not week_match:
                print(f"无法解析周次信息: {location_info}")
                continue
            
            start_week = int(week_match.group(1))
            end_week = int(week_match.group(2))
            
            week_type = "all"
            if "单" in location_info:
                week_type = "odd"
            elif "双" in location_info:
                week_type = "even"
            
            room_match = re.search(r'(?:,|，)(.*?)(?:,|，|$)', location_info)
            room = room_match.group(1) if room_match else ""
            if "【" in room:
                room = room.split('【')[0].strip()
            
            cell_id = cell.get('id', '')
            if not cell_id.startswith('TD'):
                print(f"单元格ID格式不符合预期: {cell_id}")
                continue
                
            position_match = re.match(r'TD(\d+)_\d+', cell_id)
            if not position_match:
                print(f"无法从单元格ID解析位置: {cell_id}")
                continue
                
            position = int(position_match.group(1))
            day_of_week = position // 14 + 1
            section_base = position % 14
            start_section = section_base + 1
            
            rowspan = int(cell.get('rowspan', 1))
            end_section = start_section + rowspan - 1
            
            course = {
                'name': course_name,
                'teacher': teacher,
                'course_id': course_id,
                'start_week': start_week,
                'end_week': end_week,
                'week_type': week_type,
                'day_of_week': day_of_week,
                'start_section': start_section,
                'end_section': end_section,
                'room': room
            }
            courses.append(course)
            print(f"解析到课程: {course_name}, 星期{day_of_week}, 第{start_section}-{end_section}节, {start_week}-{end_week}周{week_type}, 教室: {room}")
            
        except Exception as e:
            print(f"解析课程单元格时出错: {e}")
            continue
    
    return courses

def get_semester_start_date():
    """获取学期开始日期
    
    默认设为2025年2月17日，即2024-2025学年第二学期开始日期
    格式为 年, 月, 日
    """
    return datetime(2025, 2, 17)

def calculate_course_time(course, semester_start):
    """计算课程开始和结束时间"""
    time_slots = {
        1: ("08:00", "08:45"),
        2: ("08:50", "09:35"),
        3: ("09:50", "10:35"),
        4: ("10:40", "11:25"),
        5: ("11:30", "12:15"),
        6: ("13:00", "13:45"),
        7: ("13:50", "14:35"),
        8: ("14:50", "15:35"),
        9: ("15:40", "16:25"),
        10: ("16:30", "17:15"),
        11: ("18:00", "18:45"),
        12: ("18:50", "19:35"),
        13: ("19:40", "20:25"),
        14: ("20:30", "21:15")
    }
    
    day_of_week = course['day_of_week']
    start_week = course['start_week']
    end_week = course['end_week']
    week_type = course['week_type']
    
    events = []
    
    days_to_add = (day_of_week - 1)
    first_day = semester_start + timedelta(days=days_to_add)
    
    for week in range(start_week, end_week + 1):
        if week_type == "odd" and week % 2 == 0:
            continue
        if week_type == "even" and week % 2 == 1:
            continue
        
        current_date = first_day + timedelta(weeks=week-1)
        
        start_time_str, _ = time_slots[course['start_section']]
        _, end_time_str = time_slots[course['end_section']]
        
        start_hour, start_minute = map(int, start_time_str.split(':'))
        end_hour, end_minute = map(int, end_time_str.split(':'))
        
        start_datetime = current_date.replace(hour=start_hour, minute=start_minute)
        end_datetime = current_date.replace(hour=end_hour, minute=end_minute)
        
        event = {
            'summary': course['name'],
            'location': course['room'],
            'description': f"教师: {course['teacher']}\n课程ID: {course['course_id']}\n地点: {course['room']}",
            'dtstart': start_datetime,
            'dtend': end_datetime
        }
        events.append(event)
    
    return events

def generate_ics(courses, semester_start, output_file):
    """生成ICS文件"""
    cal = Calendar()
    
    cal.add('prodid', '-//ECNU Course Schedule//CN')
    cal.add('version', '2.0')
    cal.add('calscale', 'GREGORIAN')
    cal.add('method', 'PUBLISH')
    cal.add('X-WR-CALNAME', 'ECNU课表')
    cal.add('X-WR-TIMEZONE', 'Asia/Shanghai')
    
    tz = pytz.timezone('Asia/Shanghai')
    
    for course in courses:
        events = calculate_course_time(course, semester_start)
        
        for event_data in events:
            event = Event()
            event.add('summary', event_data['summary'])
            event.add('location', event_data['location'])
            event.add('description', event_data['description'])
            
            event.add('dtstart', event_data['dtstart'].replace(tzinfo=tz))
            event.add('dtend', event_data['dtend'].replace(tzinfo=tz))
            
            event.add('dtstamp', datetime.now(tz))
            event.add('uid', f"{event_data['dtstart'].strftime('%Y%m%dT%H%M%S')}-{course['name']}@ecnu.edu.cn")
            
            cal.add_component(event)
    
    with open(output_file, 'wb') as f:
        f.write(cal.to_ical())
    
    print(f"成功生成ICS文件: {output_file}")

def get_local_ip():
    """获取本机在局域网中的IP地址"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def start_http_server(directory, port=8000):
    """启动一个简单的HTTP服务器"""
    handler = http.server.SimpleHTTPRequestHandler
    
    class CustomHTTPServer(socketserver.TCPServer):
        allow_reuse_address = True
    
    os.chdir(directory)
    httpd = CustomHTTPServer(("", port), handler)
    
    print(f"服务器启动在端口 {port}")
    
    server_thread = threading.Thread(target=httpd.serve_forever)
    server_thread.daemon = True
    server_thread.start()
    
    return httpd

def generate_qrcode(url, output_file='qrcode.png'):
    """生成包含URL的二维码"""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(url)
    qr.make(fit=True)
    
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(output_file)
    print(f"二维码已保存到: {output_file}")
    
    try:
        from PIL import Image
        Image.open(output_file).show()
    except Exception:
        pass
    
def main():
    """主函数，处理输入参数并执行课表转换流程"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    html_file = input("请输入HTML文件路径 (默认为当前目录下的courseTable.html): ").strip() or "courseTableForStd!courseTable.action.html"
    
    if not os.path.isabs(html_file):
        html_file = os.path.join(script_dir, html_file)
    
    if not os.path.exists(html_file):
        print(f"错误: 文件不存在: {html_file}")
        return
    
    date_input = input("请输入学期开始日期 (格式: YYYY-MM-DD，默认为2025-02-17): ").strip() or "2025-02-17"
    try:
        semester_start = datetime.strptime(date_input, "%Y-%m-%d")
    except ValueError:
        print("日期格式错误，使用默认日期：2025-02-17")
        semester_start = datetime(2025, 2, 17)
    
    output_file = input("请输入输出文件名称 (默认为 ecnu_course.ics): ").strip() or "ecnu_course.ics"
    if not os.path.isabs(output_file):
        output_file = os.path.join(script_dir, output_file)
    
    try:
        html_content = read_html_file(html_file)
        courses = parse_course_table(html_content)
        
        if not courses:
            print("未找到课程信息")
            return
        
        print(f"共解析到 {len(courses)} 门课程:")
        for i, course in enumerate(courses, 1):
            print(f"{i}. {course['name']} - {course['teacher']} - 星期{course['day_of_week']} 第{course['start_section']}-{course['end_section']}节 - {course['room']}")
        
        generate_ics(courses, semester_start, output_file)
        print(f"课表已成功导出到 {output_file}")
        
        share_option = input("是否创建二维码以便手机扫码下载? (y/n): ").strip().lower()
        if share_option == 'y':
            file_dir = os.path.dirname(os.path.abspath(output_file))
            file_name = os.path.basename(output_file)
            
            port = 8000
            server = start_http_server(file_dir, port)
            
            local_ip = get_local_ip()
            
            url = f"http://{local_ip}:{port}/{file_name}"
            qrcode_file = os.path.join(file_dir, "ecnu_course_qrcode.png")
            generate_qrcode(url, qrcode_file)
            
            print(f"\n请用手机扫描二维码下载课表，或访问以下地址:")
            print(f"{url}")
            print("\n按Enter键退出服务器...")
            input()
            server.shutdown()
            
    except Exception as e:
        print(f"发生错误: {e}")

if __name__ == "__main__":
    main()