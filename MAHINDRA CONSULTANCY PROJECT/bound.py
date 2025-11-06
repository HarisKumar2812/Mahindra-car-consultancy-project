import cv2
from datetime import datetime
from ultralytics import YOLO
import easyocr
import openpyxl
from openpyxl.styles import Font
import re
import os

# Load Model
model = YOLO(r"D:\intern\codsoftintern\Bus-Time-Management-5\best.pt")

# OCR Reader
reader = easyocr.Reader(['en'])

# Video Input
video = cv2.VideoCapture(r"D:\intern\codsoftintern\Bus-Time-Management-5\newrequirements\just.mp4")

# Video Properties
frame_width = int(video.get(cv2.CAP_PROP_FRAME_WIDTH))
frame_height = int(video.get(cv2.CAP_PROP_FRAME_HEIGHT))
fps = video.get(cv2.CAP_PROP_FPS)

# Save Output Video
output_video = cv2.VideoWriter("output.avi", cv2.VideoWriter_fourcc(*"MJPG"), fps, (frame_width, frame_height))

# ROI Variables
selecting_roi = False
roi_start = (0, 0)
roi_end = (0, 0)

plate_text_dict = {}
previous_plates = {}
excel_file = "plate_data.xlsx"

# Excel Setup
if os.path.exists(excel_file):
    workbook = openpyxl.load_workbook(excel_file)
    worksheet = workbook.active
    sno = worksheet.max_row
else:
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.append(["Sno.", "Number Plate", "Date", "Time"])
    for cell in worksheet["1:1"]:
        cell.font = Font(bold=True)
    sno = 1

def is_valid_license_plate(text):
    pattern = r'^(AP|AR|AS|BR|CG|GA|GJ|HR|HP|JH|KA|KL|MP|MH|MN|ML|MZ|NL|OD|PB|RJ|SK|TN|TS|TR|UP|UK|WB|AN|CH|DD|DL|JK|LA|LD|PY)\s?[0-9]{2}\s?[A-Z]{1,2}\s?[0-9]{1,4}$'
    return bool(re.match(pattern, text))

def clean_text(text):
    text = text.replace('.', '')
    corrected = []
    for i, char in enumerate(text):
        if char == 'O' and i in [2, 3, 6, 7]:
            corrected.append('0')
        else:
            corrected.append(char)
    return ''.join(corrected)

def draw_roi(event, x, y, flags, param):
    global selecting_roi, roi_start, roi_end
    if event == cv2.EVENT_LBUTTONDOWN:
        selecting_roi = True
        roi_start = (x, y)
    elif event == cv2.EVENT_MOUSEMOVE and selecting_roi:
        roi_end = (x, y)
    elif event == cv2.EVENT_LBUTTONUP:
        selecting_roi = False
        roi_end = (x, y)

cv2.namedWindow("Select ROI", cv2.WINDOW_NORMAL)
cv2.setMouseCallback("Select ROI", draw_roi)

while True:
    ret, frame = video.read()
    if not ret:
        break

    if selecting_roi:
        temp = frame.copy()
        cv2.rectangle(temp, roi_start, roi_end, (0, 255, 0), 2)
        cv2.imshow("Select ROI", temp)
    else:
        results = model.predict(frame, conf=0.5, classes=0)
        for r in results:
            for box in r.boxes.xyxy:
                x1, y1, x2, y2 = map(int, box.tolist())

                if x1 >= roi_start[0] and x2 <= roi_end[0] and y1 >= roi_start[1] and y2 <= roi_end[1]:
                    plate_img = frame[y1:y2, x1:x2]
                    plate_gray = cv2.cvtColor(plate_img, cv2.COLOR_BGR2GRAY)
                    ocr_result = reader.readtext(plate_gray)

                    if ocr_result:
                        current_time = datetime.now()
                        raw_text = ' '.join([t[1] for t in ocr_result])
                        plate = clean_text(raw_text)

                        if is_valid_license_plate(plate):
                            if plate not in previous_plates or (current_time - previous_plates[plate]).seconds >= 10:
                                previous_plates[plate] = current_time
                                plate_text_dict[current_time] = plate
                                worksheet.append([sno, plate, current_time.strftime("%Y-%m-%d"), current_time.strftime("%H:%M:%S")])
                                sno += 1

                            cv2.rectangle(frame, (x1, y1), (x2, y2), (0, 255, 0), 2)
                            cv2.putText(frame, plate, (x1, y1 - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0,255,0), 2)

    cv2.imshow("Frame", frame)
    output_video.write(frame)

    key = cv2.waitKey(1)
    if key == ord("q"):
        break

video.release()
output_video.release()
cv2.destroyAllWindows()
workbook.save(excel_file)

print("\nDetected Plates:")
for time, plate in plate_text_dict.items():
    print(f"{time}: {plate}")
