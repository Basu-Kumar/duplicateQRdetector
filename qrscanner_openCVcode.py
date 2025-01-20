import cv2
from pyzbar.pyzbar import decode
import pandas as pd
from datetime import datetime
import numpy as np
import RPi.GPIO as GPIO
import time
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# GPIO Buzzer setup
BUZZER_PIN = 18  # GPIO18

class MultiQRCodeScanner:
    def __init__(self, camera_index=0, excel_path="multi_qr_codes.xlsx"):
        self.camera_index = camera_index
        self.excel_path = excel_path
        self.captured_codes = set()  # To avoid duplicates
        self.qr_data = []
        self.cap = None
        self.frame_count = 0
        self.last_save_time = datetime.now()
        
        # Initialize GPIO for buzzer
        GPIO.setmode(GPIO.BCM)
        GPIO.setup(BUZZER_PIN, GPIO.OUT)
        GPIO.output(BUZZER_PIN, GPIO.LOW)
        
    def buzz_alert(self, duration=0.2):
        """Activate buzzer for duplicate detection"""
        try:
            GPIO.output(BUZZER_PIN, GPIO.HIGH)
            time.sleep(duration)
            GPIO.output(BUZZER_PIN, GPIO.LOW)
        except Exception as e:
            print(f"Buzzer error: {str(e)}")
        
    def process_frame(self, frame):
        """Process a single frame and detect QR codes"""
        self.frame_count += 1
        decoded_objects = decode(frame)
        new_codes_in_frame = 0
        
        for obj in decoded_objects:
            try:
                data = obj.data.decode('utf-8')
                is_duplicate = data in self.captured_codes
                
                # Always record the entry, whether it's new or duplicate
                self.qr_data.append({
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'frame_number': self.frame_count,
                    'qr_content': data,
                    'qr_type': obj.type,
                    'status': 'duplicate' if is_duplicate else 'new'
                })
                
                # Handle duplicate or new code
                if is_duplicate:
                    print(f"Duplicate QR Code detected: {data}")
                    self.buzz_alert()  # Activate buzzer for duplicate
                else:
                    new_codes_in_frame += 1
                    self.captured_codes.add(data)
                    print(f"New QR Code detected: {data}")
                    
            except Exception as e:
                print(f"Error processing QR code: {str(e)}")
        
        # Draw information on frame
        frame = self.draw_qr_info(frame, decoded_objects)
        
        # Add text showing total unique codes captured
        cv2.putText(frame, f"Total unique codes captured: {len(self.captured_codes)}", 
                   (10, 60), cv2.FONT_HERSHEY_SIMPLEX, 
                   1, (0, 255, 0), 2)
        
        # Save to Excel if any codes were detected
        if len(decoded_objects) > 0:
            self.save_to_excel()
            
        return frame

    def save_to_excel(self):
        """Save QR code data to Excel with red highlighting for duplicates"""
        try:
            # First save using pandas
            df = pd.DataFrame(self.qr_data)
            df.to_excel(self.excel_path, index=False)
            
            # Then apply formatting with openpyxl
            workbook = load_workbook(self.excel_path)
            worksheet = workbook.active
            
            # Define red fill for duplicates
            red_fill = PatternFill(start_color='FFFF0000',
                                 end_color='FFFF0000',
                                 fill_type='solid')
            
            # Apply red background to duplicate entries (starting from row 2 to skip header)
            for row_idx, row in enumerate(worksheet.iter_rows(min_row=2), start=2):
                status = row[4].value  # Assuming 'status' is the 5th column
                if status == 'duplicate':
                    for cell in row:
                        cell.fill = red_fill
            
            # Save the formatted workbook
            workbook.save(self.excel_path)
            print(f"Excel file updated with {len(self.qr_data)} total records")
            
        except Exception as e:
            print(f"Error saving to Excel: {str(e)}")

    def draw_qr_info(self, frame, decoded_objects):
        """Draw QR code information on frame"""
        for idx, obj in enumerate(decoded_objects):
            points = obj.polygon
            if points:
                pts = np.array([(p.x, p.y) for p in points], np.int32)
                pts = pts.reshape((-1, 1, 2))
                
                # Draw in red for duplicates, green for new codes
                try:
                    data = obj.data.decode('utf-8')
                    color = (0, 0, 255) if data in self.captured_codes else (0, 255, 0)
                    cv2.polylines(frame, [pts], True, color, 2)
                    
                    x = points[0].x
                    y = points[0].y
                    truncated_data = data[:20] + "..." if len(data) > 20 else data
                    cv2.putText(frame, f"QR {idx+1}: {truncated_data}", 
                              (x, y-10), cv2.FONT_HERSHEY_SIMPLEX, 
                              0.5, color, 2)
                except:
                    pass
                
        return frame
    
    def cleanup(self):
        """Cleanup GPIO and other resources"""
        if self.cap is not None:
            self.cap.release()
        cv2.destroyAllWindows()
        GPIO.cleanup()  # Clean up GPIO on exit
        self.save_to_excel()

    def start_scanning(self):
        """Start the scanning process"""
        try:
            self.cap = cv2.VideoCapture(self.camera_index)
            if not self.cap.isOpened():
                raise Exception("Could not open camera")
                
            print("Starting QR code scanning... Press 'q' to quit")
            
            while True:
                ret, frame = self.cap.read()
                if not ret:
                    print("Failed to grab frame")
                    break
                
                processed_frame = self.process_frame(frame)
                cv2.imshow('QR Code Scanner', processed_frame)
                
                if cv2.waitKey(1) & 0xFF == ord('q'):
                    break
                
        except Exception as e:
            print(f"Error during scanning: {str(e)}")
        
        finally:
            self.cleanup()

def main():
    scanner = MultiQRCodeScanner(
        camera_index=0,
        excel_path="multi_qr_codes.xlsx"
    )
    scanner.start_scanning()

if __name__ == "__main__":
    main()