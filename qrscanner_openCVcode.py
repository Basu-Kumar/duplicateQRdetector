import cv2
from pyzbar.pyzbar import decode
import pandas as pd
from datetime import datetime
import numpy as np
import time

class MultiQRCodeScanner:
    def __init__(self, camera_index=0, excel_path="multi_qr_codes.xlsx"):
        self.camera_index = camera_index
        self.excel_path = excel_path
        self.captured_codes = set()  # To avoid duplicates
        self.qr_data = []
        self.cap = None
        self.frame_count = 0
        self.last_save_time = datetime.now()
        
    def initialize_camera(self):
        """Initialize the camera capture"""
        # Use GSTREAMER pipeline for Pi camera
        gstreamer_pipeline = (
            "libcamerasrc ! "
            "video/x-raw, width=1280, height=720, framerate=30/1 ! "
            "videoconvert ! "
            "appsink"
        )
        self.cap = cv2.VideoCapture(gstreamer_pipeline, cv2.CAP_GSTREAMER)
        
        if not self.cap.isOpened():
            # Fallback to legacy Pi camera interface if GSTREAMER fails
            self.cap = cv2.VideoCapture(0)
            if not self.cap.isOpened():
                raise Exception("Could not open Pi camera. Please check if camera is properly connected and enabled.")
        
        # Give camera sensor some time to warm up
        time.sleep(2)
            
    def draw_qr_info(self, frame, decoded_objects):
        """Draw QR code information on frame"""
        for idx, obj in enumerate(decoded_objects):
            # Get the QR code points
            points = obj.polygon
            if points:
                # Draw the QR code boundary
                pts = np.array([(p.x, p.y) for p in points], np.int32)
                pts = pts.reshape((-1, 1, 2))
                cv2.polylines(frame, [pts], True, (0, 255, 0), 2)
                
                # Add a label with QR code content
                x = points[0].x
                y = points[0].y
                try:
                    data = obj.data.decode('utf-8')
                    truncated_data = data[:20] + "..." if len(data) > 20 else data
                    cv2.putText(frame, f"QR {idx+1}: {truncated_data}", 
                              (x, y-10), cv2.FONT_HERSHEY_SIMPLEX, 
                              0.5, (0, 255, 0), 2)
                except:
                    pass
                
        # Add counter for total QR codes in frame
        cv2.putText(frame, f"QR Codes in frame: {len(decoded_objects)}", 
                   (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 
                   1, (0, 255, 0), 2)
        
        return frame
        
    def process_frame(self, frame):
        """Process a single frame and detect QR codes"""
        self.frame_count += 1
        decoded_objects = decode(frame)
        new_codes_in_frame = 0
        
        for obj in decoded_objects:
            try:
                data = obj.data.decode('utf-8')
                
                # Only process if we haven't seen this QR code before
                if data not in self.captured_codes:
                    new_codes_in_frame += 1
                    self.captured_codes.add(data)
                    
                    self.qr_data.append({
                        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'frame_number': self.frame_count,
                        'qr_content': data,
                        'qr_type': obj.type,
                        'status': 'success'
                    })
                    
                    print(f"New QR Code detected: {data}")
                    
            except Exception as e:
                print(f"Error processing QR code: {str(e)}")
        
        # Draw information on frame
        frame = self.draw_qr_info(frame, decoded_objects)
        
        # Add text showing total unique codes captured
        cv2.putText(frame, f"Total unique codes captured: {len(self.captured_codes)}", 
                   (10, 60), cv2.FONT_HERSHEY_SIMPLEX, 
                   1, (0, 255, 0), 2)
        
        # Save to Excel if new codes were detected
        if new_codes_in_frame > 0:
            self.save_to_excel()
            
        return frame
    
    def save_to_excel(self):
        """Save the captured QR code data to Excel"""
        try:
            df = pd.DataFrame(self.qr_data)
            df.to_excel(self.excel_path, index=False)
            print(f"Excel file updated with {len(self.qr_data)} total records")
        except Exception as e:
            print(f"Error saving to Excel: {str(e)}")
    
    def start_scanning(self):
        """Start the scanning process"""
        try:
            self.initialize_camera()
            print("Starting multi QR code scanning...")
            print("Press 'q' to quit")
            
            while True:
                ret, frame = self.cap.read()
                if not ret:
                    print("Failed to grab frame")
                    break
                
                # Process the frame
                processed_frame = self.process_frame(frame)
                
                # Display the frame
                cv2.imshow('Multi QR Code Scanner', processed_frame)
                
                # Break the loop if 'q' is pressed
                if cv2.waitKey(1) & 0xFF == ord('q'):
                    break
                
        except Exception as e:
            print(f"Error during scanning: {str(e)}")
        
        finally:
            if self.cap is not None:
                self.cap.release()
            cv2.destroyAllWindows()
            self.save_to_excel()  # Final save

def main():
    scanner = MultiQRCodeScanner(
        camera_index=0,  # Change this if you have multiple cameras
        excel_path="multi_qr_codes.xlsx"
    )
    scanner.start_scanning()

if __name__ == "__main__":
    main()