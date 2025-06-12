
# OMR Marks Automation using Python 📝📷

This project automates the evaluation of OMR (Optical Mark Recognition) sheets using Python and computer vision. It uses live video feed from an Android smartphone camera via the IP Webcam app to detect marked answers, calculates the score in real time, and overlays the results on the original sheet. The data is also saved into an Excel file for future reference.

---

## 📌 Features

- Live OMR Sheet Detection using phone camera
- Automated answer detection using image processing
- Marks calculation with validation (multiple markings handled)
- Real-time overlay of result and answers
- Name/Roll number recognition using Tesseract OCR
- Result storage in Excel file (`scores.xlsx`)
- Modular utility functions in `utils.py`

---

## 🧠 Technologies Used

- Python
- OpenCV
- NumPy
- Pandas
- Tesseract OCR
- IP Webcam Android App

---

## 📸 System Workflow

1. Capture OMR sheet using IP Webcam app.
2. Preprocess the frame (grayscale, blur, edge detection).
3. Detect contours to extract OMR sheet and grade area.
4. Warp perspective to flatten OMR region.
5. Threshold the sheet and detect filled bubbles.
6. Use Tesseract OCR to extract name/roll.
7. Compare answers and calculate score.
8. Show correct/wrong answers on sheet.
9. Store data in Excel if not already present.

---

## 🔧 Installation & Setup

1. **Clone the Repository**
   ```bash
   git clone https://github.com/your-username/omr-marks-automation.git
   cd omr-marks-automation
   ```

2. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Install Tesseract OCR**
   - Download from: https://github.com/tesseract-ocr/tesseract  
   - Update the Tesseract path in your `main.py`:
     ```python
     pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
     ```

4. **Run the Application**
   ```bash
   python main.py
   ```

5. **Set Up IP Webcam**
   - Install [IP Webcam](https://play.google.com/store/apps/details?id=com.pas.webcam) app on your Android phone.
   - Start the server and get the video stream URL (e.g., `http://192.168.x.x:8080/video`).
   - Replace the placeholder in `main.py`:
     ```python
     address = "http://your-ip-address:8080/video"
     ```

---

## 📂 File Structure

```
.
├── main.py               # Main application logic
├── utils.py              # Helper functions for image processing
├── scores.xlsx           # Excel file to store name and score
├── FinalResult.jpg       # Sample generated result image
├── 1.jpg                 # Sample OMR sheet image
├── README.md             # Project documentation
```

---

## 🚀 Future Improvements

- Support for multiple OMR layouts
- GUI interface for ease of use
- Cloud storage integration for results
- MCQ key customization from GUI or file

---

## 👤 Author

**Dev Patel**

> Developed as part of a practical learning initiative using OpenCV and Tesseract OCR.

---

## 📃 License

This project is licensed under the MIT License.
