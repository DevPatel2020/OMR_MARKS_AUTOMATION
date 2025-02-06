import cv2
import numpy as np
import pandas as pd
import pytesseract
import utlis

# Set the Tesseract executable path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

path = "1.jpg"
widthImg = 700
heightImg = 700
questions = 5
choices = 5
ans = [0, 1, 4, 3, 2]
webcamFeed = True
cameraNo = 0
excel_file = 'scores.xlsx'


# Initialize the video capture
cap = cv2.VideoCapture(cameraNo)
address = "http://192.168.7.154:8080///video"
cap.open(address)
cap.set(10, 150)

# Load existing Excel file or create a new one
try:
    df = pd.read_excel(excel_file)
    print(f"Excel file {excel_file} found, loading existing data.")
except FileNotFoundError:
    df = pd.DataFrame(columns=['Name', 'Score'])
    print(f"Excel file {excel_file} not found, creating a new one.")

def get_name_from_image(image):
    # Crop the region where the name is written (adjust the coordinates as needed)
    name_region = image[50:150, 100:400]
    # Convert the region to grayscale
    gray_name = cv2.cvtColor(name_region, cv2.COLOR_BGR2GRAY)
    # Apply OCR to extract text
    name = pytesseract.image_to_string(gray_name, config='--psm 7')  # Adjust config as needed
    return name.strip()

while True:
    if webcamFeed:
        success, img = cap.read()
    else:
        img = cv2.imread(path)

    # Preprocessing
    img = cv2.resize(img, (widthImg, heightImg))
    imgContours = img.copy()
    imgFinal = img.copy()
    imgBiggestContours = img.copy()
    imgGray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    imgBlur = cv2.GaussianBlur(imgGray, (5, 5), 1)
    imgCanny = cv2.Canny(imgBlur, 10, 50)

    try:
        # Find All Contours
        contours, hierarchy = cv2.findContours(imgCanny, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
        cv2.drawContours(imgContours, contours, -1, (255, 73, 7), 10)

        # Find Rectangle
        rectCon = utlis.rectCountour(contours)
        biggestContour = utlis.getCornerPoints(rectCon[0])
        gradePoints = utlis.getCornerPoints(rectCon[1])

        if biggestContour.size != 0 and gradePoints.size != 0:
            cv2.drawContours(imgBiggestContours, biggestContour, -1, (0, 255, 0), 20)
            cv2.drawContours(imgBiggestContours, gradePoints, -1, (255, 0, 0), 20)

            biggestContour = utlis.reorder(biggestContour)
            gradePoints = utlis.reorder(gradePoints)

            pt1 = np.float32(biggestContour)
            pt2 = np.float32([[0, 0], [widthImg, 0], [0, heightImg], [widthImg, heightImg]])
            matrix = cv2.getPerspectiveTransform(pt1, pt2)
            imgWarpColored = cv2.warpPerspective(img, matrix, (widthImg, heightImg))

            ptG1 = np.float32(gradePoints)
            ptG2 = np.float32([[0, 0], [325, 0], [0, 150], [325, 150]])
            matrixG = cv2.getPerspectiveTransform(ptG1, ptG2)
            imgGrayedDisplay = cv2.warpPerspective(img, matrixG, (325, 150))

            # Apply Threshold
            imgWarpGray = cv2.cvtColor(imgWarpColored, cv2.COLOR_BGR2GRAY)
            imgThresh = cv2.threshold(imgWarpGray, 150, 255, cv2.THRESH_BINARY_INV)[1]

            boxes = utlis.splitBoxes(imgThresh)

            # Getting No Zero Pixel Value Of Each Box
            myPixelVal = np.zeros((questions, choices))
            countC = 0
            countR = 0
            for image in boxes:
                totalPixels = cv2.countNonZero(image)
                myPixelVal[countR][countC] = totalPixels
                countC += 1
                if countC == choices:
                    countR += 1
                    countC = 0

            # Finding Index Value Of Markings
            # Finding Index Value of Markings
            myIndex = []
            grading = []

            for x in range(questions):
                arr = myPixelVal[x]  # Get the pixel values for the current question
                maxVal = np.amax(arr)  # Find the maximum pixel value in the row
                markedChoices = np.where(arr >= maxVal * 0.8)[
                    0]  # Consider bubbles marked if they are close to the max value

                if len(markedChoices) > 1:  # If more than one bubble is marked
                    myIndex.append(-1)  # Use -1 to represent invalid answers (multiple marks)
                    grading.append(0)  # Mark it as 0 for grading
                else:
                    myIndex.append(int(markedChoices[0]))  # Append the single marked choice
                    grading.append(1 if ans[x] == markedChoices[0] else 0)  # Compare with the answer key

            print("Detected Answers:", myIndex)
            print("Grading:", grading)

            # Calculate final score
            score = (sum(grading) / questions) * 100
            print(f"Final Score: {score}%")

            # Displaying Answer
            imgResult = imgWarpColored.copy()
            imgResult = utlis.showAnswers(imgResult, myIndex, grading, ans, questions, choices)
            imgRawDrawing = np.zeros_like(imgWarpColored)
            imgRawDrawing = utlis.showAnswers(imgRawDrawing, myIndex, grading, ans, questions, choices)
            invMatrix = cv2.getPerspectiveTransform(pt2, pt1)
            imgInvWarp = cv2.warpPerspective(imgRawDrawing, invMatrix, (widthImg, heightImg))

            imgRawGrade = np.zeros_like(imgGrayedDisplay)
            cv2.putText(imgRawGrade, str(int(score)) + "%", (60, 100), cv2.FONT_HERSHEY_COMPLEX, 3, (0, 255, 255), 3)
            invMatrixG = cv2.getPerspectiveTransform(ptG2, ptG1)
            imgInvGradeDisplay = cv2.warpPerspective(imgRawGrade, invMatrixG, (widthImg, heightImg))

            imgFinal = cv2.addWeighted(imgFinal, 1, imgInvWarp, 1, 0)
            imgFinal = cv2.addWeighted(imgFinal, 1, imgInvGradeDisplay, 1, 0)

            # Extract name/roll number
            name = get_name_from_image(imgWarpColored)
            print(f"Detected Name/Roll Number: {name}")

            # Save results to Excel
            if df[(df['Name'] == name) & (df['Score'] == score)].empty:
                new_entry = pd.DataFrame({'Name': [name], 'Score': [score]})
                df = pd.concat([df, new_entry], ignore_index=True)
                df.to_excel(excel_file, index=False)
                print(f"New entry added and saved to {excel_file}")
            else:
                print(f"Entry for {name} with score {score} already exists. No new entry added.")

        imgBlank = np.zeros_like(img)
        imgArray = ([img, imgGray, imgBlur, imgCanny],
                    [imgContours, imgBiggestContours, imgWarpColored, imgThresh],
                    [imgResult, imgRawDrawing, imgInvWarp, imgFinal])
    except Exception as e:
        print(f"Error: {e}")
        imgBlank = np.zeros_like(img)
        imgArray = ([img, imgGray, imgBlur, imgCanny],
                    [imgBlank, imgBlank, imgBlank, imgBlank],
                    [imgBlank, imgBlank, imgBlank, imgBlank])
    labels = [["Original", "Gray", "Blur", "Canny"],
              ["Contours", "Biggest Con", "Wrap", "Threshold"],
              ["Result", "Raw Drawing", "Inv Warp", "Final"]]
    imgStacked = utlis.stackImages(imgArray, 0.4, labels)

    cv2.imshow("Final Result", imgFinal)
    cv2.imshow("Stacked Images", imgStacked)


    cv2.imshow("Final Result", imgFinal)
    cv2.imshow("Stacked Images", imgStacked)
    if cv2.waitKey(1) & 0xFF == ord('s'):
        cv2.imwrite("FinalResult.jpg", imgFinal)
        cv2.waitKey(300)
