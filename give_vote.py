from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch

def speak(str1):
    speak = Dispatch('SAPI.Spvoice')
    speak.speak(str1)

video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
if not os.path.exists('data/'):
    os.makedirs('data/')

with open('data/names.pkl', 'rb') as f:
    LABELS = pickle.load(f)

with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

# Define the size of the GUI window

gui_width = 1400  
gui_height = 700  

# Load and resize the background image to match the GUI size
imgBackground = cv2.imread("background_img.png")
imgBackground = cv2.resize(imgBackground, (gui_width, gui_height))

COL_NAMES = ['NAME', 'VOTE', 'DATE', 'TIME']

def check_if_exists(value):
    try:
        with open("Votes.csv", "r") as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                if row and row[0] == value:
                    return True
    except FileNotFoundError:
        print("File not found or unable to open the csv file.")
    return False

while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)
    output = None  

    offset = 5 

    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
        exist = os.path.isfile("Votes.csv")
        
        # Adjust the coordinates to move the rectangle and text to the left
        cv2.rectangle(frame, (x - offset, y), (x + w - offset, y + h), (0, 0, 255), 1)
        cv2.rectangle(frame, (x - offset, y), (x + w - offset, y + h), (50, 50, 255), 2)
        cv2.rectangle(frame, (x - offset, y - 40), (x + w - offset, y), (50, 50, 255), -1)
        cv2.putText(frame, str(output[0]), (x - offset, y - 15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
        attendance = [output[0], timestamp]

    # Resize the captured frame to fit within the GUI
    frame_resized = cv2.resize(frame, (640, 480))

    # Calculate the position to place the resized frame on the background image
    x_offset = (gui_width - 640) // 2
    y_offset = (gui_height - 480) // 2
    imgBackground[y_offset:y_offset + 480, x_offset:x_offset + 640] = frame_resized

    cv2.imshow('frame', imgBackground)
    k = cv2.waitKey(1)

    if output is not None:  
        voter_exist = check_if_exists(output[0])
        if voter_exist:
            speak("You have already voted")
            break

        if k == ord('1'):
            speak("Your vote has been recorded")
            time.sleep(3)
            if exist:
                with open("Votes.csv", "a") as csvfile:
                    writer = csv.writer(csvfile)
                    attendance = [output[0], "United Congress Party", date, timestamp]
                    writer.writerow(attendance)
            else:
                with open("Votes.csv", "a") as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(COL_NAMES)
                    attendance = [output[0], "United Congress Party", date, timestamp]
                    writer.writerow(attendance)
            speak("Thank you for participating in the elections")
            break

        if k == ord('2'):
            speak("Your vote has been recorded")
            time.sleep(5)
            if exist:
                with open("Votes.csv", "a") as csvfile:
                    writer = csv.writer(csvfile)
                    attendance = [output[0], "United Republican Front", date, timestamp]
                    writer.writerow(attendance)
            else:
                with open("Votes.csv", "a") as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(COL_NAMES)
                    attendance = [output[0], "United Republican Front", date, timestamp]
                    writer.writerow(attendance)
            speak("Thank you for participating in the elections")
            break

        if k == ord('3'):
            speak("Your vote has been recorded")
            time.sleep(5)
            if exist:
                with open("Votes.csv", "a") as csvfile:
                    writer = csv.writer(csvfile)
                    attendance = [output[0], "United Left Front", date, timestamp]
                    writer.writerow(attendance)
            else:
                with open("Votes.csv", "a") as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(COL_NAMES)
                    attendance = [output[0], "United Left Front", date, timestamp]
                    writer.writerow(attendance)
            speak("Thank you for participating in the elections")
            break

        if k == ord('4'):
            speak("Your vote has been recorded")
            time.sleep(5)
            if exist:
                with open("Votes.csv", "a") as csvfile:
                    writer = csv.writer(csvfile)
                    attendance = [output[0], "New Independent Party", date, timestamp]
                    writer.writerow(attendance)
            else:
                with open("Votes.csv", "a") as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(COL_NAMES)
                    attendance = [output[0], "New Independent Party", date, timestamp]
                    writer.writerow(attendance)
            speak("Thank you for participating in the elections")
            break

video.release()
cv2.destroyAllWindows()
