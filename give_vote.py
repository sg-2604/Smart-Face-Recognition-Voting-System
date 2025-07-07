from sklearn.neighbors import KNeighborsClassifier
import numpy as np
import pickle
import os
import cv2
import time
import csv
from datetime import datetime
from win32com.client import Dispatch

def speak(str):
    speak=Dispatch(("SAPI.SpVoice"))
    speak.Speak(str)
video=cv2.VideoCapture(0)
facedetect=cv2.CascadeClassifier(cv2.data.haarcascades+'haarcascade_frontalface_default.xml')
if not os.path.exists('data/'):
    os.makedirs('data/')

with open('data/names.pkl','rb') as f:
    LABELS=pickle.load(f)
with open('data/faces_data.pkl','rb') as f:
    FACES=pickle.load(f)

model=KNeighborsClassifier(n_neighbors=5)
model.fit(FACES,LABELS)
img_path=r"C:\Users\Sankalp Gupta\Projects\smart face recognition voting system\background.jpeg"
imgBackground=cv2.imread(img_path)
if imgBackground is None:
    print("Image not found")
COL_NAMES=['NAME','VOTE','DATE','TIME']

while True:
    ret,frame=video.read()
# if not ret or frame is None:
#     print("Error: Couldn't capture frame from webcam.")
#     # continue
    
# print("Background Image Shape:", imgBackground.shape)
# print("Webcam Frame Shape:", frame.shape)
    gray=cv2.cvtColor(frame,cv2.COLOR_BGR2GRAY)
    faces=facedetect.detectMultiScale(gray,1.3,5)
    predicted=None
    for (x,y,w,h) in faces:
        crop_img=frame[y:y+h,x:x+w]
        resized_img=cv2.resize(crop_img,(50,50)).flatten().reshape(1,-1)
        predicted=model.predict(resized_img)
        ts=time.time()
        date=datetime.fromtimestamp(ts).strftime('%d-%m-%Y')
        timestamp=datetime.fromtimestamp(ts).strftime('%H:%M:%S')
        exist=os.path.isfile("Votes"+".csv")
        cv2.rectangle(frame,(x,y),(x+w,y+h),(0,0,255),1)
        cv2.rectangle(frame,(x,y),(x+w,y+h),(50,50,255),2)
        cv2.rectangle(frame,(x,y-40),(x+w,y),(50,50,255),-1)
        cv2.putText(frame,str(predicted[0]),(x,y-15),cv2.FONT_HERSHEY_SIMPLEX,1,(255,255,255),1)
        cv2.rectangle(frame,(x,y),(x+w,y+h),(50,50,255),1)
        attendance=[predicted[0],timestamp]
    imgBackground[200:200+480,100:100+640]=frame


    cv2.imshow('frame',imgBackground)
    k=cv2.waitKey(1)

    def check_if_exists(value):
        try:
            with open('Votes.csv', 'r') as file:
                reader = csv.reader(file)
                for row in reader:
                    if row and value == row[0]:
                        return True
        except FileNotFoundError:
            print("File not found or unable to open CSV file.")
        return False
    if predicted is not None:
        voter_exists = check_if_exists(predicted[0])

        if voter_exists:
            speak("You have already voted")
            print("You have already voted")
            break
        else:
            if k==ord('1'):
                speak("Your vote has been recorded")
                print("Your vote has been recorded")
                time.sleep(3)
                if exist:
                    with open('Votes.csv','+a',newline='') as f:
                        writer=csv.writer(f)
                        attendance=[predicted[0],'BJP',date,timestamp]
                        writer.writerow(attendance)
                    f.close()
                else:
                    with open('Votes.csv','w',newline='') as f:
                        writer=csv.writer(f)
                        writer.writerow(COL_NAMES)
                        attendance=[predicted[0],'BJP',date,timestamp]
                        writer.writerow(attendance)
                    f.close()
                speak("Thank you for voting")
                print("Thank you for voting")
                break
            elif k==ord('2'):
                speak("Your vote has been recorded")
                print("Your vote has been recorded")
                time.sleep(3)
                if exist:
                    with open('Votes.csv','+a') as f:
                        writer=csv.writer(f)
                        attendance=[predicted[0],'CONGRESS',date,timestamp]
                        writer.writerow(attendance)
                    f.close()
                else:
                    with open('Votes.csv','w',newline='') as f:
                        writer=csv.writer(f)
                        writer.writerow(COL_NAMES)
                        attendance=[predicted[0],'CONGRESS',date,timestamp]
                        writer.writerow(attendance)
                    f.close()
                speak("Thank you for voting")
                print("Thank you for voting")
                break
            elif k==ord('3'):
                speak("Your vote has been recorded")
                print("Your vote has been recorded")
                time.sleep(3)
                if exist:
                    with open('Votes.csv','+a') as f:
                        writer=csv.writer(f)
                        attendance=[predicted[0],'AAP',date,timestamp]
                        writer.writerow(attendance)
                    f.close()
                else:
                    with open('Votes.csv','w',newline='') as f:
                        writer=csv.writer(f)
                        writer.writerow(COL_NAMES)
                        attendance=[predicted[0],'AAP',date,timestamp]
                        writer.writerow(attendance)
                    f.close()
                speak("Thank you for voting")
                print("Thank you for voting")
                break
            elif k==ord('4'):
                speak("Your vote has been recorded")
                print("Your vote has been recorded")
                time.sleep(3)
                if exist:
                    with open('Votes.csv','+a') as f:
                        writer=csv.writer(f)
                        attendance=[predicted[0],'NOTA',date,timestamp]
                        writer.writerow(attendance)
                    f.close()
                else:
                    with open('Votes.csv','w',newline='') as f:
                        writer=csv.writer(f)
                        writer.writerow(COL_NAMES)
                        attendance=[predicted[0],'NOTA',date,timestamp]
                        writer.writerow(attendance)
                    f.close()
                speak("Thank you for voting")
                print("Thank you for voting")
                break        
video.release()
cv2.destroyAllWindows()

