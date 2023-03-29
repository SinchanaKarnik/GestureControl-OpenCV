import cv2
import time
import math
import numpy as np
import HandTrackingModule as htm
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

wCam, hCam = 680, 480  # default

cap = cv2.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)
pTime = 0

detector = htm.HandDetector()
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_
    , CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
minVol = volume.GetVolumeRange()[0]
maxVol = volume.GetVolumeRange()[1]
volBar = 300
volPer = 0
while True:
    success, img = cap.read()
    img = detector.findHands(img)
    landmarks = detector.findPosition(img, draw=False)
    if len(landmarks) != 0:
        # print(landmarks[4], landmarks[8])
        x1, y1 = landmarks[4][1], landmarks[4][2]
        x2, y2 = landmarks[8][1], landmarks[8][2]
        cx, cy = (x1 + x2) // 2, (y1 + y2) // 2

        cv2.circle(img, (x1, y1), 10, (255, 0, 0), cv2.FILLED)
        cv2.circle(img, (x2, y2), 10, (255, 0, 0), cv2.FILLED)
        cv2.line(img, (x1, y1), (x2, y2), (255, 0, 0), 3)
        cv2.circle(img, (cx, cy), 10, (255, 0, 0), cv2.FILLED)

        length = math.hypot(x2 - x1, y2 - y1)
        print(length)

        # hand range 30-300, Vol range -65 - 0
        vol = np.interp(length, [50, 300], [minVol, maxVol])
        volBar = np.interp(length, [50, 300], [300, 158])
        volPer = np.interp(length, [50, 300], [0, 100])
        # print(vol)
        volume.SetMasterVolumeLevel(vol, None)

        if length < 50:
            cv2.circle(img, (cx, cy), 10, (0, 255, 0), cv2.FILLED)

    cv2.rectangle(img, (50, 158), (85, 300), (0, 255, 0), 3)
    cv2.rectangle(img, (50, int(volBar)), (85, 300), (0, 255, 0), cv2.FILLED)
    cv2.putText(img, "Vol %: " + str(int(volPer)), (40, 350), cv2.FONT_HERSHEY_PLAIN, 2, (255, 0, 0), 2)
    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime
    cv2.putText(img, str(int(fps)) + " FPS", (5, 50), cv2.FONT_HERSHEY_PLAIN, 2, (255, 0, 0), 2)
    cv2.imshow('Img', img)
    cv2.waitKey(1)
