import cv2

# Test if OpenCV can open the default camera
cam = cv2.VideoCapture(0)

if cam.isOpened():
    print("Camera opened successfully.")
else:
    print("Failed to open camera.")
