import cv2
from ultralytics import YOLO
import os
import time

# Loo YOLO mudel (näiteks valge YOLO v8 mudel)
model = YOLO("yolov8n.pt")  # või tee oma mudel

# Loo kaamera objekt
cap = cv2.VideoCapture(0)

# Loo kaust piltide salvestamiseks
save_dir = "salvestatud_pildid"
os.makedirs(save_dir, exist_ok=True)

# Arvuta juba olemasolevad pildid, et loendur jätkuks
existing_files = os.listdir(save_dir)
counter = len(existing_files) + 1

print("Kaamera töötab. Vajuta ENTER, et teha pilt ja tuvastada objektid. Q sulgeb programmi.")

while True:
    ret, frame = cap.read()
    if not ret:
        print("Kaamerast pilti ei saanud!")
        break

    cv2.imshow("Kaamera", frame)
    key = cv2.waitKey(1) & 0xFF

    if key == ord("\r") or key == 13:  # ENTER-klahv
        # Salvestamiseks tee koopia kaadrist
        image_to_save = frame.copy()

        # Ennusta objektid
        results = model(image_to_save)

        # Tõmba tulemused pildile (kasti ja nimega)
        annotated_frame = results[0].plot()  # plot() tagastab numpy array

        # Loo failinimi järjekorranumbriga
        filename = os.path.join(save_dir, f"pilt_{counter}.jpg")
        cv2.imwrite(filename, annotated_frame)
        print(f"Pilt salvestatud: {filename}")

        counter += 1  # suurenda loendurit järgmise pildi jaoks

        # Näita salvestatud pilti mõneks sekundiks
        cv2.imshow("Tuvastatud objektid", annotated_frame)
        cv2.waitKey(2000)  # 2 sekundit
        cv2.destroyWindow("Tuvastatud objektid")  # Sulge see aken

    elif key == ord("q") or key == ord("Q"):
        print("Programm lõpetatakse.")
        break

cap.release()
cv2.destroyAllWindows()