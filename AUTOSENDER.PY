import pyautogui
import time

# Script shuru hone se pehle aapko 5 seconds ka time milega
# Taaki aap us chat box par click kar sakein jahan message bhejna hai
print("Aapke paas 5 seconds hain... Us chat par click karein jahan message bhejna hai!")
time.sleep(5)

# Jo message bhejna hai wo yahan likhein
message = "Hello! Ye ek automated message hai."

# Kitni baar bhejna hai (100 se zyada)
total_messages = 105

for i in range(total_messages):
    # Message type karega
    pyautogui.typewrite(f"{message} (#{i+1})") 
    
    # Send karne ke liye Enter press karega
    pyautogui.press('enter')
    
    # Do messages ke beech me thoda gap (0.5 second) taaki system hang na ho
    time.sleep(0.5)

print("Saare messages successfully bhej diye gaye hain!")