import pyautogui
import time
import random

# Install pyautogui if not already: pip install pyautogui
# Note: This enhanced script mimics natural user mouse movements more realistically. 
# It uses random directions, distances, speeds, and intervals to avoid patterns.
# Run with caution, as it controls your mouse cursor.
# Press Ctrl+C to stop. Failsafe: Move mouse to top-left corner to emergency stop.

# Enable failsafe
pyautogui.FAILSAFE = True

# Optional: Get screen size to stay within bounds
screen_width, screen_height = pyautogui.size()


def random_mouse_movement():
    # Get current position
    x, y = pyautogui.position()

    # Random delta: small natural movements (e.g., 5-50 pixels)
    dx = random.randint(-50, 50)
    dy = random.randint(-50, 50)

    # Ensure we don't go off-screen (with some margin)
    new_x = max(0, min(screen_width - 1, x + dx))
    new_y = max(0, min(screen_height - 1, y + dy))

    # Random duration for movement (0.1 to 1 second, mimicking human speed)
    duration = random.uniform(0.1, 1.0)

    # Move to new position
    pyautogui.moveTo(new_x, new_y, duration=duration)

    # Optional: Occasionally simulate a small scroll (uncomment if needed)
    # if random.random() < 0.1:  # 10% chance
    #     pyautogui.scroll(random.choice([-1, 1]))  # Scroll up or down by 1 unit


try:
    while True:
        random_mouse_movement()

        # Random pause between movements (1-10 seconds, like idle browsing)
        sleep_time = random.uniform(1, 10)
        time.sleep(sleep_time)

except KeyboardInterrupt:
    print("Stopped mimicking user activity.")
