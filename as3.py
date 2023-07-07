import pyautogui
import time
from PIL import ImageGrab
from moviepy.editor import *

def convert_ppt_to_video(ppt_path, video_path):
    # Open the PowerPoint presentation
    pyautogui.press('win')  # Open the Windows start menu
    time.sleep(1)  # Wait for the start menu to open
    pyautogui.write('PowerPoint')  # Search for PowerPoint
    time.sleep(1)  # Wait for search results to appear
    pyautogui.press('enter')  # Open PowerPoint
    time.sleep(2)  # Wait for PowerPoint to open

    # Open the PPT file
    pyautogui.press('ctrl+o')  # Open the file dialog
    time.sleep(1)  # Wait for the dialog to appear
    pyautogui.write(ppt_path)  # Enter the PPT file path
    pyautogui.press('enter')  # Press Enter to open the file
    time.sleep(2)  # Wait for the file to open

    # Set up the video recording
    app = pyautogui.getWindowsWithTitle('PowerPoint')[0]
    left, top, width, height = app.left, app.top, app.width, app.height

    # Start recording the video
    start_time = time.time()
    frames = []
    while time.time() - start_time < 10:  # Adjust the duration as needed
        frame = ImageGrab.grab(bbox=(left, top, left + width, top + height))
        frames.append(frame)

    # Create the video clip
    video_clip = ImageSequenceClip(frames, fps=30)

    # Save the video
    video_clip.write_videofile(video_path, codec='libx264')

    # Close PowerPoint
    app.close()

# Example usage
ppt_path = "C:\\dir\\first.pptx"
video_path = "output_video.mp4"
convert_ppt_to_video(ppt_path, video_path)
import pyautogui
import time
from PIL import ImageGrab
from moviepy.editor import *

def convert_ppt_to_video(ppt_path, video_path):
    # Open the PowerPoint presentation
    pyautogui.press('win')  # Open the Windows start menu
    time.sleep(1)  # Wait for the start menu to open
    pyautogui.write('PowerPoint')  # Search for PowerPoint
    time.sleep(1)  # Wait for search results to appear
    pyautogui.press('enter')  # Open PowerPoint
    time.sleep(2)  # Wait for PowerPoint to open

    # Open the PPT file
    pyautogui.press('ctrl+o')  # Open the file dialog
    time.sleep(1)  # Wait for the dialog to appear
    pyautogui.write("C:\\dir\\first.pptx")  # Enter the PPT file path
    pyautogui.press('enter')  # Press Enter to open the file
    time.sleep(2)  # Wait for the file to open

    # Set up the video recording
    app = pyautogui.getWindowsWithTitle('PowerPoint')[0]
    left, top, width, height = app.left, app.top, app.width, app.height

    # Start recording the video
    start_time = time.time()
    frames = []
    while time.time() - start_time < 10:  # Adjust the duration as needed
        frame = ImageGrab.grab(bbox=(left, top, left + width, top + height))
        frames.append(frame)

    # Create the video clip
    video_clip = ImageSequenceClip(frames, fps=30)

    # Save the video
    video_clip.write_videofile(video_path, codec='libx264')

    # Close PowerPoint
    app.close()

# Example usage
ppt_path = "C:\\dir\\first.pptx"
video_path = "output_video.mp4"
convert_ppt_to_video(ppt_path, video_path)
