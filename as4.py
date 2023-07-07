from moviepy.editor import *

def sync_voiceovers_with_slides(voiceover_path, slides_path, output_video_path):
    # Load the voiceover file
    voiceover = AudioFileClip(voiceover_path)

    # Load the slides as an image sequence
    slides = ImageSequenceClip(slides_path, fps=1)

    # Set the duration of each slide to match the duration of the voiceover
    slides = slides.set_duration(voiceover.duration)

    # Set the audio of each slide to the voiceover
    slides = slides.set_audio(voiceover)

    # Write the final video with synced voiceovers and slides
    slides.write_videofile(output_video_path, codec='libx264', audio_codec='aac')

# Example usage
voiceover_path = "voiceover.wav"
slides_path = ["slide1.jpg", "slide2.jpg", "slide3.jpg"]  # List of slide image paths
output_video_path = "presentation_video.mp4"
sync_voiceovers_with_slides(voiceover_path, slides_path, output_video_path)
