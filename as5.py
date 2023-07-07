from moviepy.editor import VideoFileClip, ImageSequenceClip, concatenate_videoclips

# Paths to the voiceover files and slide images
voiceover_paths = ["voiceover1.mp3", "voiceover2.mp3", "voiceover3.mp3"]
slide_image_paths = ["slide1.png", "slide2.png", "slide3.png"]

# Load the voiceovers as audio clips
voiceovers = [AudioFileClip(voiceover_path) for voiceover_path in voiceover_paths]

# Load the slide images as image clips
slides = [ImageSequenceClip([slide_image_path], durations=[5]) for slide_image_path in slide_image_paths]

# Concatenate the voiceovers and slides together
final_clips = [concatenate_videoclips([slides[i], voiceovers[i]]) for i in range(len(slides))]

# Combine all the final clips into a single video presentation
final_video = concatenate_videoclips(final_clips)

# Export the final video
final_video.write_videofile("presentation.mp4", codec="libx264", audio_codec="aac")
