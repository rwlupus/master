import trackanimation
from trackanimation.animation import AnimationTrack

# Simple example
input_directory = "d:/robi/Turystyka/2018_Tajlandia/trk/"

tha_trk = trackanimation.read_track(input_directory)
tha_trk = tha_trk.time_video_normalize(time=10, framerate=10)
fig = AnimationTrack(df_points=tha_trk, dpi=300, bg_map=True, map_transparency=0.9)
# fig = AnimationTrack(df_points=tha_trk, dpi=300, bg_map=True, map_transparency=0.5)

# print(type(fig))
fig.make_video(output_file='thailand4', framerate=10, linewidth=2.0)
