This is the initial release of a controller that utilizes undocumented DirectShow IAMCameraControl properties of the OBSBot Tiny USB web cameras for silky smooth Pan, Tilt and Zoom (PTZ).

Choose your camera based on device name and click "Start OSC Server". For multiple cameras run multiple copies of the program and run on different ports.

Use something like Chataigne to send OSC commands to control the camera, sample Chataigne file included using a Nintendo Switch JoyCon as sample.noisette.

This code is drived from the DxPropPages sample from DirectShow.NET located here: https://sourceforge.net/projects/directshownet/files/DirectShowSamples/2010-February/

Usage:

OSC Addresses that should work on all cameras that support these functions. Be sure to click "Dump Settings" to see what ranges you cameras can do. 
/PAN int
/TILT int
/ZOOM int
/EXPOSURE int
/FOCUS int


# The following are likely OBSBot only! Utilizing undocumented IAMCameraControl values 10, 11 and 13. 
# 10 is pan   speed limit is +-178
# 11 is tilt  speed limit is +-178
# 13 is zoom  speed limit is +-100

# Pan, Tilt, Zoom, Speed. Can be used for presets! Tweening is a work in progress but mostly working.
/PTZS int int int int 

# Start moving in the specified direction by speed. +-178 for pan or tilt.
/FLYXY int int

# Zoom in or out by speed, +-100
/FLYZ int