This is the initial release of a controller that utilizes undocumented DirectShow IAMCameraControl properties of the OBSBot Tiny USB web cameras for silky smooth Pan, Tilt and Zoom (PTZ).

Choose your camera based on device path and click "Start OSC Server".

Use something like Chataigne to send OSC commands to control the camera, sample Chataigne file included using a Nintendo Switch JoyCon as sample.noisette.

This code is drived from the DxPropPages sample from DirectShow.NET located here: https://sourceforge.net/projects/directshownet/files/DirectShowSamples/2010-February/

Usage:

OSC Addresses:

/PAN int

/TILT int

/ZOOM int

/EXPOSURE




# Likely OBSBot only! Utilizing undocumented IAMCameraControl values 10, 11 and 13. 
# 10 is pan   speed limit is +-178
# 11 is tilt  speed limit is +-178
# 13 is zoom  speed limit is +-100

# Can be used for presets
/PTZS int int int int 



/FLYXY int int

/FLYZ int