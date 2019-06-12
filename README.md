# XRD-data-analyser

Hi there,

Romain here, just to let you know a bit about this software and what it allows to do.
To begin with, let's talk a little about the background. 
Basically, during my master thesis about nanomechanical devices at EPFL, I had to carry out some characterization measurements thanks to X-ray diffraction to check the composition of the devices I was producing. The result of this measurement is a curve with several peaks giving informations about the orientation of the crystalline structure of the different materials you have in your device (I am not going to go into details here, don't worry). The interesting information are mainly the center and full width half maximum of each peak.
Given that I had quite a lot of measurements to do and that each curve contained around 3 to 7 peaks, I thought it would be a good idea for me (and the other people who needed it) to build a little software that would facilitate the exploitation of these measurement files.
Therefore, to what extent use this soft ?

This soft automatically plots the curve contained in your measurment file.
It automatically detects peaks (amplitude > 20) up to a maximum of 10 different peaks (for now at least) and automatically fit a gaussian onto it. Therefore, based on each gaussian found, it automatically find the center, amplitude and full width half maximum and plot it next to the graph.

It is also possible to open a (.txt) file or an excel file.
It is also possible to apply a linear fit or a moving average to your data.

Please, be mindful that this software has been tested on a non exhaustive list of files and might still contain some glitches. Therefore, I advise to check the results given by the software to avoid any problem.
