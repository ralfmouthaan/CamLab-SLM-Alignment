# CamLab-SLM-Alignment

VB.NET code to manually control an SLM connected to the PC via HDMI. The SLM is assumed to be set up as a second screen. Mostly intended for coarse-aligning a displayed Fourier hologram with a multimode fibre to subsequently measure a transmission matrix or excite discrete modes in the fibre. Allows Zernike polynomials to be manually adjusted, and both blank and custom holograms to be displayed. This code also provides a good example of how the RPM CamLab VB.NET classes work. 

An attempt is made to speed up the code by using some native Windows functionality and parallel processing.

Originally built to work with a Jasper SLM, but should work with other nematic liquid crystal SLMs just by changing the expected SLM resolution.
