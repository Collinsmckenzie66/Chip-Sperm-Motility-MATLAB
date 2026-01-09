# Chip-Sperm-Motility-MATLAB
contains current MATLAB script that takes spot info from trackmate and outputs excel sheets with categorized sperm tracks and track speeds comparable to CASA systems.  

This script uses the simple moving average smoothing method to calculate the average path of the sperm, something track mate can not calculate. VSL and VCL readings calculated through this script are exactly the same as the ones track mate produces. This script also generates Straightness and Wobble measurements, also unique to this script compared to track mate.

# Information on Track Excel Outputs
TrackID- unique track identifier  

VCL (um/s)- Curvilinear Velocity; Total Distance/Time  

VSL (um/s)- Straight Line Velocity; Displacement/Time  

VAP (um/s)- Average Velocity of the Smoothed Path  

TotalDist (um)- Total Distance the sperm traveled  

Displacement (um)- Sperm Origin minus Sperm destination  

Linearity (%)- VSL/VCL; Measure of the progressiveness of a track. Higher numbers= Correlation with highly progressive sperm. Lower numbers= correlation with hyperactivation.  

Straightness (%)- VSL/VAP; Measure of the directness of a track. Higher numbers = correlation with Hyperactivation  

Wobble (%)- VAP/VCL; Measures erractness or sideways movement of a track. Higher numbers= correlation with defective sperm.  

Hyperactivated (T/F)- Indicates if a sperm is hyperactive or not pased on predetermined parameters  


# Summary Tab Information
Rapid Progressive- % of sperm whos Linearity is greater than 0.5 and VAP is greater than 50  

Medium Progressive- % of sperm whos Linearity is less than 0.5 and VAP is greater than 25 but less than 50  

Non-Progressive- % of sperm whos VAP is greater than 5 but less than 25, and whos VSL is less than 25 or Linearity is less than 0.5  

Hyperactivated- % of sperm whos linearity is less than 0.5 and VCL is greater than 100.  

These categorizations show up both as total counts and percentages of the total population.
