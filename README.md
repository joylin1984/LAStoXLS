# LAStoXLS
Files with LAS extension (Log ASCII Standard) stores data for the methods of geophysical research. 
One file stores information for one well.<br>
The file consists of several sections, but the most interesting - ~W [Well information]. This section contains the following information:
<li>START – the initial depth</li>
<li>STOP – the initial depth</li>
<li>STEP – the quantization step</li>
<li>NULL code for missing information (typically -9999 or -999.25)</li>
<li>WELL – the name of the well</li>
<li>DATE – date of logging</li>
<li>UWI – the unique well code</li>
Program parses all LAS files in folder, and stores information about all wells into one XLS file.
