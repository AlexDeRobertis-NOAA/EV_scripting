
from nate 3/18/24
Check out the two attached scripts:

1. EchoviewExport.py -- The method 'export_py_MB2 is where the bulk of the scripting happens, see Lines 1040 - 1207 (following the condition if self.exportType  == 0, which is the normal non-multifrequency export).

2. EvFileMaker.py -- Within the method, 'makeFile', you'll find the scripting file creation, see Lines 477 - 631.

Let me know if you have any questions or want to follow-up on any specific parts.