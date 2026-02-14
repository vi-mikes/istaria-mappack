HashVersionWriter.exe needs to be ran beside the latest compiled MapPackSyncTool.exe . This is a console exe!

If MapPackSyncTool.exe is missing then it will exit.

It then looks for version.txt to overwrite. If version.txt does not exist then it will create it first.

It will write version.txt as:

Line 1: Version Number (example: 0.0.2) which is derived from the MapPackSyncTool.exe file.
  Note version number is included in the MapPackSyncTool.rc file.

Line 2: SHA-256 value of MapPackSyncTool.exe.

This file gets uploaded to S3: /istaria-mappack/version.txt
