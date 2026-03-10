# Ubehage's Bit Manipulator v1.0.2

A small and portable program for flipping or removing a single bit anywhere in a file.  

## Screenshot
![Main Window](./screenshot.png)

## Safety notice
- This program has no warning dialogs, nor does it ask you if you are sure what you are doing.
  It does exactly what you ask it, immediately.
- Remember to take backup of important data.

## Recent fixes
- 1.0.2
  - When copying to a new file, the program would crash if the byte edited was the first or last byte.
  - Known problems:
    - If you remove the first bit in a byte, the result might not be as expected.
	  I am looking into that.
	  
## License
Copyright © Ubehage 2026.  
MIT License. All code is free to use and modify.