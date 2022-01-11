# VBA-PowerPoint

VBA Macros for PowerPoint

This repository contains useful macros for PowerPoint. The macros are organized into _packages_. Each subdirectory in this repository contains an independent package. Each package directory also has a `Tests` subdirectory, which contains a .PPTM which has test cases and usage instructions for the package, as well as the package code. 

## Using these macros

### To use an individual package

1. Download the .PPTM file in the package's `Tests` subdirectory.
2. Open it, remembering to enable macro content.
3. Invoke the macro from the Macros dialog available via the View ribbon tab. 

### To use all packages

1. Download the MaxOfficeVBAMacros.PPTM file from the [latest release](https://github.com/MaxOffice/VBA-PowerPoint/releases/latest).
2. Open it, remembering to enable macro content.
3. Invoke any macro from the Macros dialog available via the View ribbon tab.

## To use packages of your choice

1. Create or open a .PPTM file
2. Use `File/Import File...` to import all files found directly under all your chosen package subdirectories. These may include `.bas`, `.cls`, `.frm` and `.frx` files. Do  not import anything in the `Tests` subirectories. 
3. Save the file, and keep it open.
4. Invoke any macro from the Macros dialog available via the View ribbon tab.
