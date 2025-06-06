# AttackSurfaceReduction

This repository contains example VBA macros and minimal artifacts used to demonstrate Microsoft Defender Attack Surface Reduction (ASR) scenarios. These macros are intentionally designed to perform actions such as downloading files and launching external processes so that ASR rules can be evaluated against them.

## Warning

Running the macros in this repository may download files from the Internet or launch other programs on your system. Only open and execute the files in a controlled environment (for example, an isolated virtual machine) where you can safely observe the behavior.

## Prerequisites

- Windows operating system
- Microsoft Office that supports VBA macros (Office 2016 or newer is recommended)

## Opening the Macros Safely

1. Set up a test virtual machine running Windows with Microsoft Office installed.
2. Copy the `.vba` files from the `Macros` folder (and `test.vba` in the repository root) to the VM.
3. Open Word or Excel and press **Alt+F11** to launch the VBA editor.
4. From the VBA editor, create a new module and paste the contents of one of the `.vba` files into it.
5. Review the macro code so you understand what it will do. Each macro includes comments describing its behavior.
6. Run the macro using the **Run** button or by pressing **F5**.
7. Observe how Defender or configured ASR rules react to the macro attempting to download files or run external commands.

Use snapshots in your VM so you can revert the environment after testing.
