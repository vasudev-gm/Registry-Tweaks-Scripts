# Registry-Tweaks-Scripts

Contains registry tweaks and simple batch scripts

I will basically divide it into 2 directories/sections called cmds and reg_tweaks which I commonly use when creating custom up-to-date windows 10 or 11 ISO using WinToolkit,NTLite and W10UI. I will include direct links for DirectX9, NetFX35 and VC++ runtimes repacks etc.

Thank you abbodi1406, Majorgeeks, WinAero for reg tweaks, repacks and scripts.

- [DirectX Repack][def] Repack made by [@abbodi1406][def2]

- [Netfx35W10][def3] Repack made by [@abbodi1406][def2]

- [Microsoft VC++ redist runtimes][def4] Repack made by [@abbodi1406][def2]

- [NetfxW7, applicable to Win 7 only][def5] Repack made by [@abbodi1406][def2]

- [MajorGeeks Registry Tweaks Pack covering Win 7-11][def6] Made by [Majorgeek][def7]

- [WHDownloader][def8] made by [@abbodi1406][def2] Used for downloading offline msu/cabs files applicable to windows and office.

- [Unlock hidden power options Win 10][def9]
- [Unlock hidden power options Win 7][def10]
- [DirectX Repack by htfx][def11]
- [Automated Script to download and install VC++ Runtimes][def12]

Feel free to suggest any changes.

You can find all regfiles and cmd scripts in this directory `cmds` and `registry\regfiles\Windows(Version)`

Update: I stopped using NTLite because it was causing some side-effects with broken start menu and other issues. I now use WinToolkit, W10UI and unattended file for creating custom ISOs. I have started using VHDXs formatted with ReFS for testing purposes.

To make VHDX files smaller after integration use sdelete from Sysinternals Suite. It is a command line tool that securely deletes files and cleans up free space on NTFS volumes. It can also be used to zero out free space in VHDX files. Afterwards, run retrim on the VHDX file to reclaim the space.

Added scripts to download Brave, DX/VC++ Redist and IrfanView to get latest versions within unattended.xml file

Thanks to abbodi1406, htfx, exurd and ThioJoe for scripts and repacks

[def]: https://forums.mydigitallife.net/threads/repack-directx-end-user-runtime-june-2010.84785/
[def2]: https://github.com/abbodi1406
[def3]: https://github.com/abbodi1406/dotNetFx35W10
[def4]: https://github.com/abbodi1406/vcredist
[def5]: https://github.com/abbodi1406/dotNetFx4xW7
[def6]: https://github.com/MajorGeek/MajorGeeks-Windows-Tweaks
[def7]: https://github.com/MajorGeek
[def8]: https://forums.mydigitallife.net/threads/whdownloader-download.66243/
[def9]: https://gist.github.com/Nt-gm79sp/1f8ea2c2869b988e88b4fbc183731693
[def10]: https://gist.github.com/theultramage/cbdfdbb733d4a5b7d2669a6255b4b94b
[def11]: https://github.com/stdin82/htfx
[def12]: https://github.com/exurd/Windows-Sandbox-Tools/blob/vcredist_aio_script/Installer%20Scripts/Install%20VC%20Redist%20AIO.ps1
