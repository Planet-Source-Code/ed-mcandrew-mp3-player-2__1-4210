This is my second release of  my mp3 player code. This is as far as the project went before
I scrapped it and started over. The code was becoming just too messy to be efficient.

The new project called "FreePlayer Mp3" will be distributed soon, if your interested, check
our website at: http://204.116.190.57

Freeplayer uses a direct call to the quartz.dll (instead of the mediaplayer.ocx), and has
skinning integrated, with full support of multiple playlists (including randomizing,
shuffling, repeat and point looping).  It also includes tray integration with just about
every feature found in high-end mp3 decoders (such as winamp).

Anyhow in order to use Mp3 Player 2:

1. Copy Mp3Player2.ini to C:\
2. Have mediaplayer installed.
3. Copy HIndic32.OCX to Windows\System (or in Nt Winnt\system32)
   (if you encounter errors, you can just install Mabry's Indicator Indc.exe)
   (The Maybry indicator ocx is shareware from http://www.mabry.com)

***Unfortunately, Microsoft's MediaPlayer doesn't use any waveout functions (that I'm aware of
   so, the whole indicator thing is not even used with this player.  However, if your soundcard
   supports waveout functions, you can open Mp3 Player 2, and run something that plays a wav
   file to debug the indicators (this was the primary reason I scratched this project in favor
   of a new)


I do not claim to be a great programmer.  Rather just another person teaching himself VB.
If you feel that the code is too messy or just don't like it. Then simply clean it
or don't use it.

If you do decide you like it, then great have fun, and do whatever you like with it.

Ed McAndrew
